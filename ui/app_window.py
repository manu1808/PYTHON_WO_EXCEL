"""
Ventana principal de Tkinter con ejecución asíncrona (Threading) para no congelar la GUI.
"""

import os
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core.excel_transformer import ExcelTransformer
from core.overdue_reporter import OverdueReporter
from services.drive_service import DriveService
from utils.file_utils import temporary_output_path, get_overdue_temp_path, cleanup_files

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("WO Transformer — CNOC")
        self.resizable(False, False)
        self.configure(bg="#f0f4f8")

        # HEADER
        tk.Label(
            self,
            text="WO Export  →  WO CNOC",
            font=("Segoe UI", 13, "bold"),
            bg="#1b4332",
            fg="white",
            padx=14,
            pady=11
        ).pack(fill="x")

        frame = tk.Frame(self, bg="#f0f4f8", padx=24, pady=20)
        frame.pack(fill="both", expand=True)

        # ARCHIVO RAW
        tk.Label(
            frame,
            text="Archivo RAW (WoExport):",
            bg="#f0f4f8",
            font=("Segoe UI", 10)
        ).grid(row=0, column=0, sticky="w", pady=(0, 3))

        self.raw_var = tk.StringVar()

        tk.Entry(
            frame,
            textvariable=self.raw_var,
            width=54,
            font=("Segoe UI", 9)
        ).grid(row=1, column=0, sticky="ew", padx=(0, 8))

        tk.Button(
            frame,
            text="Buscar…",
            command=self.pick_raw,
            bg="#52b788",
            fg="white",
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=10
        ).grid(row=1, column=1)

        # DESTINO
        tk.Label(
            frame,
            text="Destino: Google Drive → TECHNICAL_LAL → WO_CNOC → 08 - AGOSTO",
            bg="#f0f4f8",
            fg="#555555",
            font=("Segoe UI", 9)
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(14, 0))

        # ESTADO
        self.status_var = tk.StringVar(value="Selecciona el archivo RAW para comenzar.")

        tk.Label(
            frame,
            textvariable=self.status_var,
            bg="#f0f4f8",
            font=("Segoe UI", 9),
            fg="#444444"
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(14, 4))

        # PROGRESS BAR
        self.progress = ttk.Progressbar(frame, mode="indeterminate", length=430)
        self.progress.grid(row=4, column=0, columnspan=2, sticky="ew")

        # BOTÓN TRANSFORMAR
        self.transform_button = tk.Button(
            frame,
            text="▶  Transformar y subir",
            command=self.run,
            bg="#1b4332",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            padx=24,
            pady=9
        )
        self.transform_button.grid(row=5, column=0, columnspan=2, pady=(22, 0))

        frame.columnconfigure(0, weight=1)

    def pick_raw(self):
        path = filedialog.askopenfilename(
            title="Selecciona el archivo RAW",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.raw_var.set(path)
            self.set_status("Archivo RAW seleccionado.")

    def set_status(self, msg: str):
        self.status_var.set(msg)
        self.update_idletasks()

    def run(self):
        raw = self.raw_var.get().strip()

        if not raw:
            messagebox.showwarning("Falta archivo", "Selecciona el archivo RAW primero.")
            return

        if not os.path.isfile(raw):
            messagebox.showerror("No encontrado", f"No existe:\n{raw}")
            return

        clean_out = temporary_output_path(raw)
        overdue_out = get_overdue_temp_path(clean_out)

        # Limpieza de archivos temporales previos
        delete_errors = cleanup_files(clean_out, overdue_out)
        if delete_errors:
            messagebox.showerror(
                "Archivo temporal ocupado",
                "No se puede utilizar el archivo temporal. Verifica que no esté abierto en Excel."
            )
            return

        self.transform_button.config(state="disabled")
        self.progress.start(12)
        self.set_status("Procesando…")

        # Hilo secundario para mantener fluida la interfaz gráfica
        threading.Thread(
            target=self._run_process_thread,
            args=(raw, clean_out, overdue_out),
            daemon=True
        ).start()

    def _run_process_thread(self, raw: str, clean_out: str, overdue_out: str):
        try:
            # 1. Crear archivo limpio
            colored = ExcelTransformer.transform(
                raw, clean_out, progress_cb=self.set_status
            )

            # 2. Crear reporte de vencidas
            overdue_count = OverdueReporter.create_overdue_report(
                clean_out, overdue_out, progress_cb=self.set_status
            )

            # 3. Subir archivo limpio
            self.set_status("Subiendo archivo limpio a Google Drive…")
            uploaded_clean = DriveService.upload_file(
                clean_out, progress_cb=self.set_status
            )

            # 4. Subir reporte de vencidas
            self.set_status("Subiendo reporte de vencidas a Google Drive…")
            uploaded_overdue = DriveService.upload_file(
                overdue_out, progress_cb=self.set_status
            )

            # 5. Eliminar temporales
            self.set_status("Eliminando archivos temporales…")
            delete_errors = cleanup_files(clean_out, overdue_out)

            # 6. Notificar éxito en el hilo de la GUI
            self.after(0, self._on_success, uploaded_clean, uploaded_overdue, overdue_count, colored, delete_errors)

        except Exception as exc:
            cleanup_files(clean_out, overdue_out)
            err_msg = str(exc) + "\n\n" + traceback.format_exc()[-1200:]
            self.after(0, self._on_error, err_msg)

    def _on_success(self, uploaded_clean, uploaded_overdue, overdue_count, colored, delete_errors):
        self.progress.stop()
        self.set_status("✓ Archivos enviados correctamente a Google Drive.")
        self.transform_button.config(state="normal")

        clean_name = uploaded_clean.get("name", "WO_CNOC")
        overdue_name = uploaded_overdue.get("name", "WO_CNOC_VENCIDAS")

        local_message = (
            "\n\nAdvertencia: uno o más archivos temporales no pudieron eliminarse del PC."
            if delete_errors else
            "\n\nLos archivos temporales locales fueron eliminados."
        )

        messagebox.showinfo(
            "¡Proceso completado!",
            "Los archivos fueron procesados y enviados correctamente a Google Drive.\n\n"
            f"ARCHIVO LIMPIO:\n{clean_name}\n\n"
            f"REPORTE DE VENCIDAS:\n{overdue_name}\n\n"
            f"WOs vencidas: {overdue_count}\n\n"
            f"Celdas negativas resaltadas: {colored}"
            + local_message
        )

    def _on_error(self, err_msg: str):
        self.progress.stop()
        self.set_status("Error durante el proceso.")
        self.transform_button.config(state="normal")
        messagebox.showerror("Error", err_msg)