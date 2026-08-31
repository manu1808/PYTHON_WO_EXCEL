"""
Utilidades para manejo de archivos temporales y operaciones de sistema.
"""

import os
import datetime
import tempfile

def temporary_output_path(raw_path: str) -> str:
    """
    Genera la ruta absoluta de un archivo temporal con sufijo de fecha actual.
    """
    today = datetime.date.today().strftime("%d%m%Y")
    file_name = f"WO_CNOC_{today}.xlsx"
    temp_dir = tempfile.gettempdir()
    return os.path.join(temp_dir, file_name)

def get_overdue_temp_path(clean_out_path: str) -> str:
    """
    Genera la ruta absoluta temporal para el reporte de vencidas.
    """
    temp_dir = os.path.dirname(clean_out_path)
    today = datetime.date.today().strftime("%d%m%Y")
    return os.path.join(temp_dir, f"VENCIDAS_CNOC_{today}.xlsx")

def cleanup_files(*file_paths: str) -> list:
    """
    Intenta eliminar una lista de archivos y retorna los errores que ocurran.
    """
    errors = []
    for path in file_paths:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception as e:
                errors.append(f"{path}: {e}")
    return errors