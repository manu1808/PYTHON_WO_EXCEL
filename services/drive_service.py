"""
Servicio de integración con Google Drive API v3 mediante OAuth 2.0.
"""

import sys
import os
import json
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request

from config import GOOGLE_TOKEN_FILE, DRIVE_SCOPES, DRIVE_FOLDER_ID

class DriveService:
    @staticmethod
    def get_google_credentials() -> Credentials:
        """
        Obtiene y refresca las credenciales de OAuth2 desde el token.json.
        """
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            root_dir = os.path.abspath(os.path.join(base_dir, ".."))
            if os.path.isfile(os.path.join(root_dir, GOOGLE_TOKEN_FILE)):
                base_dir = root_dir

        token_path = os.path.join(base_dir, GOOGLE_TOKEN_FILE)

        if not os.path.isfile(token_path):
            raise FileNotFoundError(
                "No se encontró token.json.\n\n"
                f"Archivo esperado:\n{token_path}"
            )

        with open(token_path, "r", encoding="utf-8") as f:
            token_data = json.load(f)

        credentials = Credentials(
            token=None,
            refresh_token=token_data["refresh_token"],
            token_uri=token_data.get("token_uri", "https://oauth2.googleapis.com/token"),
            client_id=token_data["client_id"],
            client_secret=token_data["client_secret"],
            scopes=DRIVE_SCOPES
        )

        credentials.refresh(Request())
        return credentials

    @classmethod
    def upload_file(cls, file_path: str, progress_cb=None) -> dict:
        """
        Sube un archivo de hoja de cálculo a la carpeta configurada en Google Drive.
        """
        def progress(msg):
            if progress_cb:
                progress_cb(msg)

        progress("Autenticando con Google Drive…")
        credentials = cls.get_google_credentials()

        service = build("drive", "v3", credentials=credentials)
        file_name = os.path.basename(file_path)

        progress(f"Subiendo {file_name} a Google Drive…")

        metadata = {
            "name": file_name,
            "parents": [DRIVE_FOLDER_ID]
        }

        media = MediaFileUpload(
            file_path,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=True
        )

        uploaded_file = (
            service.files()
            .create(
                body=metadata,
                media_body=media,
                fields="id,name,webViewLink"
            )
            .execute()
        )

        progress("Subida completada correctamente.")
        return uploaded_file