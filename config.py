"""
Módulo de Configuración Global
Contiene las constantes, rutas, esquemas de columnas y estilos visuales reutilizables.
"""

from openpyxl.styles import PatternFill

# Columnas que se deben conservar en el orden especificado
COLUMNS_TO_KEEP = [
    'WO code',
    'WO Name',
    'WO Status',
    'Create Persion',
    'Create Time',
    'CD',
    'FT',
    'Priority',
    'Start Time',
    'End Time',
    'Remain time (H)'
]

# Ancho personalizado por columna
COL_WIDTHS = {
    'WO code':          38.71,
    'WO Name':          78.14,
    'WO Status':        14.86,
    'Create Persion':   19.57,
    'Create Time':      19.0,
    'CD':               32.29,
    'FT':               20.57,
    'Priority':         7.57,
    'Start Time':       19.0,
    'End Time':         19.0,
    'Remain time (H)':  14.86,
}

# Alturas de filas predeterminadas
HEADER_ROW_HEIGHT = 24.75
DATA_ROW_HEIGHT   = 12.0

# Estilos de resaltado para valores negativos
NEG_FILL = PatternFill("solid", fgColor="FFFF00")  # Amarillo
NEG_COLOR = "FF0000"                               # Texto Rojo

# Configuración de Google Drive
DRIVE_FOLDER_ID = "14W9fUS6cFmKOSR09stVS7c2KyvLHeSa9"
GOOGLE_TOKEN_FILE = "token.json"
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]