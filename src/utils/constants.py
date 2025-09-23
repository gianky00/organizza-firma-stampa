import os

# --- PATHS ---
# Determine the base path for the application, works for both script and frozen exe
if getattr(os.sys, 'frozen', False):
    APPLICATION_PATH = os.path.dirname(os.sys.executable)
else:
    APPLICATION_PATH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# --- DEFAULT FOLDER AND FILE NAMES ---
FIRMA_EXCEL_INPUT_DIR = "FILE EXCEL DA FIRMARE"
FIRMA_PDF_OUTPUT_DIR = "PDF"
FIRMA_IMAGE_NAME = "TIMBRO.png"
APP_ICON_NAME = "app_icon.ico"

ORGANIZZA_SOURCE_DIR = "SCHEDE DA ORGANIZZARE"
ORGANIZZA_DEST_DIR = "SCHEDE ORGANIZZATE"

RINOMINA_DEFAULT_DIR = "SCHEDE SENZA DATA"

CONFIG_FILE_NAME = "config_programma.json"

# --- NETWORK AND EXTERNAL PATHS ---
# These are unlikely to change but are kept here for centralization
CANONI_GIORNALIERA_BASE_DIR = r"\\192.168.11.251\Database_Tecnico_SMI\Giornaliere"
CANONI_CONSUNTIVI_BASE_DIR = r"\\192.168.11.251\Database_Tecnico_SMI\Contabilita' strumentale"
CANONI_WORD_DEFAULT_PATH = r"C:\Users\Coemi\Desktop\foglioNuovoCanone.docx"
DEFAULT_GHOSTSCRIPT_PATH = r"C:\Program Files\gs\gs10.05.0\bin\gswin64c.exe"

# --- APPLICATION DATA ---
MESI_GIORNALIERA_MAP = {
    "Gennaio": "01", "Febbraio": "02", "Marzo": "03", "Aprile": "04",
    "Maggio": "05", "Giugno": "06", "Luglio": "07", "Agosto": "08",
    "Settembre": "09", "Ottobre": "10", "Novembre": "11", "Dicembre": "12"
}
NOMI_MESI_ITALIANI = list(MESI_GIORNALIERA_MAP.keys())

DEFAULT_MACRO_NAME = "Modulo42.StampaFogli"
