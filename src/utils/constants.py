import os

# --- PATHS ---
# Determine the base path for the application, works for both script and frozen exe
if getattr(os.sys, 'frozen', False):
    APPLICATION_PATH = os.path.dirname(os.sys.executable)
else:
    # Go up two levels from this file's location (src/utils/constants.py) to get to the project root.
    APPLICATION_PATH = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

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

# --- Email Feature Constants ---
TCL_CONTACTS = {
    "Domenico Passanisi": "dpassanisi@isab.com",
    "Francesco Naselli": "fnaselli@isab.com",
    "Ferdinando Caldarella": "fcaldarella@isab.com",
    "Manuel Prezzavento": "mprezzavento@isab.com",
    "Ivan Messina": "imessina@isab.com"
}

# Placeholders: {name} for the first name, {file_list} for the list of files.
EMAIL_BODY_INFORMAL = "Ciao {name},\n\ndi seguito elenco delle schede in allegato firmate:\n\n{file_list}\n\nSaluti,"
EMAIL_BODY_FORMAL = "Buongiorno {name},\n\nin allegato la documentazione richiesta.\n\nElenco file:\n{file_list}\n\nCordiali Saluti,"

EMAIL_BODY_GENERIC_INFORMAL = "Ciao,\n\ndi seguito elenco delle schede in allegato firmate:\n\n{file_list}\n\nSaluti,"
EMAIL_BODY_GENERIC_FORMAL = "Buongiorno,\n\nin allegato la documentazione richiesta.\n\nElenco file:\n{file_list}\n\nCordiali Saluti,"
