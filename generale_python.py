import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
import os
import sys
import subprocess
import shutil
import re
import threading
import json
from datetime import datetime

# Import specifici per le varie funzionalità
import win32com.client
import pythoncom
import openpyxl  # Mantenuto per riferimenti futuri
import xlrd      # Mantenuto per riferimenti futuri
import traceback
import win32print # NUOVO IMPORT per elencare le stampanti


# --- CONFIGURAZIONE GLOBALE ---
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

# Nomi delle cartelle e file di default
FIRMA_EXCEL_INPUT_DIR = "FILE EXCEL DA FIRMARE"
FIRMA_PDF_OUTPUT_DIR = "PDF"
FIRMA_IMAGE_NAME = "TIMBRO.png"
APP_ICON_NAME = "app_icon.ico"

ORGANIZZA_SOURCE_DIR = "SCHEDE DA ORGANIZZARE"
ORGANIZZA_DEST_DIR = "SCHEDE ORGANIZZATE"

RINOMINA_DEFAULT_DIR = "SCHEDE SENZA DATA"

CANONI_GIORNALIERA_BASE_DIR = r"\\192.168.11.251\Database_Tecnico_SMI\Giornaliere"
# MODIFICATO: Il percorso base dei consuntivi ora è più generico
CANONI_CONSUNTIVI_BASE_DIR = r"\\192.168.11.251\Database_Tecnico_SMI\Contabilita' strumentale"
CANONI_WORD_DEFAULT_PATH = r"C:\Users\Coemi\Desktop\foglioNuovoCanone.docx"

CONFIG_FILE_NAME = "config_programma.json"


class MainApplication(tk.Tk):
    """
    Applicazione unificata con interfaccia a schede (tab) per
    gestire firme, organizzazione, ridenominazione e stampa di file.
    """
    def __init__(self):
        super().__init__()
        self.title("Gestione Documenti Ufficio")
        self.center_window(950, 850)
        self.resizable(True, True)

        self._setup_style_and_icon()
        self._load_processing_data()
        self._setup_initial_folders()

        self.mesi_giornaliera_map = {
            "Gennaio": "01", "Febbraio": "02", "Marzo": "03", "Aprile": "04",
            "Maggio": "05", "Giugno": "06", "Luglio": "07", "Agosto": "08",
            "Settembre": "09", "Ottobre": "10", "Novembre": "11", "Dicembre": "12"
        }
        self.nomi_mesi_italiani = list(self.mesi_giornaliera_map.keys())
        
        current_year = datetime.now().year
        self.anni_giornaliera = [str(y) for y in range(current_year - 5, current_year + 6)]

        self._initialize_stringvars()
        self._create_widgets()
        self._load_config()
        
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # RIGA AGGIUNTA: Popola le stampanti dopo che la GUI è pronta
        self.after(100, self._populate_printers)

    def _initialize_stringvars(self):
        self.firma_excel_dir = tk.StringVar(value=os.path.join(application_path, FIRMA_EXCEL_INPUT_DIR))
        self.firma_image_path = tk.StringVar(value=os.path.join(application_path, FIRMA_IMAGE_NAME))
        self.firma_pdf_dir = tk.StringVar(value=os.path.join(application_path, FIRMA_PDF_OUTPUT_DIR))
        self.firma_ghostscript_path = tk.StringVar(value=r"C:\Program Files\gs\gs10.05.0\bin\gswin64c.exe")
        self.firma_processing_mode = tk.StringVar(value="schede")
        self.rinomina_path = tk.StringVar(value=os.path.join(application_path, RINOMINA_DEFAULT_DIR))
        self.organizza_source_dir = tk.StringVar(value=os.path.join(application_path, ORGANIZZA_SOURCE_DIR))
        self.organizza_dest_dir = tk.StringVar(value=os.path.join(application_path, ORGANIZZA_DEST_DIR))
        
        # --- MODIFICHE PER CANONI ---
        self.canoni_giornaliera_path = tk.StringVar()
        self.canoni_selected_month = tk.StringVar()
        self.canoni_selected_year = tk.StringVar()

        # NUOVE VAR: Per i numeri dei consuntivi inseriti dall'utente
        self.canoni_messina_num = tk.StringVar()
        self.canoni_naselli_num = tk.StringVar()
        self.canoni_caldarella_num = tk.StringVar()
        
        # VECCHIE VAR: Ora usate internamente per memorizzare i percorsi costruiti
        self.canoni_cons1_path = tk.StringVar(value="")
        self.canoni_cons2_path = tk.StringVar(value="")
        self.canoni_cons3_path = tk.StringVar(value="")

        # NUOVA VAR: Per la stampante selezionata
        self.selected_printer = tk.StringVar()

        self.canoni_word_path = tk.StringVar(value=CANONI_WORD_DEFAULT_PATH)
        self.canoni_macro_name = tk.StringVar(value="Modulo42.StampaFogli")
    
    def center_window(self, width, height):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def _setup_style_and_icon(self):
        try:
            icon_path = os.path.join(application_path, APP_ICON_NAME)
            if os.path.exists(icon_path): self.iconbitmap(icon_path)
        except tk.TclError:
            print(f"Attenzione: icona '{APP_ICON_NAME}' non trovata o non valida.")
        self.font_main = ("Segoe UI", 10)
        self.font_bold = ("Segoe UI", 11, "bold")
        style = ttk.Style(self)
        style.theme_use('vista')
        style.configure('.', font=self.font_main) 
        style.configure('TLabel', font=self.font_main)
        style.configure('TLabelframe.Label', font=self.font_bold)
        style.configure('info.TLabel', foreground='#333333')
        style.configure('TButton', padding=6, foreground='black')
        style.configure('primary.TButton', background='#0078D4')
        style.map('primary.TButton', background=[('active', '#005a9e')])
        style.configure('danger.TButton')
        
    def _load_processing_data(self):
        self.firma_processing_data = {
            "schedacontrolloSTRUMENTIANALOGICI": {"PrintArea": "A2:N55", "FirmaCella": "G54"},
            "schedacontrolloSTRUMENTIDIGITALI": {"PrintArea": "A2:N50", "FirmaCella": "G49"},
            "SchedacontrolloREPORTMANUTENZIONECORRETTIVA": {"PrintArea": "A2:N55", "FirmaCella": "G55"},
            "SCHEDAMANUTENZIONE": {"PrintArea": "A1:FV106", "FirmaCella": "BZ105"},
        }
        self.stampa_processing_data = {
            "schedacontrolloSTRUMENTIANALOGICI": {"PrintArea": "A2:N55"},
            "schedacontrolloSTRUMENTIDIGITALI": {"PrintArea": "A2:N50"},
            "SchedacontrolloREPORTMANUTENZIONECORRETTIVA": {"PrintArea": "A2:N55"},
            "SCHEDAMANUTENZIONE": {"PrintArea": "A1:FV106"}
        }

    def _setup_initial_folders(self):
        folders_to_create = [
            os.path.join(application_path, FIRMA_EXCEL_INPUT_DIR),
            os.path.join(application_path, FIRMA_PDF_OUTPUT_DIR),
            os.path.join(application_path, ORGANIZZA_SOURCE_DIR),
            os.path.join(application_path, ORGANIZZA_DEST_DIR),
            os.path.join(application_path, RINOMINA_DEFAULT_DIR)
        ]
        try:
            for folder in folders_to_create:
                os.makedirs(folder, exist_ok=True)
        except Exception as e:
            print(f"ERRORE CRITICO: Impossibile creare le cartelle di lavoro: {e}")
            self.after(100, self.destroy)

    def _create_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(expand=True, fill='both', padx=10, pady=10)

        self.firma_tab = ttk.Frame(notebook)
        self.rinomina_tab = ttk.Frame(notebook)
        self.organizza_tab = ttk.Frame(notebook)
        self.stampa_canoni_tab = ttk.Frame(notebook)

        notebook.add(self.firma_tab, text=' Apponi Firma ')
        notebook.add(self.rinomina_tab, text=' Aggiungi Data Schede ')
        notebook.add(self.organizza_tab, text=' Organizza e Stampa Schede ')
        notebook.add(self.stampa_canoni_tab, text=' Stampa Canoni Mensili ')

        self._create_firma_tab(self.firma_tab)
        self._create_rinomina_tab(self.rinomina_tab)
        self._create_organizza_tab(self.organizza_tab)
        self._create_stampa_canoni_tab(self.stampa_canoni_tab)

    # --- SCHEDA "APPONI FIRMA" ---
    def _create_firma_tab(self, tab):
        main_frame = ttk.Frame(tab, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        desc_text = "Questa sezione automatizza il processo di firma dei documenti. Prende i file Excel dalla cartella 'FILE EXCEL DA FIRMARE', applica la firma 'TIMBRO.png' in base al tipo di documento selezionato, li converte in PDF nella cartella 'PDF' e infine li comprime."
        desc_label = ttk.Label(main_frame, text=desc_text, wraplength=850, justify=tk.LEFT, style='info.TLabel')
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor='w')

        paths_frame = ttk.LabelFrame(main_frame, text="1. Percorsi (Firma)", padding="10")
        paths_frame.pack(fill=tk.X, pady=(0, 5))
        self._create_path_entry(paths_frame, "Cartella Excel:", self.firma_excel_dir, 0, readonly=True)
        self._create_path_entry(paths_frame, "Cartella PDF:", self.firma_pdf_dir, 1, readonly=True)
        self._create_path_entry(paths_frame, "Immagine Firma:", self.firma_image_path, 2, readonly=True)
        self._create_path_entry(paths_frame, "Ghostscript:", self.firma_ghostscript_path, 3, readonly=False, browse_command=lambda: self._select_file_dialog(self.firma_ghostscript_path, "Seleziona eseguibile Ghostscript", [("Executable", "*.exe")]))

        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)
        mode_frame = ttk.LabelFrame(main_frame, text="2. Tipo di Documento da Firmare", padding="10")
        mode_frame.pack(fill=tk.X, pady=5)
        ttk.Radiobutton(mode_frame, text="Schede (Controllo, Manutenzione, etc.)", variable=self.firma_processing_mode, value="schede").pack(anchor=tk.W, padx=5, pady=2)
        ttk.Radiobutton(mode_frame, text="Preventivi (Basato su foglio 'Consuntivo')", variable=self.firma_processing_mode, value="preventivi").pack(anchor=tk.W, padx=5, pady=2)
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)
        actions_frame = ttk.LabelFrame(main_frame, text="3. Azioni Firma", padding="10")
        actions_frame.pack(fill=tk.X, pady=5)
        self.firma_run_button = ttk.Button(actions_frame, text="▶  AVVIA PROCESSO FIRMA COMPLETO", style='primary.TButton', command=self._start_firma_process)
        self.firma_run_button.pack(fill=tk.X, ipady=8, pady=5)
        self.firma_clean_pdf_button = ttk.Button(actions_frame, text="🗑️  PULISCI CARTELLA PDF", style='danger.TButton', command=lambda: self._start_folder_cleanup(self.firma_pdf_dir.get(), FIRMA_PDF_OUTPUT_DIR, self.firma_log_area, self._toggle_firma_buttons))
        self.firma_clean_pdf_button.pack(fill=tk.X, ipady=4, pady=(8, 2))
        self.firma_clean_excel_button = ttk.Button(actions_frame, text="🗑️  PULISCI CARTELLA EXCEL (FIRMA)", style='danger.TButton', command=lambda: self._start_folder_cleanup(self.firma_excel_dir.get(), FIRMA_EXCEL_INPUT_DIR, self.firma_log_area, self._toggle_firma_buttons))
        self.firma_clean_excel_button.pack(fill=tk.X, ipady=4, pady=(2, 0))
        log_frame = ttk.LabelFrame(main_frame, text="Log Esecuzione (Firma)", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        self.firma_log_area = self._create_log_widget(log_frame)

    # --- SCHEDA "AGGIUNGI DATA SCHEDE" ---
    def _create_rinomina_tab(self, tab):
        main_frame = ttk.Frame(tab, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        desc_text = "Questa funzione analizza tutti i file Excel in una cartella, cerca la data di emissione al loro interno e rinomina i file aggiungendo la data nel formato (GG-MM-AAAA). Se un file è protetto da password, proverà ad usare la password 'coemi'."
        desc_label = ttk.Label(main_frame, text=desc_text, wraplength=850, justify=tk.LEFT, style='info.TLabel')
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor='w')

        paths_frame = ttk.LabelFrame(main_frame, text="1. Percorso di Lavoro", padding="10")
        paths_frame.pack(fill=tk.X, pady=(0, 5))
        self._create_path_entry(paths_frame, "Cartella da analizzare:", self.rinomina_path, 0, readonly=False, browse_command=lambda: self._select_folder_dialog(self.rinomina_path, "Seleziona cartella con le schede da rinominare"))
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)
        
        actions_frame = ttk.LabelFrame(main_frame, text="2. Azioni", padding="10")
        actions_frame.pack(fill=tk.X, pady=5)
        self.rinomina_run_button = ttk.Button(actions_frame, text="▶  AVVIA PROCESSO DI RINOMINA", style='primary.TButton', command=self._start_rinomina_process)
        self.rinomina_run_button.pack(fill=tk.X, ipady=8, pady=5)

        log_frame = ttk.LabelFrame(main_frame, text="Log Esecuzione (Aggiungi Data)", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        self.rinomina_log_area = self._create_log_widget(log_frame)

    def _start_rinomina_process(self):
        root_path = self.rinomina_path.get()
        if not os.path.isdir(root_path):
            self._log(self.rinomina_log_area, f"ERRORE: Percorso non valido o inesistente: '{root_path}'", "ERROR")
            return
        
        self._clear_log(self.rinomina_log_area)
        self._toggle_rinomina_buttons('disabled')
        threading.Thread(target=self._run_rinomina_thread, args=(root_path,), daemon=True).start()

    def _run_rinomina_thread(self, root_path):
        pythoncom.CoInitialize()
        try:
            self._rename_excel_files_in_place(root_path, self.rinomina_log_area)
        except Exception as e:
            self._log(self.rinomina_log_area, f"ERRORE CRITICO INASPETTATO: {e}", "ERROR")
            self._log(self.rinomina_log_area, traceback.format_exc(), "ERROR")
        finally:
            self._toggle_rinomina_buttons('normal')
            pythoncom.CoUninitialize()
            
    def _toggle_rinomina_buttons(self, state):
        self.rinomina_run_button.config(state=state)
        
    def _rename_excel_files_in_place(self, root_path, log_widget):
        self._log(log_widget, "[FASE 1/2] Raccolta file Excel...", "HEADER")
        excel_files = []
        for root, _, filenames in os.walk(root_path):
            for filename in filenames:
                if filename.lower().endswith(('.xlsx', '.xlsm', '.xls')) and not filename.startswith('~'):
                    excel_files.append(os.path.join(root, filename))

        if not excel_files:
            self._log(log_widget, "Nessun file Excel trovato.", "WARNING"); return
        
        self._log(log_widget, f"Trovati {len(excel_files)} file Excel. Inizio analisi.", "INFO")
        self._log(log_widget, "[FASE 2/2] Analisi e ridenominazione...", "HEADER")

        DATE_IN_FILENAME_REGEX = re.compile(r'\s*\(\d{2}-\d{2}-\d{4}\)')
        corrected, already_ok, no_date_found, errors = 0, 0, 0, 0
        
        excel_app = None
        try:
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.DisplayAlerts = False

            for file_path in excel_files:
                self._log(log_widget, f"Analisi: {os.path.basename(file_path)}...")
                wb = None
                try:
                    # Tentativo di apertura con e senza password
                    try:
                        wb = excel_app.Workbooks.Open(file_path, ReadOnly=True)
                    except Exception:
                        self._log(log_widget, f"  -> File protetto. Tentativo con password 'coemi'...", "WARNING")
                        wb = excel_app.Workbooks.Open(file_path, ReadOnly=True, Password="coemi")

                    ws = wb.Worksheets(1)
                    
                    # Estrazione valori celle
                    all_cell_refs = ['F3', 'G3', 'C54', 'T6', 'AY3', 'B95', 'N1', 'AK2', 'Q3', 'S3', 'E2', 'F2', 'T2', 'T5', 'F56', 'F44', 'F49', 'B99', 'L46', 'B46', 'B108', 'B45', 'B50', 'B105', 'L52']
                    cell_values = {ref: ws.Range(ref).Value for ref in all_cell_refs}
                    
                    # Logica per trovare la data
                    model_n1 = self._normalize_model_string_rename(cell_values.get('N1'))
                    model_f3 = self._normalize_model_string_rename(cell_values.get('F3'))
                    model_g3 = self._normalize_model_string_rename(cell_values.get('G3'))
                    model_ay3 = self._normalize_model_string_rename(cell_values.get('AY3'))
                    model_q3 = self._normalize_model_string_rename(cell_values.get('Q3'))
                    model_t6 = self._normalize_model_string_rename(cell_values.get('T6'))
                    model_f2 = self._normalize_model_string_rename(cell_values.get('F2'))
                    model_s3 = self._normalize_model_string_rename(cell_values.get('S3'))
                    model_e2 = self._normalize_model_string_rename(cell_values.get('E2'))
                    combined_e2_t2_t5 = model_e2 or self._normalize_model_string_rename(cell_values.get('T2')) or self._normalize_model_string_rename(cell_values.get('T5'))

                    date_candidates = []
                    if model_n1 == "schedatecnicaverificadiscocalibro": date_candidates = ["AK2"]
                    elif model_f3 == "schedavalvole" or model_g3 == "schedavalvole": date_candidates = ["C54"]
                    elif "valvole" in model_ay3: date_candidates = ["B95"]
                    elif model_q3 == "schedataraturastrumentidigitali": date_candidates = ["B50"]
                    elif model_t6 == "valvolediregolazione": date_candidates = ["B108"]
                    elif model_f2 == "schedacontrollovalvole": date_candidates = ["F56"]
                    elif model_f2 == "schedacontrollostrumentidigitali": date_candidates = ["F44"]
                    elif model_f2 == "schedacontrollostrumentianalogici": date_candidates = ["F49"]
                    elif model_f2 == "schedacontrollostrumenti": date_candidates = ["F49", "F44"]
                    elif model_s3 == "schedataraturastrumentodiprocesso": date_candidates = ["B99"]
                    elif combined_e2_t2_t5 == "schedacontrollovalvole": date_candidates = ["L46", "B46", "B108"]
                    elif combined_e2_t2_t5 == "schedacontrollostrumentidigitali": date_candidates = ["B45"]
                    elif model_e2 == "schedacontrollostrumentianalogici": date_candidates = ["L52", "B45", "B50", "B108", "B99", "B105"]
                    elif model_e2 == "schedacontrolloreportmanutenzionecorrettiva": date_candidates = ["B50"]
                    else: date_candidates = ["B45", "B50", "B108", "B99", "B105", "L52", "C54"]

                    emission_date = None
                    for cell_ref in date_candidates:
                        status, date_found = self._extract_date_from_val_rename(cell_values.get(cell_ref))
                        if status == 'VALID':
                            emission_date = date_found
                            break
                    
                    if emission_date:
                        original_dir, original_filename = os.path.split(file_path)
                        base_name, ext = os.path.splitext(original_filename)
                        cleaned_base_name = DATE_IN_FILENAME_REGEX.sub('', base_name).strip()
                        cleaned_base_name = self._clean_windows_duplicate_marker_rename(cleaned_base_name)
                        new_filename = f"{cleaned_base_name} ({emission_date.strftime('%d-%m-%Y')}){ext}"
                        
                        wb.Close(SaveChanges=False) # Chiudi prima di rinominare
                        wb = None

                        if new_filename.lower() != original_filename.lower():
                            new_filepath = os.path.join(original_dir, new_filename)
                            final_path = self._get_unique_filepath_rename(new_filepath)
                            os.rename(file_path, final_path)
                            self._log(log_widget, f"  -> RINOMINATO in: {new_filename}", "SUCCESS")
                            corrected += 1
                        else:
                            self._log(log_widget, "  -> Già corretto.", "INFO")
                            already_ok += 1
                    else:
                        self._log(log_widget, "  -> Data non trovata. File non modificato.", "WARNING")
                        no_date_found += 1

                except Exception as e:
                    self._log(log_widget, f"--- ERRORE FILE: {os.path.basename(file_path)} ---", "ERROR")
                    self._log(log_widget, f"Tipo errore: {type(e).__name__} - Messaggio: {e}", "ERROR")
                    if "password" not in str(e).lower():
                         self._log(log_widget, traceback.format_exc(), "ERROR")
                    errors += 1
                finally:
                    if wb:
                        wb.Close(SaveChanges=False)

        finally:
            if excel_app:
                excel_app.Quit()
        
        self._log(log_widget, "\n--- RIEPILOGO PROCESSO RINOMINA ---", "HEADER")
        self._log(log_widget, f"File rinominati o corretti: {corrected}", "SUCCESS")
        self._log(log_widget, f"File già corretti: {already_ok}", "INFO")
        self._log(log_widget, f"File con data non trovata: {no_date_found}", "WARNING")
        self._log(log_widget, f"File con errori di lettura: {errors}", "ERROR")
        self._log(log_widget, "--- COMPLETATO ---", "HEADER")

    def _get_unique_filepath_rename(self, filepath: str) -> str:
        if not os.path.exists(filepath): return filepath
        base, ext = os.path.splitext(filepath)
        counter = 1
        while True:
            new_path = f"{base} ({counter}){ext}"
            if not os.path.exists(new_path): return new_path
            counter += 1

    def _clean_windows_duplicate_marker_rename(self, name: str) -> str:
        return re.sub(r'\s*\(\d+\)$', '', name.strip())

    def _normalize_model_string_rename(self, s: any) -> str:
        if s is None: return ""
        s_str = str(s)
        cleaned = re.sub(r'\s+', ' ', s_str).strip()
        return re.sub(r'[\W_]+', '', cleaned).lower()

    def _extract_date_from_val_rename(self, value: any) -> tuple[str, datetime | None]:
        if value is None or (isinstance(value, str) and not value.strip()): return 'EMPTY', None
        
        if hasattr(value, 'year') and hasattr(value, 'month') and hasattr(value, 'day'):
             try:
                 return 'VALID', datetime(value.year, value.month, value.day)
             except Exception:
                 pass
        
        if isinstance(value, str):
            date_str = value.strip()
            if not date_str: return 'EMPTY', None
            range_match = re.match(r'^\d{1,2}\s*-\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', date_str)
            if range_match: date_str = range_match.group(1)
            date_str = date_str.split('&')[0].strip()
            if not date_str: return 'EMPTY', None
            date_formats = ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y",
                            "%d/%m/%y", "%m/%d/%y", "%y-%m-%d", "%d-%m-%y", "%d.%m.%y")
            for fmt in date_formats:
                try:
                    dt_obj = datetime.strptime(date_str, fmt)
                    if dt_obj.year < 100:
                        current_year_base = datetime.now().year // 100 * 100
                        year_adjusted = current_year_base + dt_obj.year
                        if year_adjusted > datetime.now().year + 20: year_adjusted -= 100
                        dt_obj = dt_obj.replace(year=year_adjusted)
                    return 'VALID', dt_obj
                except ValueError: continue
            return 'TYPO', None
        return 'EMPTY', None
        
    # --- METODI PER "ORGANIZZA E STAMPA" ---
    def _create_organizza_tab(self, tab):
        main_frame = ttk.Frame(tab, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        desc_text = "Questa sezione analizza i file Excel dalla 'Cartella di origine', legge un codice ODC e li copia in sottocartelle. Poi permette di selezionare le cartelle per la stampa di gruppo."
        desc_label = ttk.Label(main_frame, text=desc_text, wraplength=850, justify=tk.LEFT, style='info.TLabel')
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor='w')

        org_frame = ttk.LabelFrame(main_frame, text="1. Elabora e Organizza per ODC", padding="10")
        org_frame.pack(fill=tk.X, pady=(0, 5))
        self._create_path_entry(org_frame, "Cartella di origine:", self.organizza_source_dir, 0, readonly=False, browse_command=lambda: self._select_folder_dialog(self.organizza_source_dir, "Seleziona cartella schede da organizzare"))
        self.organizza_run_button = ttk.Button(org_frame, text="🚀 AVVIA ORGANIZZAZIONE", style='primary.TButton', command=self._start_organizza_process)
        self.organizza_run_button.grid(row=1, column=0, columnspan=3, sticky="we", pady=(10,0))
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)
        print_frame = ttk.LabelFrame(main_frame, text="2. Stampa Schede Organizzate", padding="10")
        print_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        print_controls_frame = ttk.Frame(print_frame); print_controls_frame.pack(fill=tk.X, pady=(0, 10))
        self.stampa_run_button = ttk.Button(print_controls_frame, text="🖨️ STAMPA SELEZIONATE", command=self._start_stampa_process)
        self.stampa_run_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.stampa_refresh_button = ttk.Button(print_controls_frame, text="🔄 AGGIORNA LISTA", command=self._populate_stampa_list)
        self.stampa_refresh_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        list_container = ttk.Frame(print_frame); list_container.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(list_container, borderwidth=0, highlightthickness=0)
        self.stampa_checkbox_frame = ttk.Frame(canvas)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set); scrollbar.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        self.canvas_window = canvas.create_window((0, 0), window=self.stampa_checkbox_frame, anchor="nw")
        self.stampa_checkbox_vars = {}
        self.stampa_checkbox_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(self.canvas_window, width=e.width))
        self._populate_stampa_list()
        log_frame = ttk.LabelFrame(main_frame, text="Log Esecuzione (Organizza/Stampa)", padding="10")
        log_frame.pack(fill=tk.X, pady=(15, 0))
        self.organizza_log_area = self._create_log_widget(log_frame)

    # --- METODI PER "STAMPA CANONI MENSILI" ---
    def _create_stampa_canoni_tab(self, tab):
        main_frame = ttk.Frame(tab, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        desc_text = "Questa sezione automatizza la stampa dei canoni mensili. Seleziona i file e il periodo, poi avvia il processo per eseguire macro VBA e stampare documenti Word in sequenza."
        desc_label = ttk.Label(main_frame, text=desc_text, wraplength=850, justify=tk.LEFT, style='info.TLabel')
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor='w')

        paths_frame = ttk.LabelFrame(main_frame, text="1. File, Periodo e Macro per Stampa Canoni", padding="10")
        paths_frame.pack(fill=tk.X, pady=(0, 5))

        word_ft = [("File Word", "*.docx *.doc"), ("Tutti i file", "*.*")]

        # --- Selettori Anno/Mese ---
        giornaliera_selection_frame = ttk.Frame(paths_frame)
        giornaliera_selection_frame.grid(row=0, column=0, columnspan=3, sticky=tk.EW, pady=3)
        
        ttk.Label(giornaliera_selection_frame, text="Anno Giornaliera:").pack(side=tk.LEFT, padx=(0,5))
        self.canoni_anno_combo = ttk.Combobox(giornaliera_selection_frame, textvariable=self.canoni_selected_year, values=self.anni_giornaliera, state="readonly", width=10)
        self.canoni_anno_combo.pack(side=tk.LEFT, padx=(0,10))
        self.canoni_anno_combo.bind("<<ComboboxSelected>>", self._update_paths_from_ui)
        
        ttk.Label(giornaliera_selection_frame, text="Mese Giornaliera:").pack(side=tk.LEFT, padx=(5,5))
        self.canoni_mese_combo = ttk.Combobox(giornaliera_selection_frame, textvariable=self.canoni_selected_month, values=self.nomi_mesi_italiani, state="readonly", width=15)
        self.canoni_mese_combo.pack(side=tk.LEFT)
        self.canoni_mese_combo.bind("<<ComboboxSelected>>", self._update_paths_from_ui)

        try:
            current_year_str = str(datetime.now().year)
            if current_year_str in self.anni_giornaliera: self.canoni_anno_combo.set(current_year_str)
            else: self.canoni_anno_combo.set(self.anni_giornaliera[0])
            self.canoni_mese_combo.current(datetime.now().month - 1)
        except Exception:
            if self.anni_giornaliera: self.canoni_anno_combo.set(self.anni_giornaliera[0])
            self.canoni_mese_combo.current(0)
        
        # --- Campi di Input ---
        self._create_path_entry(paths_frame, "File Giornaliera (Auto):", self.canoni_giornaliera_path, 1, readonly=True)

        self._create_path_entry(paths_frame, "N° Canone Messina:", self.canoni_messina_num, 2, readonly=False)
        self._create_path_entry(paths_frame, "N° Canone Naselli:", self.canoni_naselli_num, 3, readonly=False)
        self._create_path_entry(paths_frame, "N° Canone Caldarella:", self.canoni_caldarella_num, 4, readonly=False)

        # Associa l'aggiornamento dei percorsi alla modifica dei campi numero
        self.canoni_messina_num.trace_add("write", self._update_paths_from_ui)
        self.canoni_naselli_num.trace_add("write", self._update_paths_from_ui)
        self.canoni_caldarella_num.trace_add("write", self._update_paths_from_ui)

        self._create_path_entry(paths_frame, "File Foglio Canone (Word):", self.canoni_word_path, 5, readonly=False, browse_command=lambda: self._select_file_dialog(self.canoni_word_path, "Seleziona Foglio Canone Word", word_ft))
        
        ttk.Label(paths_frame, text="Seleziona Stampante:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=3)
        self.printer_combo = ttk.Combobox(paths_frame, textvariable=self.selected_printer, state="readonly")
        self.printer_combo.grid(row=6, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=3)
        
        self._create_path_entry(paths_frame, "Nome Macro VBA:", self.canoni_macro_name, 7, readonly=True)
        
        # --- Frame Azioni e Log ---
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)
        actions_frame = ttk.LabelFrame(main_frame, text="2. Azione Stampa Canoni", padding="10")
        actions_frame.pack(fill=tk.X, pady=5)
        self.canoni_run_button = ttk.Button(actions_frame, text="▶  AVVIA PROCESSO STAMPA CANONI", style='primary.TButton', command=self._start_stampa_canoni_process)
        self.canoni_run_button.pack(fill=tk.X, ipady=8, pady=5)

        log_frame = ttk.LabelFrame(main_frame, text="Log Esecuzione (Stampa Canoni)", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        self.canoni_log_area = self._create_log_widget(log_frame)
        
        self.after(100, self._update_paths_from_ui)

    def _start_firma_process(self):
        if not self._validate_firma_paths(): return
        self._clear_log(self.firma_log_area)
        self._toggle_firma_buttons('disabled')
        threading.Thread(target=self._run_firma_thread, daemon=True).start()

    def _run_firma_thread(self):
        pythoncom.CoInitialize()
        try:
            self._log(self.firma_log_area, "--- FASE 1: ELABORAZIONE EXCEL E CONVERSIONE PDF ---", 'HEADER')
            if not self._process_excel_files_firma():
                self._log(self.firma_log_area, "Fase 1 terminata con errori. Processo interrotto.", 'ERROR'); return
            self._log(self.firma_log_area, "--- FASE 2: COMPRESSIONE PDF ---", 'HEADER')
            self._compress_pdfs_firma()
            self._log(self.firma_log_area, "--- PROCESSO FIRMA COMPLETATO ---", 'SUCCESS')
        except Exception as e: self._log(self.firma_log_area, f"ERRORE CRITICO: {e}", "ERROR")
        finally: self._toggle_firma_buttons('normal'); pythoncom.CoUninitialize()
        
    def _validate_firma_paths(self):
        paths_to_check = { "Immagine Firma": self.firma_image_path.get(), "Eseguibile Ghostscript": self.firma_ghostscript_path.get() }
        for name, path in paths_to_check.items():
            if not path or not os.path.isfile(path):
                self._log(self.firma_log_area, f"ERRORE: '{name}' non trovato: {path}", 'ERROR'); return False
        return True

    def _process_excel_files_firma(self):
        excel_path = self.firma_excel_dir.get()
        if not os.path.isdir(excel_path):
            self._log(self.firma_log_area, f"Cartella input non trovata: '{excel_path}'", "ERROR"); return False
        excel_files = [f for f in os.listdir(excel_path) if f.lower().endswith(('.xlsx', '.xls', '.xlsm'))]
        if not excel_files:
            self._log(self.firma_log_area, f"Nessun file Excel in: {FIRMA_EXCEL_INPUT_DIR}", 'WARNING'); return True
        self._log(self.firma_log_area, f"Trovati {len(excel_files)} file Excel.")
        excel = None
        try:
            excel = win32com.client.Dispatch('Excel.Application'); excel.Visible = False; excel.DisplayAlerts = False
            self._log(self.firma_log_area, "Applicazione Excel inizializzata.", 'INFO')
            mode = self.firma_processing_mode.get()
            for file_name in excel_files:
                file_path = os.path.join(excel_path, file_name)
                self._log(self.firma_log_area, "-" * 50); self._log(self.firma_log_area, f"Elaborazione: {file_name}", 'INFO')
                workbook = excel.Workbooks.Open(file_path, 0, True) 
                if mode == "schede": self._firma_mode_schede(workbook, file_name)
                elif mode == "preventivi": self._firma_mode_preventivi(workbook, file_name)
                workbook.Close(SaveChanges=False)
        except Exception as e: self._log(self.firma_log_area, f"ERRORE Excel: {e}", 'ERROR'); return False
        finally:
            if excel: excel.Quit()
            self._log(self.firma_log_area, "Applicazione Excel chiusa.", 'INFO')
        return True
    
    def _firma_mode_schede(self, workbook, file_name):
        try:
            ws = workbook.Worksheets(1)
            valE2 = ws.Cells(2, 5).Text.strip(); valT2 = ws.Cells(2, 20).Text.strip(); valT5 = ws.Cells(5, 20).Text.strip()
            model_value = valE2 or valT2 or valT5
            cleaned_model = ''.join(filter(str.isalnum, model_value))
            self._log(self.firma_log_area, f"Modello: '{cleaned_model}'", 'INFO')
            if cleaned_model in self.firma_processing_data:
                data = self.firma_processing_data[cleaned_model]; ws.PageSetup.PrintArea = data["PrintArea"]
                img_width, img_height = (105, 35) if cleaned_model == "SCHEDAMANUTENZIONE" else (150, 50)
                cell_address = data["FirmaCella"]; col_str = ''.join(re.findall("[A-Z]+", cell_address)); row_str = ''.join(re.findall(r"\d+", cell_address))
                target_cell = ws.Cells(int(row_str), self._col_to_num(col_str))
                points_per_cm = 28.35; offset_1cm = 1.0 * points_per_cm; offset_03cm = 0.3 * points_per_cm
                top_pos = max(0, target_cell.Top - (offset_03cm if cleaned_model == "SCHEDAMANUTENZIONE" else offset_1cm))
                left_pos = max(0, target_cell.Left - offset_1cm)
                ws.Shapes.AddPicture(self.firma_image_path.get(), True, True, left_pos, top_pos, img_width, img_height)
                self._log(self.firma_log_area, "Immagine firma aggiunta.", 'SUCCESS')
                pdf_file_path = os.path.join(self.firma_pdf_dir.get(), f"{os.path.splitext(file_name)[0]}.pdf")
                workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_file_path)
                self._log(self.firma_log_area, f"File PDF esportato.", 'SUCCESS')
            else: self._log(self.firma_log_area, f"Modello non gestito.", 'WARNING')
        except Exception as e: self._log(self.firma_log_area, f"ERRORE: {e}", 'ERROR')
    
    def _firma_mode_preventivi(self, workbook, file_name):
        try:
            ws = next((s for s in workbook.Worksheets if s.Name == "Consuntivo"), None)
            if ws is None: self._log(self.firma_log_area, "Foglio 'Consuntivo' non trovato.", 'WARNING'); return
            ws.Activate(); ws.PageSetup.PrintArea = "A3:L63"
            target_cell = ws.Cells(59, 3); top_position = target_cell.Top + 10
            ws.Shapes.AddPicture(self.firma_image_path.get(), True, True, target_cell.Left, top_position, 150, 50)
            self._log(self.firma_log_area, "Immagine firma aggiunta.", 'SUCCESS')
            pdf_file_path = os.path.join(self.firma_pdf_dir.get(), f"{os.path.splitext(file_name)[0]}.pdf")
            ws.ExportAsFixedFormat(0, pdf_file_path); self._log(self.firma_log_area, f"File PDF esportato.", 'SUCCESS')
        except Exception as e: self._log(self.firma_log_area, f"ERRORE: {e}", 'ERROR')

    def _compress_pdfs_firma(self):
        pdf_path = self.firma_pdf_dir.get()
        pdf_files = [f for f in os.listdir(pdf_path) if f.lower().endswith('.pdf')]
        if not pdf_files: self._log(self.firma_log_area, "Nessun PDF da comprimere.", 'WARNING'); return
        self._log(self.firma_log_area, f"Trovati {len(pdf_files)} PDF.")
        gs_exe = self.firma_ghostscript_path.get()
        for pdf_file in pdf_files:
            input_pdf = os.path.join(pdf_path, pdf_file); temp_output_pdf = os.path.join(pdf_path, f"temp_{pdf_file}")
            self._log(self.firma_log_area, f"Compressione: {pdf_file}", 'INFO')
            args = [gs_exe, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", "-dPDFSETTINGS=/ebook", "-dNOPAUSE", "-dBATCH", "-dQUIET", f"-sOutputFile={temp_output_pdf}", input_pdf]
            try:
                subprocess.run(args, check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
                if os.path.exists(temp_output_pdf) and os.path.getsize(temp_output_pdf) > 100:
                    os.remove(input_pdf); os.rename(temp_output_pdf, input_pdf); self._log(self.firma_log_area, "Compressione OK.", 'SUCCESS')
                else:
                    self._log(self.firma_log_area, "ERRORE: File compresso non valido.", 'ERROR')
                    if os.path.exists(temp_output_pdf): os.remove(temp_output_pdf)
            except Exception as e:
                self._log(self.firma_log_area, f"ERRORE compressione: {e}", 'ERROR')
                if os.path.exists(temp_output_pdf): os.remove(temp_output_pdf)

    def _toggle_firma_buttons(self, state):
        for btn in [self.firma_run_button, self.firma_clean_pdf_button, self.firma_clean_excel_button]: btn.config(state=state)

    def _populate_stampa_list(self):
        for widget in self.stampa_checkbox_frame.winfo_children(): widget.destroy()
        self.stampa_checkbox_vars.clear()
        dest_path = self.organizza_dest_dir.get()
        if not os.path.isdir(dest_path): return
        try:
            folders = sorted([d for d in os.listdir(dest_path) if os.path.isdir(os.path.join(dest_path, d))])
            for folder_name in folders:
                var = tk.IntVar(); cb = ttk.Checkbutton(self.stampa_checkbox_frame, text=folder_name, variable=var)
                cb.pack(anchor="w", padx=5, fill='x')
                self.stampa_checkbox_vars[folder_name] = {"var": var, "path": os.path.join(dest_path, folder_name)}
        except Exception as e: self._log(self.organizza_log_area, f"Errore lettura cartelle: {e}", "ERROR")

    def _start_organizza_process(self):
        self._clear_log(self.organizza_log_area); self._toggle_organizza_buttons('disabled')
        threading.Thread(target=self._run_organizza_thread, daemon=True).start()

    def _run_organizza_thread(self):
        pythoncom.CoInitialize()
        try:
            self._clear_folder_content(self.organizza_dest_dir.get(), ORGANIZZA_DEST_DIR, self.organizza_log_area)
            self._organize_files()
            self.after(0, self._populate_stampa_list)
        finally: self._toggle_organizza_buttons('normal'); pythoncom.CoUninitialize()
    
    def _organize_files(self):
        source_dir, dest_dir, log_w = self.organizza_source_dir.get(), self.organizza_dest_dir.get(), self.organizza_log_area
        self._log(log_w, "--- Inizio Organizzazione ---", "HEADER")
        if not os.path.isdir(source_dir): self._log(log_w, f"ERRORE: Origine '{source_dir}' non esiste.", "ERROR"); return
        excel_ext = ('.xls', '.xlsx', '.xlsm', '.xlsb')
        try:
            excel_files = [os.path.join(r,f) for r, _, fs in os.walk(source_dir) for f in fs if f.lower().endswith(excel_ext)]
        except Exception as e: self._log(log_w, f"ERRORE accesso a '{source_dir}': {e}", "ERROR"); return
        if not excel_files: self._log(log_w, f"Nessun file Excel in: {source_dir}", "WARNING"); return
        self._log(log_w, f"Trovati {len(excel_files)} file Excel.")
        excel, proc_count = None, 0
        try:
            excel = win32com.client.Dispatch("Excel.Application"); excel.Visible = False; excel.DisplayAlerts = False
            for fp in excel_files:
                self._log(log_w, f"Processando: {os.path.basename(fp)}")
                wb = None
                try:
                    wb = excel.Workbooks.Open(fp); ws = wb.Worksheets(1)
                    odc_v = next((ws.Range(c).Value for c in ["L50","L45","DB14","DB17"] if ws.Range(c).Value is not None and str(ws.Range(c).Value).strip()!=""),None)
                    odc_s = str(int(odc_v)) if isinstance(odc_v,(int,float)) else (str(odc_v).strip() if isinstance(odc_v,str) else "")
                    wb.Close(SaveChanges=False); wb = None
                    dest_f_p = os.path.join(dest_dir, re.sub(r'[\\/:*?"<>|]', '', odc_s)) if odc_s and odc_s.upper()!="NA" else os.path.join(dest_dir, "Schede senza ODC")
                    os.makedirs(dest_f_p, exist_ok=True); shutil.copy2(fp, dest_f_p)
                    self._log(log_w, f"  -> Copiato in: {os.path.basename(dest_f_p)}", "SUCCESS"); proc_count+=1
                except Exception as e: self._log(log_w, f"ERRORE file {os.path.basename(fp)}: {e}", "ERROR");
                finally: 
                    if wb: wb.Close(SaveChanges=False)
        finally:
            if excel: excel.Quit()
            self._log(log_w, f"--- Organizzazione Completata ({proc_count}/{len(excel_files)}) ---", "HEADER")
            
    def _start_stampa_process(self):
        sel_folders = [d["path"] for d in self.stampa_checkbox_vars.values() if d["var"].get()==1]
        if not sel_folders: self._log(self.organizza_log_area, "Nessuna cartella per stampa.", "WARNING"); return
        self._toggle_organizza_buttons('disabled')
        threading.Thread(target=self._run_stampa_thread, args=(sel_folders,), daemon=True).start()

    def _run_stampa_thread(self, folders_to_print):
        pythoncom.CoInitialize()
        try:
            self._log(self.organizza_log_area, f"--- Avvio Stampa per {len(folders_to_print)} cartelle ---", "HEADER")
            self._print_files_in_folders(folders_to_print)
            self._log(self.organizza_log_area, "--- Stampa Completata ---", "SUCCESS")
        finally: self._toggle_organizza_buttons('normal'); pythoncom.CoUninitialize()

    def _print_files_in_folders(self, folder_list):
        log_w, excel = self.organizza_log_area, None
        excel_ext = ('.xls', '.xlsx', '.xlsm', '.xlsb')
        try:
            excel = win32com.client.Dispatch("Excel.Application"); excel.Visible = False; excel.DisplayAlerts = False
            for folder_p in folder_list:
                self._log(log_w, f"Stampa cartella: {os.path.basename(folder_p)}")
                excel_fs = [os.path.join(folder_p, f) for f in os.listdir(folder_p) if f.lower().endswith(excel_ext)]
                if not excel_fs: self._log(log_w, "  -> Nessun file Excel.", "WARNING"); continue
                for fp in excel_fs:
                    wb = None
                    try:
                        wb = excel.Workbooks.Open(fp); ws = wb.Worksheets(1)
                        m = next((str(ws.Cells(r,c).Value).strip() for r,c in [(2,5),(2,20),(5,20)] if ws.Cells(r,c).Value and str(ws.Cells(r,c).Value).strip()),"")
                        cm = re.sub(r'\W','',m)
                        if cm in self.stampa_processing_data:
                            ws.PageSetup.PrintArea = self.stampa_processing_data[cm]["PrintArea"]; wb.PrintOut()
                            self._log(log_w, f"  -> Stampa: {os.path.basename(fp)}", "SUCCESS")
                        else: self._log(log_w, f"  -> Ignorato (modello non trovato): {os.path.basename(fp)}", "WARNING")
                    finally: 
                        if wb: wb.Close(SaveChanges=False)
        finally:
            if excel: excel.Quit()
            
    def _toggle_organizza_buttons(self, state):
        for btn in [self.organizza_run_button, self.stampa_run_button, self.stampa_refresh_button]: btn.config(state=state)

    def _populate_printers(self):
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            self.printer_combo['values'] = printers
            
            saved_printer = self.selected_printer.get()
            if saved_printer and saved_printer in printers:
                self.printer_combo.set(saved_printer)
            else:
                default_printer = win32print.GetDefaultPrinter()
                if default_printer in printers:
                    self.printer_combo.set(default_printer)
                elif printers:
                    self.printer_combo.set(printers[0])
        except Exception as e:
            self._log(self.canoni_log_area, f"Errore nel caricamento delle stampanti: {e}", "ERROR")

    def _update_paths_from_ui(self, *args):
        self._update_giornaliera_path()
        self._update_consuntivo_paths()

    def _update_giornaliera_path(self):
        selected_month_name = self.canoni_selected_month.get()
        selected_year_str = self.canoni_selected_year.get()

        if selected_month_name and selected_year_str:
            month_number = self.mesi_giornaliera_map.get(selected_month_name)
            if month_number:
                year_folder_name = f"Giornaliere {selected_year_str}"
                file_name = f"Giornaliera {month_number}-{selected_year_str}.xlsm"
                full_path = os.path.join(CANONI_GIORNALIERA_BASE_DIR, year_folder_name, file_name)
                self.canoni_giornaliera_path.set(full_path)
            else:
                self.canoni_giornaliera_path.set("Mese non valido")
        else:
            self.canoni_giornaliera_path.set("Seleziona Anno e Mese")

    def _update_consuntivo_paths(self):
        year = self.canoni_selected_year.get()
        if not year:
            return

        cons_dir = os.path.join(CANONI_CONSUNTIVI_BASE_DIR, year, "CONSUNTIVI", year)

        nums_to_find = {
            "messina": self.canoni_messina_num.get(),
            "naselli": self.canoni_naselli_num.get(),
            "caldarella": self.canoni_caldarella_num.get()
        }
        
        path_vars = {
            "messina": self.canoni_cons1_path,
            "naselli": self.canoni_cons2_path,
            "caldarella": self.canoni_cons3_path
        }

        if not os.path.isdir(cons_dir):
            for var in path_vars.values():
                var.set(f"ERRORE: Cartella non trovata")
            return

        try:
            files_in_dir = os.listdir(cons_dir)
            for key, num_str in nums_to_find.items():
                if not num_str.strip().isdigit():
                    path_vars[key].set("Inserire un numero valido")
                    continue
                
                found_path = None
                for filename in files_in_dir:
                    if filename.startswith(f"{num_str}-") or filename.startswith(f"{num_str} "):
                        found_path = os.path.join(cons_dir, filename)
                        break
                
                if found_path:
                    path_vars[key].set(found_path)
                else:
                    path_vars[key].set(f"File non trovato per il n° {num_str}")

        except Exception as e:
            self._log(self.canoni_log_area, f"Errore ricerca consuntivi: {e}", "ERROR")

    def _start_stampa_canoni_process(self):
        if not self._validate_stampa_canoni_paths(): return
        self._clear_log(self.canoni_log_area)
        self._toggle_stampa_canoni_buttons('disabled')
        threading.Thread(target=self._run_stampa_canoni_thread, daemon=True).start()

    def _validate_stampa_canoni_paths(self):
        self._update_paths_from_ui()

        paths_to_check = {
            "File Giornaliera": self.canoni_giornaliera_path.get(),
            "Canone Messina": self.canoni_cons1_path.get(),
            "Canone Naselli": self.canoni_cons2_path.get(),
            "Canone Caldarella": self.canoni_cons3_path.get(),
            "File Foglio Canone": self.canoni_word_path.get(),
        }
        for name, path in paths_to_check.items():
            if not path or not os.path.isfile(path):
                self._log(self.canoni_log_area, f"ERRORE: Percorso per '{name}' non valido o file non trovato: '{path}'", 'ERROR')
                return False
        
        if not self.canoni_macro_name.get().strip():
            self._log(self.canoni_log_area, "ERRORE: Nome della macro VBA non specificato.", 'ERROR'); return False
            
        if not self.selected_printer.get():
             self._log(self.canoni_log_area, "ERRORE: Nessuna stampante selezionata.", 'ERROR'); return False
             
        return True

    def _run_stampa_canoni_thread(self):
        pythoncom.CoInitialize()
        excel_app, word_app = None, None
        docs_to_close = []  
        log_w = self.canoni_log_area
        macro_name = self.canoni_macro_name.get().strip()
        try:
            self._log(log_w, "Avvio applicazioni Excel e Word...", 'INFO')
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.DisplayAlerts = False
            word_app = win32com.client.Dispatch("Word.Application")

            selected_printer_name = self.selected_printer.get()
            if selected_printer_name:
                word_app.ActivePrinter = selected_printer_name
                self._log(log_w, f"Stampante impostata su: '{selected_printer_name}'", "SUCCESS")

            self._log(log_w, "--- Apertura file sorgente ---", 'HEADER')
            giornaliera_path = self.canoni_giornaliera_path.get()
            try:
                self._log(log_w, f"Apertura file Giornaliera: {os.path.basename(giornaliera_path)}...", 'INFO')
                wb_giornaliera = excel_app.Workbooks.Open(giornaliera_path)
                docs_to_close.append(wb_giornaliera)
                self._log(log_w, "File Giornaliera aperto. Rimarrà aperto per i riferimenti.", 'SUCCESS')
            except Exception as e:
                self._log(log_w, f"ERRORE CRITICO apertura Giornaliera: {e}", "ERROR"); raise
            
            cons_paths = [self.canoni_cons1_path.get(), self.canoni_cons2_path.get(), self.canoni_cons3_path.get()]
            word_path = self.canoni_word_path.get()
            wb_cons_list = []
            
            for i, p in enumerate(cons_paths):
                try:
                    self._log(log_w, f"Apertura Consuntivo {i+1}: {os.path.basename(p)}...", 'INFO')
                    wb = excel_app.Workbooks.Open(p); wb_cons_list.append(wb); docs_to_close.append(wb)
                except Exception as e: self._log(log_w, f"ERRORE apertura Consuntivo {i+1}: {e}", "ERROR"); raise
            
            try:
                self._log(log_w, f"Apertura documento Word: {os.path.basename(word_path)}...", 'INFO')
                doc_word = word_app.Documents.Open(word_path); docs_to_close.append(doc_word)
            except Exception as e: self._log(log_w, f"ERRORE apertura Documento Word: {e}", "ERROR"); raise
            
            self._log(log_w, "--- Inizio sequenza operazioni ---", 'HEADER')
            for i, cons_wb in enumerate(wb_cons_list):
                leaf_name = cons_wb.Name
                self._log(log_w, f"Esecuzione macro '{macro_name}' su {leaf_name}...", 'INFO')
                excel_app.Run(f"'{leaf_name}'!{macro_name}")
                self._log(log_w, f"Macro su Consuntivo {i+1} completata.", 'SUCCESS')
                if i < len(wb_cons_list) - 1:
                    self._log(log_w, f"Stampa file Word: {doc_word.Name}...", 'INFO')
                    doc_word.PrintOut()
                    self._log(log_w, "Comando di stampa Word inviato.", 'SUCCESS')
                    
            self._log(log_w, "--- PROCESSO STAMPA CANONI COMPLETATO ---", 'SUCCESS')
        except Exception as e:
            self._log(log_w, f"ERRORE CRITICO nel processo: {e}", "ERROR")
            self._log(log_w, traceback.format_exc(), "ERROR")
            self._log(log_w, "Il processo è stato interrotto.", "WARNING")
        finally:
            self._log(log_w, "Chiusura file e applicazioni...", 'INFO')
            for doc in reversed(docs_to_close): 
                try: doc.Close(SaveChanges=0)
                except Exception as e_cl: self._log(log_w, f"Errore durante la chiusura di un documento: {e_cl}", "WARNING")
            if word_app: word_app.Quit(); self._log(log_w, "Applicazione Word chiusa.", 'INFO')
            if excel_app: excel_app.Quit(); self._log(log_w, "Applicazione Excel chiusa.", 'INFO')
            self._toggle_stampa_canoni_buttons('normal'); pythoncom.CoUninitialize()

    def _toggle_stampa_canoni_buttons(self, state):
        self.canoni_run_button.config(state=state)

    # --- METODI GENERICI E DI UTILITÀ ---
    def _create_path_entry(self, parent, label_text, string_var, row, readonly=False, browse_command=None):
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        entry_width = 60 if browse_command or readonly else 70 
        entry = ttk.Entry(parent, textvariable=string_var, width=entry_width)
        if readonly: entry.config(state='readonly')
        entry.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)
        if browse_command:
            button = ttk.Button(parent, text="Sfoglia...", command=browse_command, width=10)
            button.grid(row=row, column=2, sticky=tk.E, padx=5, pady=3)
            parent.columnconfigure(1, weight=1) 
        else: parent.grid_columnconfigure(1, weight=1)

    def _select_folder_dialog(self, string_var, title):
        folder_selected = filedialog.askdirectory(title=title)
        if folder_selected: string_var.set(folder_selected)

    def _select_file_dialog(self, string_var, title, filetypes, initialdir=None):
        file_selected = filedialog.askopenfilename(title=title, filetypes=filetypes, initialdir=initialdir)
        if file_selected: string_var.set(file_selected)

    def _create_log_widget(self, parent):
        log_widget = scrolledtext.ScrolledText(parent, wrap=tk.WORD, height=10, font=("Courier New", 9))
        log_widget.pack(fill=tk.BOTH, expand=True); log_widget.config(state='disabled')
        tags = {'INFO':'black', 'SUCCESS':'#008744', 'WARNING':'#ffa700', 'ERROR':'#d62d20', 'HEADER':'#0057e7'}
        for tag, color in tags.items():
            fw = "bold" if tag in ['SUCCESS', 'ERROR', 'HEADER'] else "normal"
            log_widget.tag_config(tag.upper(), foreground=color, font=("Courier New", 9, fw))
        return log_widget

    def _log(self, widget, message, level='INFO'):
        if not widget: return
        self.after(0, self._update_log_on_main_thread, widget, message, level)

    def _update_log_on_main_thread(self, widget, message, level):
        widget.config(state='normal'); widget.insert(tk.END, f"> {message}\n", level.upper())
        widget.config(state='disabled'); widget.see(tk.END); self.update_idletasks()

    def _clear_log(self, widget):
        widget.config(state='normal'); widget.delete(1.0, tk.END); widget.config(state='disabled')

    def _clear_folder_content(self, folder_path, folder_display_name, log_widget):
        self._log(log_widget, f"--- Pulizia '{folder_display_name}' in corso... ---", 'HEADER')
        if os.path.isdir(folder_path):
            for item_name in os.listdir(folder_path):
                item_path = os.path.join(folder_path, item_name)
                try:
                    if os.path.isdir(item_path): shutil.rmtree(item_path)
                    else: os.remove(item_path)
                except Exception as e: self._log(log_widget, f"Impossibile eliminare {item_name}: {e}", 'ERROR')
        self._log(log_widget, f"--- Pulizia '{folder_display_name}' completata. ---", 'SUCCESS')

    def _start_folder_cleanup(self, folder_to_clean, folder_name, log_widget, toggle_func):
        toggle_func('disabled')
        threading.Thread(target=self._run_folder_cleanup_thread, args=(folder_to_clean, folder_name, log_widget, toggle_func), daemon=True).start()

    def _run_folder_cleanup_thread(self, folder_to_clean, folder_display_name, log_widget, toggle_buttons_func):
        self._clear_folder_content(folder_to_clean, folder_display_name, log_widget)
        toggle_buttons_func('normal')
        
    def _col_to_num(self, col_str):
        num = 0
        for char in col_str: num = num * 26 + (ord(char.upper()) - ord('A')) + 1
        return num

    # --- GESTIONE SALVATAGGIO/CARICAMENTO CONFIGURAZIONE ---
    def _load_config(self):
        config_path = os.path.join(application_path, CONFIG_FILE_NAME)
        defaults = {
            "firma_ghostscript_path": r"C:\Program Files\gs\gs10.05.0\bin\gswin64c.exe",
            "rinomina_path": os.path.join(application_path, RINOMINA_DEFAULT_DIR),
            "organizza_source_dir": os.path.join(application_path, ORGANIZZA_SOURCE_DIR),
            "canoni_selected_year": str(datetime.now().year),
            "canoni_selected_month": self.nomi_mesi_italiani[datetime.now().month - 1],
            "canoni_messina_num": "", "canoni_naselli_num": "", "canoni_caldarella_num": "",
            "canoni_word_path": CANONI_WORD_DEFAULT_PATH,
            "selected_printer": "" 
        }
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r') as f: settings = json.load(f)
            else: settings = defaults
        except (json.JSONDecodeError, IOError) as e:
            print(f"Errore caricamento {CONFIG_FILE_NAME}: {e}. Uso valori predefiniti.")
            settings = defaults

        self.firma_ghostscript_path.set(settings.get("firma_ghostscript_path", defaults["firma_ghostscript_path"]))
        self.rinomina_path.set(settings.get("rinomina_path", defaults["rinomina_path"]))
        self.organizza_source_dir.set(settings.get("organizza_source_dir", defaults["organizza_source_dir"]))
        
        loaded_year = settings.get("canoni_selected_year", defaults["canoni_selected_year"])
        if loaded_year in self.anni_giornaliera: self.canoni_selected_year.set(loaded_year)
        else: self.canoni_selected_year.set(defaults["canoni_selected_year"])
        
        loaded_month = settings.get("canoni_selected_month", defaults["canoni_selected_month"])
        if loaded_month in self.nomi_mesi_italiani: self.canoni_selected_month.set(loaded_month)
        else: self.canoni_selected_month.set(defaults["canoni_selected_month"])
        
        if hasattr(self, 'canoni_anno_combo'): self.canoni_anno_combo.set(self.canoni_selected_year.get())
        if hasattr(self, 'canoni_mese_combo'): self.canoni_mese_combo.set(self.canoni_selected_month.get())

        self.canoni_messina_num.set(settings.get("canoni_messina_num", defaults["canoni_messina_num"]))
        self.canoni_naselli_num.set(settings.get("canoni_naselli_num", defaults["canoni_naselli_num"]))
        self.canoni_caldarella_num.set(settings.get("canoni_caldarella_num", defaults["canoni_caldarella_num"]))
        self.canoni_word_path.set(settings.get("canoni_word_path", defaults["canoni_word_path"]))
        
        self.selected_printer.set(settings.get("selected_printer", defaults["selected_printer"]))

        self._update_paths_from_ui()

    def _save_config(self):
        settings = {
            "firma_ghostscript_path": self.firma_ghostscript_path.get(),
            "rinomina_path": self.rinomina_path.get(),
            "organizza_source_dir": self.organizza_source_dir.get(),
            "canoni_selected_year": self.canoni_selected_year.get(),
            "canoni_selected_month": self.canoni_selected_month.get(),
            "canoni_messina_num": self.canoni_messina_num.get(),
            "canoni_naselli_num": self.canoni_naselli_num.get(),
            "canoni_caldarella_num": self.canoni_caldarella_num.get(),
            "canoni_word_path": self.canoni_word_path.get(),
            "selected_printer": self.selected_printer.get()
        }
        config_path = os.path.join(application_path, CONFIG_FILE_NAME)
        try:
            with open(config_path, 'w') as f:
                json.dump(settings, f, indent=4)
        except IOError as e:
            print(f"Errore salvataggio configurazione: {e}")

    def _on_closing(self):
        self._save_config()
        self.destroy()

if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()