import tkinter as tk
from tkinter import ttk
import os
from datetime import datetime, timedelta

from src.utils import constants as const
from src.utils.config_manager import ConfigManager
from src.utils.ui_utils import create_log_widget, log_message, clear_log

from src.gui.tabs.signature_tab import SignatureTab
from src.gui.tabs.rename_tab import RenameTab
from src.gui.tabs.organize_tab import OrganizeTab
from src.gui.tabs.fees_tab import FeesTab
from src.logic.monthly_fees import MonthlyFeesProcessor

class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestione Documenti Ufficio (Refactored)")
        try:
            self.state('zoomed')
        except tk.TclError:
            self.geometry("1200x900")
            self.center_window(1200, 900)
        self.resizable(True, True)

        self.config_manager = ConfigManager()
        self.config_manager.load()

        self._initialize_stringvars()
        self._setup_style()
        self._create_widgets()
        self._load_config_into_vars()

        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def center_window(self, width, height):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def _setup_style(self):
        self.font_main = ("Segoe UI", 10)
        self.font_bold = ("Segoe UI", 11, "bold")
        self.background_color = "#f0f0f0"

        style = ttk.Style(self)
        style.theme_use('clam')

        style.configure('.', font=self.font_main, background=self.background_color)
        style.configure('TLabel', font=self.font_main, background=self.background_color)
        style.configure('TLabelframe', background=self.background_color, bordercolor="#cccccc")
        style.configure('TLabelframe.Label', font=self.font_bold, background=self.background_color)
        style.configure('info.TLabel', foreground='#333333', background=self.background_color)

        style.configure('TButton', padding=6, font=self.font_main)
        style.map('TButton',
                  background=[('active', '#e0e0e0')],
                  foreground=[('disabled', '#a0a0a0')])

        style.configure('primary.TButton', background='#0078D4', foreground='white', font=self.font_bold)
        style.map('primary.TButton',
                  background=[('active', '#005a9e'), ('disabled', '#a0a0a0')],
                  foreground=[('disabled', '#ffffff')])

        style.configure('TNotebook', background=self.background_color, borderwidth=0)
        style.configure('TNotebook.Tab', padding=[12, 6], font=self.font_main)
        style.map('TNotebook.Tab',
                  background=[('selected', self.background_color), ('!selected', '#d0d0d0')],
                  expand=[("selected", [0, 2, 0, 0])])

    def _initialize_stringvars(self):
        # ... (this method is unchanged)
        self.FIRMA_EXCEL_INPUT_DIR = const.FIRMA_EXCEL_INPUT_DIR
        self.ORGANIZZA_DEST_DIR = const.ORGANIZZA_DEST_DIR
        self.CANONI_GIORNALIERA_BASE_DIR = const.CANONI_GIORNALIERA_BASE_DIR
        self.CANONI_CONSUNTIVI_BASE_DIR = const.CANONI_CONSUNTIVI_BASE_DIR
        self.mesi_giornaliera_map = const.MESI_GIORNALIERA_MAP
        self.nomi_mesi_italiani = const.NOMI_MESI_ITALIANI
        self.TCL_CONTACTS = const.TCL_CONTACTS
        self.EMAIL_BODY_INFORMAL = const.EMAIL_BODY_INFORMAL
        self.EMAIL_BODY_FORMAL = const.EMAIL_BODY_FORMAL
        self.EMAIL_BODY_GENERIC_INFORMAL = const.EMAIL_BODY_GENERIC_INFORMAL
        self.EMAIL_BODY_GENERIC_FORMAL = const.EMAIL_BODY_GENERIC_FORMAL
        self.firma_excel_dir = tk.StringVar(value=os.path.join(const.APPLICATION_PATH, const.FIRMA_EXCEL_INPUT_DIR))
        self.firma_image_path = tk.StringVar(value=os.path.join(const.APPLICATION_PATH, 'src', 'assets', const.FIRMA_IMAGE_NAME))
        self.firma_pdf_dir = tk.StringVar(value=os.path.join(const.APPLICATION_PATH, const.FIRMA_PDF_OUTPUT_DIR))
        self.firma_ghostscript_path = tk.StringVar()
        self.firma_processing_mode = tk.StringVar(value="schede")
        self.email_to = tk.StringVar()
        self.email_subject = tk.StringVar()
        self.email_tcl = tk.StringVar()
        self.email_is_formal = tk.BooleanVar(value=False)
        self.email_size_limit = tk.StringVar(value="6")
        self.rinomina_path = tk.StringVar()
        self.rinomina_password = tk.StringVar()
        self.organizza_source_dir = tk.StringVar()
        self.organizza_dest_dir = tk.StringVar(value=os.path.join(const.APPLICATION_PATH, const.ORGANIZZA_DEST_DIR))
        self.canoni_selected_year = tk.StringVar()
        self.canoni_selected_month = tk.StringVar()
        self.canoni_messina_num = tk.StringVar()
        self.canoni_naselli_num = tk.StringVar()
        self.canoni_caldarella_num = tk.StringVar()
        self.canoni_word_path = tk.StringVar()
        self.selected_printer = tk.StringVar()
        self.canoni_macro_name = tk.StringVar(value=const.DEFAULT_MACRO_NAME)
        self.canoni_giornaliera_path = tk.StringVar()
        self.canoni_cons1_path = tk.StringVar()
        self.canoni_cons2_path = tk.StringVar()
        self.canoni_cons3_path = tk.StringVar()

    def _load_config_into_vars(self):
        # ... (this method is unchanged)
        self.firma_ghostscript_path.set(self.config_manager.get("firma_ghostscript_path"))
        self.rinomina_path.set(self.config_manager.get("rinomina_path"))
        self.rinomina_password.set(self.config_manager.get("rinomina_password"))
        today = datetime.now()
        prev_month_date = today - timedelta(days=20)
        prev_month_year_str = str(prev_month_date.year)
        fees_tab_month_name = const.NOMI_MESI_ITALIANI[prev_month_date.month - 1]
        self.canoni_selected_year.set(prev_month_year_str)
        self.canoni_selected_month.set(fees_tab_month_name)
        organize_folder_month_str = f"{prev_month_date.month:02d} - {fees_tab_month_name.upper()}"
        organize_default_path = os.path.join(const.ORGANIZZA_BASE_DIR, prev_month_year_str, organize_folder_month_str)
        self.organizza_source_dir.set(organize_default_path)
        self.canoni_messina_num.set(self.config_manager.get("canoni_messina_num"))
        self.canoni_naselli_num.set(self.config_manager.get("canoni_naselli_num"))
        self.canoni_caldarella_num.set(self.config_manager.get("canoni_caldarella_num"))
        self.canoni_word_path.set(self.config_manager.get("canoni_word_path"))
        self.selected_printer.set(self.config_manager.get("selected_printer"))
        self.email_to.set(self.config_manager.get("email_to"))
        self.email_subject.set(self.config_manager.get("email_subject"))
        self.email_tcl.set(self.config_manager.get("email_tcl"))
        self.email_is_formal.set(self.config_manager.get("email_is_formal"))
        self.email_size_limit.set(self.config_manager.get("email_size_limit"))

    def _create_widgets(self):
        self.configure(background=self.background_color)
        main_container = ttk.Frame(self, padding="10")
        main_container.pack(fill=tk.BOTH, expand=True)

        # --- Notebook for Tabs ---
        notebook = ttk.Notebook(main_container)
        notebook.pack(expand=True, fill='both')

        # --- Create Tab Containers ---
        self.firma_container = ttk.Frame(notebook, padding="15")
        self.rinomina_container = ttk.Frame(notebook, padding="15")
        self.organizza_container = ttk.Frame(notebook, padding="15")
        self.canoni_container = ttk.Frame(notebook, padding="15")

        self.firma_container.columnconfigure(0, weight=1)
        self.rinomina_container.columnconfigure(0, weight=1)
        self.organizza_container.columnconfigure(0, weight=1)
        self.canoni_container.columnconfigure(0, weight=1)

        notebook.add(self.firma_container, text=' Apponi Firma ')
        notebook.add(self.rinomina_container, text=' Aggiungi Data Schede ')
        notebook.add(self.organizza_container, text=' Organizza e Stampa Schede ')
        notebook.add(self.canoni_container, text=' Stampa Canoni Mensili ')

        # --- Create Log Widgets ---
        self.log_widget_firma = self._create_log_frame(self.firma_container, "Log Esecuzione (Firma)")
        self.log_widget_rinomina = self._create_log_frame(self.rinomina_container, "Log Esecuzione (Aggiungi Data)")
        self.log_widget_organizza = self._create_log_frame(self.organizza_container, "Log Esecuzione (Organizza/Stampa)")
        self.log_widget_canoni = self._create_log_frame(self.canoni_container, "Log Esecuzione (Stampa Canoni)")

        # --- Dependency Injection and Tab Creation ---
        self.signature_tab = SignatureTab(self.firma_container, self, lambda msg, level='INFO': log_message(self.log_widget_firma, msg, level))
        self.signature_tab.pack(fill='both', expand=True)
        self.log_widget_firma.master.pack_forget() # Hide log frame initially
        self.log_widget_firma.master.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))

        self.rename_tab = RenameTab(self.rinomina_container, self, lambda msg, level='INFO': log_message(self.log_widget_rinomina, msg, level))
        self.rename_tab.pack(fill='both', expand=True)
        self.log_widget_rinomina.master.pack_forget()
        self.log_widget_rinomina.master.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))

        self.fees_tab = FeesTab(self.canoni_container, self, lambda msg, level='INFO': log_message(self.log_widget_canoni, msg, level))
        self.fees_tab.pack(fill='both', expand=True)
        self.log_widget_canoni.master.pack_forget()
        self.log_widget_canoni.master.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))

        self.organize_tab = OrganizeTab(self.organizza_container, self, lambda msg, level='INFO': log_message(self.log_widget_organizza, msg, level), self.fees_tab.processor)
        self.organize_tab.pack(fill='both', expand=True)
        self.log_widget_organizza.master.pack_forget()
        self.log_widget_organizza.master.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))

    def _create_log_frame(self, parent, title):
        log_frame = ttk.LabelFrame(parent, text=title, padding="10")
        # The frame is packed by the caller
        log_widget = create_log_widget(log_frame)
        return log_widget

    def _on_closing(self):
        # ... (this method is unchanged)
        current_config = {
            "firma_ghostscript_path": self.firma_ghostscript_path.get(),
            "rinomina_path": self.rinomina_path.get(),
            "rinomina_password": self.rinomina_password.get(),
            "canoni_messina_num": self.canoni_messina_num.get(),
            "canoni_naselli_num": self.canoni_naselli_num.get(),
            "canoni_caldarella_num": self.canoni_caldarella_num.get(),
            "canoni_word_path": self.canoni_word_path.get(),
            "selected_printer": self.selected_printer.get(),
            "email_to": self.email_to.get(),
            "email_subject": self.email_subject.get(),
            "email_tcl": self.email_tcl.get(),
            "email_is_formal": self.email_is_formal.get(),
            "email_size_limit": self.email_size_limit.get()
        }
        self.config_manager.save(current_config)
        self.destroy()
