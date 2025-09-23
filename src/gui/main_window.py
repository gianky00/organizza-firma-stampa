import tkinter as tk
from tkinter import ttk
import os
from datetime import datetime

from src.utils import constants as const
from src.utils.config_manager import ConfigManager
from src.utils.ui_utils import create_log_widget, log_message, clear_log

# Import tab classes
from src.gui.tabs.signature_tab import SignatureTab
from src.gui.tabs.rename_tab import RenameTab
from src.gui.tabs.organize_tab import OrganizeTab
from src.gui.tabs.fees_tab import FeesTab

class MainApplication(tk.Tk):
    """
    The main window of the application, hosting the notebook with all the tabs.
    """
    def __init__(self):
        super().__init__()
        self.title("Gestione Documenti Ufficio (Refactored)")
        try:
            self.state('zoomed')
        except tk.TclError:
            # Fallback for other OSes or environments that don't support 'zoomed'
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

    def _initialize_stringvars(self):
        """
        Initializes all tk.StringVars that hold the application's state.
        """
        # Common constants
        self.FIRMA_EXCEL_INPUT_DIR = const.FIRMA_EXCEL_INPUT_DIR
        self.ORGANIZZA_DEST_DIR = const.ORGANIZZA_DEST_DIR
        self.CANONI_GIORNALIERA_BASE_DIR = const.CANONI_GIORNALIERA_BASE_DIR
        self.CANONI_CONSUNTIVI_BASE_DIR = const.CANONI_CONSUNTIVI_BASE_DIR
        self.mesi_giornaliera_map = const.MESI_GIORNALIERA_MAP
        self.nomi_mesi_italiani = const.NOMI_MESI_ITALIANI

        # Signature Tab Vars
        self.firma_excel_dir = tk.StringVar(value=os.path.join(const.APPLICATION_PATH, const.FIRMA_EXCEL_INPUT_DIR))
        self.firma_image_path = tk.StringVar(value=os.path.join(const.APPLICATION_PATH, 'src', 'assets', const.FIRMA_IMAGE_NAME))
        self.firma_pdf_dir = tk.StringVar(value=os.path.join(const.APPLICATION_PATH, const.FIRMA_PDF_OUTPUT_DIR))
        self.firma_ghostscript_path = tk.StringVar()
        self.firma_processing_mode = tk.StringVar(value="schede")
        self.email_to = tk.StringVar()
        self.email_subject = tk.StringVar()
        # The email body doesn't use a StringVar as it's a Text widget

        # Rename Tab Vars
        self.rinomina_path = tk.StringVar()
        self.rinomina_password = tk.StringVar()

        # Organize Tab Vars
        self.organizza_source_dir = tk.StringVar()
        self.organizza_dest_dir = tk.StringVar(value=os.path.join(const.APPLICATION_PATH, const.ORGANIZZA_DEST_DIR))

        # Fees Tab Vars
        self.canoni_selected_year = tk.StringVar()
        self.canoni_selected_month = tk.StringVar()
        self.canoni_messina_num = tk.StringVar()
        self.canoni_naselli_num = tk.StringVar()
        self.canoni_caldarella_num = tk.StringVar()
        self.canoni_word_path = tk.StringVar()
        self.selected_printer = tk.StringVar()
        self.canoni_macro_name = tk.StringVar(value=const.DEFAULT_MACRO_NAME)

        # Internal path holders for Fees tab
        self.canoni_giornaliera_path = tk.StringVar()
        self.canoni_cons1_path = tk.StringVar()
        self.canoni_cons2_path = tk.StringVar()
        self.canoni_cons3_path = tk.StringVar()

    def _load_config_into_vars(self):
        """
        Loads values from the ConfigManager into the tk.StringVars.
        """
        self.firma_ghostscript_path.set(self.config_manager.get("firma_ghostscript_path"))
        self.rinomina_path.set(self.config_manager.get("rinomina_path"))
        self.rinomina_password.set(self.config_manager.get("rinomina_password"))
        self.organizza_source_dir.set(self.config_manager.get("organizza_source_dir"))

        self.canoni_selected_year.set(self.config_manager.get("canoni_selected_year"))
        self.canoni_selected_month.set(self.config_manager.get("canoni_selected_month"))
        self.canoni_messina_num.set(self.config_manager.get("canoni_messina_num"))
        self.canoni_naselli_num.set(self.config_manager.get("canoni_naselli_num"))
        self.canoni_caldarella_num.set(self.config_manager.get("canoni_caldarella_num"))
        self.canoni_word_path.set(self.config_manager.get("canoni_word_path"))
        self.selected_printer.set(self.config_manager.get("selected_printer"))
        self.email_to.set(self.config_manager.get("email_to"))
        self.email_subject.set(self.config_manager.get("email_subject"))

    def _create_widgets(self):
        # --- Create a main container with a scrollbar ---
        main_container = ttk.Frame(self)
        main_container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_container)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")

        # --- Place the notebook inside the scrollable frame ---
        notebook = ttk.Notebook(scrollable_frame)
        notebook.pack(expand=True, fill='both', padx=10, pady=10)

        # Create a container frame for each tab that includes the tab content and the log area
        firma_container = ttk.Frame(notebook)
        rinomina_container = ttk.Frame(notebook)
        organizza_container = ttk.Frame(notebook)
        canoni_container = ttk.Frame(notebook)

        firma_container.pack(fill='both', expand=True)
        rinomina_container.pack(fill='both', expand=True)
        organizza_container.pack(fill='both', expand=True)
        canoni_container.pack(fill='both', expand=True)

        notebook.add(firma_container, text=' Apponi Firma ')
        notebook.add(rinomina_container, text=' Aggiungi Data Schede ')
        notebook.add(organizza_container, text=' Organizza e Stampa Schede ')
        notebook.add(canoni_container, text=' Stampa Canoni Mensili ')

        # --- Create Log Widgets ---
        log_frame_firma = ttk.LabelFrame(firma_container, text="Log Esecuzione (Firma)", padding="10")
        log_frame_firma.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))
        log_widget_firma = create_log_widget(log_frame_firma)

        log_frame_rinomina = ttk.LabelFrame(rinomina_container, text="Log Esecuzione (Aggiungi Data)", padding="10")
        log_frame_rinomina.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))
        log_widget_rinomina = create_log_widget(log_frame_rinomina)

        log_frame_organizza = ttk.LabelFrame(organizza_container, text="Log Esecuzione (Organizza/Stampa)", padding="10")
        log_frame_organizza.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))
        log_widget_organizza = create_log_widget(log_frame_organizza)

        log_frame_canoni = ttk.LabelFrame(canoni_container, text="Log Esecuzione (Stampa Canoni)", padding="10")
        log_frame_canoni.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))
        log_widget_canoni = create_log_widget(log_frame_canoni)

        # --- Create and Pack Tab Content ---
        # Pass the application config (self) and the specific logger to each tab
        signature_tab = SignatureTab(firma_container, self, lambda msg, level='INFO': log_message(log_widget_firma, msg, level))
        signature_tab.pack(fill='both', expand=True, before=log_frame_firma)

        rename_tab = RenameTab(rinomina_container, self, lambda msg, level='INFO': log_message(log_widget_rinomina, msg, level))
        rename_tab.pack(fill='both', expand=True, before=log_frame_rinomina)

        organize_tab = OrganizeTab(organizza_container, self, lambda msg, level='INFO': log_message(log_widget_organizza, msg, level))
        organize_tab.pack(fill='both', expand=True, before=log_frame_organizza)

        fees_tab = FeesTab(canoni_container, self, lambda msg, level='INFO': log_message(log_widget_canoni, msg, level))
        fees_tab.pack(fill='both', expand=True, before=log_frame_canoni)

    def _on_closing(self):
        """
        Saves the current configuration and closes the application.
        """
        current_config = {
            "firma_ghostscript_path": self.firma_ghostscript_path.get(),
            "rinomina_path": self.rinomina_path.get(),
            "rinomina_password": self.rinomina_password.get(),
            "organizza_source_dir": self.organizza_source_dir.get(),
            "canoni_selected_year": self.canoni_selected_year.get(),
            "canoni_selected_month": self.canoni_selected_month.get(),
            "canoni_messina_num": self.canoni_messina_num.get(),
            "canoni_naselli_num": self.canoni_naselli_num.get(),
            "canoni_caldarella_num": self.canoni_caldarella_num.get(),
            "canoni_word_path": self.canoni_word_path.get(),
            "selected_printer": self.selected_printer.get(),
            "email_to": self.email_to.get(),
            "email_subject": self.email_subject.get()
        }
        self.config_manager.save(current_config)
        self.destroy()
