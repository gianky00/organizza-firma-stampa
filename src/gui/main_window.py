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
        self.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 40)), "units")

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
        main_container = ttk.Frame(self)
        main_container.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(main_container)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=self.canvas.yview)
        scrollable_frame = ttk.Frame(self.canvas)
        scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))
        self.canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")

        notebook = ttk.Notebook(scrollable_frame)
        notebook.pack(expand=True, fill='both', padx=10, pady=10)

        firma_container = ttk.Frame(notebook)
        rinomina_container = ttk.Frame(notebook)
        organizza_container = ttk.Frame(notebook)
        canoni_container = ttk.Frame(notebook)

        notebook.add(firma_container, text=' Apponi Firma ')
        notebook.add(rinomina_container, text=' Aggiungi Data Schede ')
        notebook.add(organizza_container, text=' Organizza e Stampa Schede ')
        notebook.add(canoni_container, text=' Stampa Canoni Mensili ')

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

        # --- Create and Pack Tab Content with correct dependency injection ---

        # 1. Create Signature and Rename tabs (no dependencies)
        signature_tab = SignatureTab(firma_container, self, lambda msg, level='INFO': log_message(log_widget_firma, msg, level))
        signature_tab.pack(fill='both', expand=True, before=log_frame_firma)

        rename_tab = RenameTab(rinomina_container, self, lambda msg, level='INFO': log_message(log_widget_rinomina, msg, level))
        rename_tab.pack(fill='both', expand=True, before=log_frame_rinomina)

        # 2. Create FeesTab, which creates its own processor
        fees_tab = FeesTab(canoni_container, self, lambda msg, level='INFO': log_message(log_widget_canoni, msg, level))
        fees_tab.pack(fill='both', expand=True, before=log_frame_canoni)

        # 3. Get the processor from FeesTab and pass it to OrganizeTab
        fees_processor = fees_tab.processor
        organize_tab = OrganizeTab(organizza_container, self, lambda msg, level='INFO': log_message(log_widget_organizza, msg, level), fees_processor)
        organize_tab.pack(fill='both', expand=True, before=log_frame_organizza)

        # Pack containers into the notebook
        firma_container.pack(fill='both', expand=True)
        rinomina_container.pack(fill='both', expand=True)
        organizza_container.pack(fill='both', expand=True)
        canoni_container.pack(fill='both', expand=True)

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
