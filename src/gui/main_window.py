import os
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import ttk

from src.gui.tabs.fees_tab import FeesTab
from src.gui.tabs.organize_tab import OrganizeTab
from src.gui.tabs.rename_tab import RenameTab
from src.gui.tabs.settings_tab import SettingsTab
from src.gui.tabs.signature_tab import SignatureTab
from src.utils import constants as const
from src.utils.config_manager import ConfigManager
from src.utils.ui_utils import create_log_widget, log_message


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestione Documenti Ufficio (Refactored)")
        try:
            self.state("zoomed")
        except tk.TclError:
            self.geometry("1200x900")
            self.center_window(1200, 900)
        self.resizable(True, True)

        self.config_manager = ConfigManager()
        self.config_manager.load()

        self._initialize_stringvars()
        self._load_config_into_vars()  # Carica i dati PRIMA di creare i widget
        self._setup_style()
        self._create_widgets()

        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def center_window(self, width, height):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _setup_style(self):
        self.font_main = ("Segoe UI", 10)
        self.font_bold = ("Segoe UI", 11, "bold")
        self.background_color = "#f0f0f0"

        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(".", font=self.font_main, background=self.background_color)
        style.configure("TLabel", font=self.font_main, background=self.background_color)
        style.configure("TLabelframe", background=self.background_color, bordercolor="#cccccc")
        style.configure("TLabelframe.Label", font=self.font_bold, background=self.background_color)
        style.configure("info.TLabel", foreground="#333333", background=self.background_color)

        style.configure("TButton", padding=6, font=self.font_main)
        style.map("TButton", background=[("active", "#e0e0e0")], foreground=[("disabled", "#a0a0a0")])

        style.configure("primary.TButton", background="#0078D4", foreground="white", font=self.font_bold)
        style.map(
            "primary.TButton",
            background=[("active", "#005a9e"), ("disabled", "#a0a0a0")],
            foreground=[("disabled", "#ffffff")],
        )

        style.configure("TNotebook", background=self.background_color, borderwidth=0)
        style.configure("TNotebook.Tab", padding=[12, 6], font=self.font_main)
        style.map(
            "TNotebook.Tab",
            background=[("selected", self.background_color), ("!selected", "#d0d0d0")],
            expand=[("selected", [0, 2, 0, 0])],
        )

    def _initialize_stringvars(self):
        # ... (this method is unchanged)
        self.FIRMA_EXCEL_INPUT_DIR = const.FIRMA_EXCEL_INPUT_DIR
        self.ORGANIZZA_DEST_DIR = const.ORGANIZZA_DEST_DIR
        self.CANONI_GIORNALIERA_BASE_DIR = const.CANONI_GIORNALIERA_BASE_DIR
        self.CANONI_CONSUNTIVI_BASE_DIR = const.CANONI_CONSUNTIVI_BASE_DIR
        self.mesi_giornaliera_map = const.MESI_GIORNALIERA_MAP
        self.nomi_mesi_italiani = const.NOMI_MESI_ITALIANI
        self.TCL_CONTACTS = const.TCL_CONTACTS
        self.EMAIL_TCL_SCHEDE = const.EMAIL_TCL_SCHEDE
        self.EMAIL_BODY_INFORMAL = const.EMAIL_BODY_INFORMAL
        self.EMAIL_BODY_FORMAL = const.EMAIL_BODY_FORMAL
        self.EMAIL_BODY_GENERIC_INFORMAL = const.EMAIL_BODY_GENERIC_INFORMAL
        self.EMAIL_BODY_GENERIC_FORMAL = const.EMAIL_BODY_GENERIC_FORMAL
        self.firma_excel_dir = tk.StringVar(value=os.path.join(const.APPLICATION_PATH, const.FIRMA_EXCEL_INPUT_DIR))
        self.firma_image_path = tk.StringVar(
            value=os.path.join(const.APPLICATION_PATH, "src", "assets", const.FIRMA_IMAGE_NAME)
        )
        self.firma_pdf_dir = tk.StringVar(value=os.path.join(const.APPLICATION_PATH, const.FIRMA_PDF_OUTPUT_DIR))
        self.firma_ghostscript_path = tk.StringVar()
        self.firma_processing_mode = tk.StringVar(value="schede")
        self.email_to = tk.StringVar()
        self.email_cc = tk.StringVar()
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

        # Le variabili dinamiche dei TCL verranno popolate in _load_config_into_vars
        self.canoni_tcl_vars = []

        self.canoni_word_path = tk.StringVar()
        self.selected_printer = tk.StringVar()
        self.canoni_macro_name = tk.StringVar(value=const.DEFAULT_MACRO_NAME)
        self.canoni_giornaliera_path = tk.StringVar()
        self.canoni_cons1_path = tk.StringVar()
        self.canoni_cons2_path = tk.StringVar()
        self.canoni_cons3_path = tk.StringVar()
        self.canoni_cons4_path = tk.StringVar()

        # New dynamic settings
        self.canoni_giornaliera_base_dir = tk.StringVar()
        self.canoni_consuntivi_base_dir = tk.StringVar()
        self.organizza_base_dir = tk.StringVar()

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

        tcl_data = self.config_manager.get("canoni_tcl_list")
        self.canoni_tcl_vars = []
        for ref in tcl_data:
            self.canoni_tcl_vars.append(
                {
                    "name": tk.StringVar(value=ref.get("name", "")),
                    "tcl": tk.StringVar(value=ref.get("tcl", "")),
                    "num": tk.StringVar(value=ref.get("num", "")),
                    "print": tk.BooleanVar(value=ref.get("print", False)),
                    "path": tk.StringVar(value=""),
                }
            )

        self.canoni_word_path.set(self.config_manager.get("canoni_word_path"))
        self.selected_printer.set(self.config_manager.get("selected_printer"))
        self.email_to.set(self.config_manager.get("email_to"))
        self.email_cc.set(self.config_manager.get("email_cc"))
        self.email_subject.set(self.config_manager.get("email_subject"))
        self.email_tcl.set(self.config_manager.get("email_tcl"))
        self.email_is_formal.set(bool(self.config_manager.get("email_is_formal")))
        self.email_size_limit.set(self.config_manager.get("email_size_limit"))
        self.canoni_giornaliera_base_dir.set(self.config_manager.get("canoni_giornaliera_base_dir"))
        self.canoni_consuntivi_base_dir.set(self.config_manager.get("canoni_consuntivi_base_dir"))
        self.organizza_base_dir.set(self.config_manager.get("organizza_base_dir"))

        organize_default_path = os.path.join(
            self.organizza_base_dir.get(), prev_month_year_str, organize_folder_month_str
        )
        self.organizza_source_dir.set(organize_default_path)

    def _create_widgets(self):
        self.configure(background=self.background_color)

        # --- Header con Progress Bar Globale ---
        self.header_frame = ttk.Frame(self, padding=(15, 5))
        self.header_frame.pack(fill=tk.X, side=tk.TOP)
        self.header_frame.columnconfigure(0, weight=1)

        # Info App a sinistra
        app_info_lbl = ttk.Label(self.header_frame, text="GESTIONE DOCUMENTI - SMI", font=("Segoe UI", 9, "bold"), foreground="#666666")
        app_info_lbl.grid(row=0, column=0, sticky="w")

        from src.utils.ui_utils import ProgressWithETA
        self.global_progress_container = ttk.Frame(self.header_frame)
        self.global_progress_container.grid(row=0, column=1, sticky="e")

        self.global_progress = ProgressWithETA(self.global_progress_container)
        # La teniamo inizialmente invisibile tramite il metodo hide_global_progress()
        self.hide_global_progress()

        main_container = ttk.Frame(self, padding="10")
        main_container.pack(fill=tk.BOTH, expand=True)

        # --- Notebook for Tabs ---
        notebook = ttk.Notebook(main_container)
        notebook.pack(expand=True, fill="both")

        # --- Create Tab Containers ---
        self.firma_container = ttk.Frame(notebook, padding="15")
        self.rinomina_container = ttk.Frame(notebook, padding="15")
        self.organizza_container = ttk.Frame(notebook, padding="15")
        self.canoni_container = ttk.Frame(notebook, padding="15")
        self.impostazioni_container = ttk.Frame(notebook, padding="15")

        self.firma_container.columnconfigure(0, weight=1)
        self.rinomina_container.columnconfigure(0, weight=1)
        self.organizza_container.columnconfigure(0, weight=1)
        self.canoni_container.columnconfigure(0, weight=1)
        self.impostazioni_container.columnconfigure(0, weight=1)

        notebook.add(self.rinomina_container, text=" Aggiungi Data Schede ")
        notebook.add(self.firma_container, text=" Apponi Firma ")
        notebook.add(self.organizza_container, text=" Organizza e Stampa Schede ")
        notebook.add(self.canoni_container, text=" Stampa Canoni Mensili ")
        notebook.add(self.impostazioni_container, text=" Impostazioni Avanzate ")

        # --- Create Log Widgets ---
        self.log_widget_firma = self._create_log_frame(self.firma_container, "Log Esecuzione (Firma)")
        self.log_widget_rinomina = self._create_log_frame(self.rinomina_container, "Log Esecuzione (Aggiungi Data)")
        self.log_widget_organizza = self._create_log_frame(
            self.organizza_container, "Log Esecuzione (Organizza/Stampa)"
        )
        self.log_widget_canoni = self._create_log_frame(self.canoni_container, "Log Esecuzione (Stampa Canoni)")

        # --- Dependency Injection and Tab Creation ---
        self.signature_tab = SignatureTab(
            self.firma_container, self, lambda msg, level="INFO": log_message(self.log_widget_firma, msg, level)
        )
        self.signature_tab.pack(fill="both", expand=True)
        self.log_widget_firma.master.pack_forget()  # Hide log frame initially
        self.log_widget_firma.master.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))

        self.rename_tab = RenameTab(
            self.rinomina_container, self, lambda msg, level="INFO": log_message(self.log_widget_rinomina, msg, level)
        )
        self.rename_tab.pack(fill="both", expand=True)
        self.log_widget_rinomina.master.pack_forget()
        self.log_widget_rinomina.master.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))

        self.fees_tab = FeesTab(
            self.canoni_container, self, lambda msg, level="INFO": log_message(self.log_widget_canoni, msg, level)
        )
        self.fees_tab.pack(fill="both", expand=True)
        self.log_widget_canoni.master.pack_forget()
        self.log_widget_canoni.master.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))

        self.organize_tab = OrganizeTab(
            self.organizza_container,
            self,
            lambda msg, level="INFO": log_message(self.log_widget_organizza, msg, level),
            self.fees_tab.processor,
        )
        self.organize_tab.pack(fill="both", expand=True)
        self.log_widget_organizza.master.pack_forget()
        self.log_widget_organizza.master.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))

        self.settings_tab = SettingsTab(self.impostazioni_container, self)
        self.settings_tab.pack(fill="both", expand=True)

    def _create_log_frame(self, parent, title):
        log_frame = ttk.LabelFrame(parent, text=title, padding="10")
        # The frame is packed by the caller
        log_widget = create_log_widget(log_frame)
        return log_widget

    def _on_closing(self):
        # --- On Closing ---
        tcl_to_save = [
            {"name": ref["name"].get(), "tcl": ref["tcl"].get(), "num": ref["num"].get(), "print": ref["print"].get()}
            for ref in self.canoni_tcl_vars
        ]

        current_config = {
            "firma_ghostscript_path": self.firma_ghostscript_path.get(),
            "rinomina_path": self.rinomina_path.get(),
            "rinomina_password": self.rinomina_password.get(),
            "canoni_tcl_list": tcl_to_save,
            "canoni_word_path": self.canoni_word_path.get(),
            "selected_printer": self.selected_printer.get(),
            "email_to": self.email_to.get(),
            "email_cc": self.email_cc.get(),
            "email_subject": self.email_subject.get(),
            "email_tcl": self.email_tcl.get(),
            "email_is_formal": self.email_is_formal.get(),
            "email_size_limit": self.email_size_limit.get(),
            "canoni_giornaliera_base_dir": self.canoni_giornaliera_base_dir.get(),
            "canoni_consuntivi_base_dir": self.canoni_consuntivi_base_dir.get(),
            "organizza_base_dir": self.organizza_base_dir.get(),
        }
        self.config_manager.save(current_config)
        self.destroy()

    # --- Metodi Progress Bar Globale ---
    def setup_global_progress(self, max_value, label_text="Progresso:"):
        self.global_progress.pack(side=tk.RIGHT)
        self.global_progress.setup(max_value, label_text)
        self.header_frame.update_idletasks()

    def show_global_indeterminate(self, label_text="Elaborazione..."):
        self.global_progress.pack(side=tk.RIGHT)
        self.global_progress.setup_indeterminate(label_text)
        self.header_frame.update_idletasks()

    def update_global_progress(self, value):
        self.global_progress.update_progress(value)

    def hide_global_progress(self):
        self.global_progress.stop_indeterminate()
        self.global_progress.pack_forget()
        self.header_frame.update_idletasks()
