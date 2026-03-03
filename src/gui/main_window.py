import os
from datetime import datetime, timedelta

from PySide6.QtCore import QTimer
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from src.gui.tabs.fees_tab import FeesTab
from src.gui.tabs.organize_tab import OrganizeTab
from src.gui.tabs.rename_tab import RenameTab
from src.gui.tabs.settings_tab import SettingsTab
from src.gui.tabs.signature_tab import SignatureTab
from src.utils import constants as const
from src.utils.config_manager import ConfigManager
from src.utils.qt_vars import BooleanVar, StringVar
from src.utils.ui_utils import ProgressWithETA, create_log_widget, log_message


class MainApplication(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gestione Documenti Ufficio (PySide6)")
        self.resize(1200, 900)
        self.center_window(1200, 900)

        self.config_manager = ConfigManager()
        self.config_manager.load()

        self._initialize_stringvars()
        self._load_config_into_vars()  # Carica i dati PRIMA di creare i widget
        self._setup_style()
        self._create_widgets()

    def center_window(self, width, height):
        # We can use QScreen to find center
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        x = (screen_geometry.width() - width) // 2
        y = (screen_geometry.height() - height) // 2
        self.setGeometry(x, y, width, height)
        # Try to maximize
        self.showMaximized()

    def _setup_style(self):
        # We will use modern PySide6 styles / stylesheets
        self.background_color = "#f0f0f0"
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {self.background_color};
            }}
            QTabWidget::pane {{
                border: 1px solid #cccccc;
                background: {self.background_color};
            }}
            QTabBar::tab {{
                background: #d0d0d0;
                padding: 8px 15px;
                margin-right: 2px;
                font-family: 'Segoe UI';
                font-size: 10pt;
            }}
            QTabBar::tab:selected {{
                background: {self.background_color};
                border-bottom-color: {self.background_color};
                font-weight: bold;
            }}
            QGroupBox {{
                font-weight: bold;
                border: 1px solid #cccccc;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 15px;
                font-family: 'Segoe UI';
                font-size: 11pt;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 3px;
            }}
            QLabel {{
                font-family: 'Segoe UI';
                font-size: 10pt;
            }}
            QPushButton {{
                padding: 6px;
                font-family: 'Segoe UI';
                font-size: 10pt;
            }}
        """)

    def _initialize_stringvars(self):
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

        self.firma_excel_dir = StringVar(
            value=os.path.join(const.APPLICATION_PATH, const.FIRMA_EXCEL_INPUT_DIR), parent=self
        )
        self.firma_image_path = StringVar(
            value=os.path.join(const.APPLICATION_PATH, "src", "assets", const.FIRMA_IMAGE_NAME), parent=self
        )
        self.firma_pdf_dir = StringVar(
            value=os.path.join(const.APPLICATION_PATH, const.FIRMA_PDF_OUTPUT_DIR), parent=self
        )
        self.firma_ghostscript_path = StringVar(parent=self)
        self.firma_processing_mode = StringVar(value="schede", parent=self)
        self.email_to = StringVar(parent=self)
        self.email_cc = StringVar(parent=self)
        self.email_subject = StringVar(parent=self)
        self.email_tcl = StringVar(parent=self)
        self.email_is_formal = BooleanVar(value=False, parent=self)
        self.email_size_limit = StringVar(value="6", parent=self)
        self.rinomina_path = StringVar(parent=self)
        self.rinomina_password = StringVar(parent=self)
        self.organizza_source_dir = StringVar(parent=self)
        self.organizza_dest_dir = StringVar(
            value=os.path.join(const.APPLICATION_PATH, const.ORGANIZZA_DEST_DIR), parent=self
        )
        self.canoni_selected_year = StringVar(parent=self)
        self.canoni_selected_month = StringVar(parent=self)

        # Le variabili dinamiche dei TCL verranno popolate in _load_config_into_vars
        self.canoni_tcl_vars = []
        self.rename_models_vars = []

        self.canoni_word_path = StringVar(parent=self)
        self.selected_printer = StringVar(parent=self)
        self.canoni_macro_name = StringVar(value=const.DEFAULT_MACRO_NAME, parent=self)
        self.canoni_giornaliera_path = StringVar(parent=self)
        self.canoni_cons1_path = StringVar(parent=self)
        self.canoni_cons2_path = StringVar(parent=self)
        self.canoni_cons3_path = StringVar(parent=self)
        self.canoni_cons4_path = StringVar(parent=self)

        # New dynamic settings
        self.canoni_giornaliera_base_dir = StringVar(parent=self)
        self.canoni_consuntivi_base_dir = StringVar(parent=self)
        self.organizza_base_dir = StringVar(parent=self)

    def _load_config_into_vars(self):
        self.firma_excel_dir.set(self.config_manager.get("firma_excel_dir"))
        self.firma_pdf_dir.set(self.config_manager.get("firma_pdf_dir"))
        self.firma_image_path.set(self.config_manager.get("firma_image_path"))
        self.firma_ghostscript_path.set(self.config_manager.get("firma_ghostscript_path"))
        self.firma_processing_mode.set(self.config_manager.get("firma_processing_mode"))
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
                    "name": StringVar(value=ref.get("name", ""), parent=self),
                    "tcl": StringVar(value=ref.get("tcl", ""), parent=self),
                    "num": StringVar(value=ref.get("num", ""), parent=self),
                    "print": BooleanVar(value=ref.get("print", False), parent=self),
                    "path": StringVar(value="", parent=self),
                }
            )

        models_data = self.config_manager.get("rename_models_config")
        self.rename_models_vars = []

        # Carica modelli di default per fallback/healing
        from src.domain.models import RENAME_MODELS

        defaults_map = {m.match_value: m for m in RENAME_MODELS}

        for mod in models_data:
            match_val = mod.get("match_value", "")

            raw_id_cells = mod.get("id_cells", [])
            if not raw_id_cells and mod.get("id_cell"):
                raw_id_cells = [mod.get("id_cell")]

            if not raw_id_cells and match_val in defaults_map:
                raw_id_cells = defaults_map[match_val].id_cells

            raw_tcl_cells = mod.get("tcl_cells", [])
            if not raw_tcl_cells and mod.get("tcl_cell"):
                raw_tcl_cells = [mod.get("tcl_cell")]

            if not raw_tcl_cells and match_val in defaults_map:
                raw_tcl_cells = defaults_map[match_val].tcl_cells

            print_area = mod.get("print_area")
            if not print_area and match_val in defaults_map:
                print_area = defaults_map[match_val].print_area
            if not print_area:
                print_area = "A1:N50"

            self.rename_models_vars.append(
                {
                    "name": StringVar(value=mod.get("name", ""), parent=self),
                    "match_value": StringVar(value=match_val, parent=self),
                    "id_cells": StringVar(value=", ".join(raw_id_cells), parent=self),
                    "tcl_cell": StringVar(value=raw_tcl_cells[0] if raw_tcl_cells else "", parent=self),
                    "date_cells": StringVar(value=", ".join(mod.get("date_cells", [])), parent=self),
                    "print_area": StringVar(value=print_area, parent=self),
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
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # --- Header con Progress Bar Globale ---
        self.header_frame = QWidget()
        header_layout = QHBoxLayout(self.header_frame)
        header_layout.setContentsMargins(0, 0, 0, 0)

        # Info App a sinistra
        app_info_lbl = QLabel("GESTIONE DOCUMENTI - SMI")
        font = QFont("Segoe UI", 10, QFont.Weight.Bold)
        app_info_lbl.setFont(font)
        app_info_lbl.setStyleSheet("color: #666666;")
        header_layout.addWidget(app_info_lbl)

        header_layout.addStretch()

        self.global_progress = ProgressWithETA(self.header_frame)
        self.global_progress.hide()
        header_layout.addWidget(self.global_progress)

        main_layout.addWidget(self.header_frame)

        # --- Tab Widget ---
        self.notebook = QTabWidget()
        main_layout.addWidget(self.notebook)

        # --- Create Tab Containers ---
        self.rinomina_container = QWidget()
        self.firma_container = QWidget()
        self.organizza_container = QWidget()
        self.canoni_container = QWidget()
        self.impostazioni_container = QWidget()

        self.notebook.addTab(self.rinomina_container, " Aggiungi Data Schede ")
        self.notebook.addTab(self.firma_container, " Apponi Firma ")
        self.notebook.addTab(self.organizza_container, " Organizza e Stampa Schede ")
        self.notebook.addTab(self.canoni_container, " Stampa Canoni Mensili ")
        self.notebook.addTab(self.impostazioni_container, " Impostazioni Avanzate ")

        # --- Initialize Layouts for Tabs ---
        self.rinomina_layout = QVBoxLayout(self.rinomina_container)
        self.firma_layout = QVBoxLayout(self.firma_container)
        self.organizza_layout = QVBoxLayout(self.organizza_container)
        self.canoni_layout = QVBoxLayout(self.canoni_container)
        self.impostazioni_layout = QVBoxLayout(self.impostazioni_container)

        # --- Create Log Widgets ---
        self.log_widget_rinomina, self.log_frame_rinomina = self._create_log_frame(
            self.rinomina_container, "Log Esecuzione (Aggiungi Data)"
        )
        self.log_widget_firma, self.log_frame_firma = self._create_log_frame(
            self.firma_container, "Log Esecuzione (Firma)"
        )
        self.log_widget_organizza, self.log_frame_organizza = self._create_log_frame(
            self.organizza_container, "Log Esecuzione (Organizza/Stampa)"
        )
        self.log_widget_canoni, self.log_frame_canoni = self._create_log_frame(
            self.canoni_container, "Log Esecuzione (Stampa Canoni)"
        )

        # Instantiate Tabs (They will add themselves or provide their widgets to add)
        # Note: In PySide6, we typically pass the parent and the layout, or the Tab inherits from QWidget and we just add it to the layout.
        # For simplicity during migration, we'll instantiate them and add them to the VBoxLayouts.

        self.rename_tab = RenameTab(self, lambda msg, level="INFO": log_message(self.log_widget_rinomina, msg, level))
        self.rinomina_layout.addWidget(self.rename_tab, 1)  # stretch=1
        self.rinomina_layout.addWidget(self.log_frame_rinomina)

        self.signature_tab = SignatureTab(
            self, lambda msg, level="INFO": log_message(self.log_widget_firma, msg, level)
        )
        self.firma_layout.addWidget(self.signature_tab, 1)
        self.firma_layout.addWidget(self.log_frame_firma)

        self.fees_tab = FeesTab(self, lambda msg, level="INFO": log_message(self.log_widget_canoni, msg, level))
        self.canoni_layout.addWidget(self.fees_tab, 1)
        self.canoni_layout.addWidget(self.log_frame_canoni)

        self.organize_tab = OrganizeTab(
            self, lambda msg, level="INFO": log_message(self.log_widget_organizza, msg, level), self.fees_tab.processor
        )
        self.organizza_layout.addWidget(self.organize_tab, 1)
        self.organizza_layout.addWidget(self.log_frame_organizza)

        self.settings_tab = SettingsTab(self)
        self.impostazioni_layout.addWidget(self.settings_tab, 1)

    def _create_log_frame(self, parent, title):
        log_frame = QGroupBox(title)
        layout = QVBoxLayout(log_frame)
        log_widget = create_log_widget(log_frame)
        layout.addWidget(log_widget)
        # Prevent it from expanding too much
        log_frame.setMaximumHeight(200)
        return log_widget, log_frame

    def closeEvent(self, event):  # noqa: N802
        # --- On Closing ---
        tcl_to_save = [
            {"name": ref["name"].get(), "tcl": ref["tcl"].get(), "num": ref["num"].get(), "print": ref["print"].get()}
            for ref in self.canoni_tcl_vars
        ]

        models_to_save = []
        for mod in self.rename_models_vars:
            id_list = [d.strip() for d in mod["id_cells"].get().split(",") if d.strip()]
            tcl_val = mod["tcl_cell"].get().strip()
            models_to_save.append(
                {
                    "name": mod["name"].get(),
                    "match_value": mod["match_value"].get(),
                    "id_cells": id_list,
                    "id_cell": id_list[0] if id_list else "",  # Legacy support
                    "tcl_cells": [tcl_val] if tcl_val else [],
                    "tcl_cell": tcl_val,  # Legacy support
                    "date_cells": [d.strip() for d in mod["date_cells"].get().split(",") if d.strip()],
                    "print_area": mod["print_area"].get(),
                }
            )

        current_config = {
            "firma_ghostscript_path": self.firma_ghostscript_path.get(),
            "rinomina_path": self.rinomina_path.get(),
            "rinomina_password": self.rinomina_password.get(),
            "canoni_tcl_list": tcl_to_save,
            "rename_models_config": models_to_save,
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
        event.accept()

    # --- Metodi Progress Bar Globale ---
    def setup_global_progress(self, max_value, label_text="Progresso:"):
        def _setup():
            self.global_progress.show()
            self.global_progress.setup(max_value, label_text)

        QTimer.singleShot(0, _setup)

    def show_global_indeterminate(self, label_text="Elaborazione..."):
        def _show():
            self.global_progress.show()
            self.global_progress.setup_indeterminate(label_text)

        QTimer.singleShot(0, _show)

    def update_global_progress(self, value):
        def _update():
            self.global_progress.update_progress(value)

        QTimer.singleShot(0, _update)

    def hide_global_progress(self):
        def _hide():
            self.global_progress.stop_indeterminate()
            self.global_progress.hide()

        QTimer.singleShot(0, _hide)
