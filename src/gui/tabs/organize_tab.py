import os
import threading

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtWidgets import QCheckBox, QGroupBox, QHBoxLayout, QLabel, QPushButton, QScrollArea, QVBoxLayout, QWidget

from src.logic.organization import OrganizationProcessor
from src.utils.ui_utils import create_path_entry, open_folder_in_explorer, select_folder_dialog


class OrganizeTab(QWidget):
    log_signal = Signal(str, str)
    progress_setup_signal = Signal(float, str)
    progress_update_signal = Signal(float)
    progress_hide_signal = Signal()
    indeterminate_signal = Signal(str)
    process_finished_signal = Signal()

    def __init__(self, parent, logger, fees_processor, processor_class=None, excel_gateway_class=None):
        super().__init__(parent)
        self.app_config = parent  # MainApplication is passed as parent which acts as app_config
        self.log_widget = logger
        self.stampa_checkboxes = {}
        self.cancel_event = threading.Event()
        self.log_signal.connect(self._handle_log)
        self.progress_setup_signal.connect(self._handle_setup_progress)
        self.progress_update_signal.connect(self._handle_update_progress)
        self.progress_hide_signal.connect(self._handle_hide_progress)
        self.indeterminate_signal.connect(self._handle_indeterminate)
        self.process_finished_signal.connect(self.on_process_finished)
        self.active_process_type = None

        self._create_widgets()
        processor_cls = processor_class or OrganizationProcessor
        self.processor = processor_cls(
            self,
            self.app_config,
            fees_processor,
            self.setup_progress,
            self.update_progress,
            self.hide_progress,
            excel_gateway_class=excel_gateway_class,
        )
        QTimer.singleShot(100, self.populate_stampa_list)

    def _create_widgets(self):
        main_layout = QVBoxLayout(self)

        # --- Description ---
        desc_label = QLabel(
            "Analizza i file Excel da una cartella, li organizza in sottocartelle per ODC, e poi permette la stampa di gruppo."
        )
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #333333;")
        main_layout.addWidget(desc_label)

        # --- Organization Frame ---
        self.org_frame = QGroupBox("1. Organizza File per ODC")
        org_layout = QVBoxLayout(self.org_frame)

        path_layout = create_path_entry(
            self.org_frame,
            "Cartella di Origine:",
            self.app_config.organizza_source_dir,
            lambda: select_folder_dialog(self.app_config.organizza_source_dir, self),
            readonly=False,
        )
        org_layout.addLayout(path_layout)

        self.organize_button = QPushButton("🚀 Avvia Organizzazione")
        self.organize_button.setStyleSheet("background-color: #0078D4; color: white; font-weight: bold;")
        self.organize_button.clicked.connect(self.start_organization_process)
        org_layout.addWidget(self.organize_button)

        self.cancel_org_button = QPushButton("Annulla Organizzazione")
        self.cancel_org_button.clicked.connect(self.cancel_process)
        self.cancel_org_button.hide()
        org_layout.addWidget(self.cancel_org_button)

        main_layout.addWidget(self.org_frame)

        # --- Printing Frame ---
        self.print_frame = QGroupBox("2. Stampa Schede Organizzate")
        print_layout = QVBoxLayout(self.print_frame)

        # --- Print Controls ---
        self.print_controls_layout = QHBoxLayout()

        self.print_button = QPushButton("🖨️ Stampa Selezionate")
        self.print_button.setStyleSheet("background-color: #0078D4; color: white; font-weight: bold;")
        self.print_button.clicked.connect(self.start_printing_process)
        self.print_controls_layout.addWidget(self.print_button)

        self.refresh_button = QPushButton("🔄 Aggiorna")
        self.refresh_button.clicked.connect(self.populate_stampa_list)
        self.print_controls_layout.addWidget(self.refresh_button)

        self.open_folder_button = QPushButton("📂 Apri Cartella")
        self.open_folder_button.clicked.connect(
            lambda: open_folder_in_explorer(self.app_config.organizza_dest_dir.get())
        )
        self.print_controls_layout.addWidget(self.open_folder_button)

        self.cancel_print_button = QPushButton("Annulla Stampa")
        self.cancel_print_button.clicked.connect(self.cancel_process)
        self.cancel_print_button.hide()
        self.print_controls_layout.addWidget(self.cancel_print_button)

        print_layout.addLayout(self.print_controls_layout)

        # --- Checkbox List (Scroll Area) ---
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.stampa_checkbox_widget = QWidget()
        self.stampa_checkbox_layout = QVBoxLayout(self.stampa_checkbox_widget)
        self.stampa_checkbox_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scroll_area.setWidget(self.stampa_checkbox_widget)

        print_layout.addWidget(self.scroll_area)
        main_layout.addWidget(self.print_frame)

        self.on_process_finished()

    def start_process(self, process_type, target_func, *args):
        self.cancel_event.clear()
        self.active_process_type = process_type
        self.toggle_buttons(is_running=True)

        thread_args = (self.cancel_event, *args)

        def _wrapper():
            try:
                target_func(*thread_args)
            finally:
                self.process_finished_signal.emit()

        threading.Thread(target=_wrapper, daemon=True).start()

    def start_organization_process(self):
        self.start_process("organize", self.processor.run_organization_process)

    def start_printing_process(self):
        selected_folders = [d["path"] for d in self.stampa_checkboxes.values() if d["checkbox"].isChecked()]
        self.start_process("print", self.processor.run_printing_process, selected_folders)

    def cancel_process(self):
        self.log_organizza("Annullamento richiesto...", "WARNING")
        self.cancel_event.set()
        self.cancel_org_button.setEnabled(False)
        self.cancel_print_button.setEnabled(False)

    def on_process_finished(self):
        self.toggle_buttons(is_running=False)
        self.active_process_type = None

    def toggle_buttons(self, is_running):
        self.organize_button.setEnabled(not is_running)
        self.print_button.setEnabled(not is_running)
        self.refresh_button.setEnabled(not is_running)

        if is_running:
            if self.active_process_type == "organize":
                self.organize_button.hide()
                self.cancel_org_button.show()
                self.cancel_org_button.setEnabled(True)
            elif self.active_process_type == "print":
                self.print_button.hide()
                self.cancel_print_button.show()
                self.cancel_print_button.setEnabled(True)
        else:
            self.cancel_org_button.hide()
            self.cancel_print_button.hide()
            self.organize_button.show()
            self.print_button.show()

    def log_organizza(self, message, level="INFO"):
        self.log_signal.emit(str(message), str(level))

    def setup_progress(self, max_value, label_text="Progresso:"):
        self.progress_setup_signal.emit(float(max_value), str(label_text))

    def update_progress(self, value):
        self.progress_update_signal.emit(float(value))

    def hide_progress(self):
        self.progress_hide_signal.emit()

    def populate_stampa_list(self):
        # Clear existing checkboxes
        for i in reversed(range(self.stampa_checkbox_layout.count())):
            item = self.stampa_checkbox_layout.itemAt(i)
            if item is None:
                continue
            widget_to_remove = item.widget()
            if widget_to_remove is not None:
                widget_to_remove.setParent(None)  # type: ignore
                widget_to_remove.deleteLater()  # type: ignore

        self.stampa_checkboxes.clear()

        year = self.app_config.canoni_selected_year.get()
        month = self.app_config.canoni_selected_month.get()
        odc_map = self.processor.get_odc_to_canone_map(year, month)
        dest_path = self.app_config.organizza_dest_dir.get()

        if not os.path.isdir(dest_path):
            return

        try:
            folders = sorted([d for d in os.listdir(dest_path) if os.path.isdir(os.path.join(dest_path, d))])
            for folder_name in folders:
                folder_path = os.path.join(dest_path, folder_name)
                file_count = 0
                try:
                    file_count = len(
                        [name for name in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, name))]
                    )
                except Exception as e:
                    self.log_organizza(f"Impossibile contare i file nella cartella '{folder_name}': {e}", "WARNING")

                display_text = folder_name
                if folder_name in odc_map:
                    display_text = f"{folder_name} ({odc_map[folder_name]})"
                display_text = f"{display_text} - qt. {file_count}"

                cb = QCheckBox(display_text)
                self.stampa_checkbox_layout.addWidget(cb)
                self.stampa_checkboxes[folder_name] = {"checkbox": cb, "path": folder_path}
        except Exception as e:
            self.log_organizza(f"Errore durante la lettura delle cartelle organizzate: {e}", "ERROR")

    # Slot eseguiti nel thread principale
    def _handle_log(self, message, level):
        self.log_widget(message, level)

    def _handle_setup_progress(self, max_value, label_text):
        self.app_config.setup_global_progress(max_value, label_text)

    def _handle_update_progress(self, value):
        self.app_config.update_global_progress(value)

    def _handle_hide_progress(self):
        self.app_config.hide_global_progress()

    def _handle_indeterminate(self, label_text):
        self.app_config.show_global_indeterminate(label_text)
