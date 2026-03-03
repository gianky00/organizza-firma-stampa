import threading

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QGroupBox, QLabel, QPushButton, QVBoxLayout, QWidget

from src.logic.renaming import RenameProcessor
from src.utils.ui_utils import create_path_entry, select_folder_dialog


class RenameTab(QWidget):
    log_signal = Signal(str, str)
    progress_setup_signal = Signal(float, str)
    progress_update_signal = Signal(float)
    progress_hide_signal = Signal()
    indeterminate_signal = Signal(str)
    process_finished_signal = Signal()

    def __init__(self, parent, logger, processor_class=None, excel_gateway_class=None):
        super().__init__(parent)
        self.app_config = parent  # MainApplication
        self.log_widget = logger
        self.cancel_event = threading.Event()

        # Connect signals to slots
        self.log_signal.connect(self._handle_log)
        self.progress_setup_signal.connect(self._handle_setup_progress)
        self.progress_update_signal.connect(self._handle_update_progress)
        self.progress_hide_signal.connect(self._handle_hide_progress)
        self.indeterminate_signal.connect(self._handle_indeterminate)
        self.process_finished_signal.connect(self.on_process_finished)

        self._create_widgets()
        processor_cls = processor_class or RenameProcessor
        self.processor = processor_cls(
            self,
            self.app_config,
            self.setup_progress,
            self.update_progress,
            self.hide_progress,
            excel_gateway_class=excel_gateway_class,
        )

    def _create_widgets(self):
        main_layout = QVBoxLayout(self)

        # --- Description ---
        desc_text = "Analizza i file Excel in una cartella, trova la data di emissione e li rinomina nel formato NOME (GG-MM-AAAA). Prova ad usare una password per i file protetti."
        desc_label = QLabel(desc_text)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #333333;")
        main_layout.addWidget(desc_label)

        # --- Settings Frame ---
        settings_frame = QGroupBox("1. Impostazioni")
        settings_layout = QVBoxLayout(settings_frame)

        path_layout = create_path_entry(
            settings_frame,
            "Cartella da Analizzare:",
            self.app_config.rinomina_path,
            lambda: select_folder_dialog(self.app_config.rinomina_path, self),
            readonly=False,
        )
        settings_layout.addLayout(path_layout)

        pwd_layout = create_path_entry(
            settings_frame,
            "Password (opzionale):",
            self.app_config.rinomina_password,
            None,
            readonly=False,
        )
        settings_layout.addLayout(pwd_layout)

        main_layout.addWidget(settings_frame)

        # --- Actions Frame ---
        self.actions_frame = QGroupBox("2. Azioni")
        actions_layout = QVBoxLayout(self.actions_frame)

        self.run_button = QPushButton("▶ AVVIA PROCESSO DI RINOMINA")
        self.run_button.setStyleSheet("background-color: #0078D4; color: white; font-weight: bold; padding: 10px;")
        self.run_button.clicked.connect(self.start_rename_process)
        actions_layout.addWidget(self.run_button)

        self.cancel_button = QPushButton("Annulla Processo")
        self.cancel_button.setStyleSheet("padding: 10px;")
        self.cancel_button.clicked.connect(self.cancel_process)
        self.cancel_button.hide()
        actions_layout.addWidget(self.cancel_button)

        main_layout.addWidget(self.actions_frame)
        main_layout.addStretch()

        self.toggle_buttons(is_running=False)

    def start_rename_process(self):
        self.cancel_event.clear()
        self.toggle_buttons(is_running=True)
        threading.Thread(target=self._run_and_finish, daemon=True).start()

    def _run_and_finish(self):
        try:
            self.processor.run_rename_process(self.cancel_event)
        finally:
            self.process_finished_signal.emit()

    def cancel_process(self):
        self.log_rinomina("Annullamento richiesto...", "WARNING")
        self.cancel_event.set()
        self.cancel_button.setEnabled(False)

    def on_process_finished(self):
        self.toggle_buttons(is_running=False)

    def toggle_buttons(self, is_running):
        if is_running:
            self.run_button.hide()
            self.cancel_button.show()
            self.cancel_button.setEnabled(True)
        else:
            self.cancel_button.hide()
            self.run_button.show()
            self.run_button.setEnabled(True)

    # API chiamata dai thread secondari
    def log_rinomina(self, message, level="INFO"):
        self.log_signal.emit(str(message), str(level))

    def setup_progress(self, max_value, label_text="Progresso:"):
        self.progress_setup_signal.emit(float(max_value), str(label_text))

    def show_indeterminate(self, label_text="Inizializzazione..."):
        self.indeterminate_signal.emit(str(label_text))

    def update_progress(self, value):
        self.progress_update_signal.emit(float(value))

    def hide_progress(self):
        self.progress_hide_signal.emit()

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
