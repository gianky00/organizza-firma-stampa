import threading
from datetime import datetime

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from src.logic.monthly_fees import MonthlyFeesProcessor
from src.utils.ui_utils import create_path_entry, select_file_dialog


class FeesTab(QWidget):
    log_signal = Signal(str, str)
    progress_setup_signal = Signal(float, str)
    progress_update_signal = Signal(float)
    progress_hide_signal = Signal()
    indeterminate_signal = Signal(str)
    process_finished_signal = Signal()

    def __init__(self, parent, logger, processor_class=None, excel_gateway_class=None, word_gateway_class=None):
        super().__init__(parent)
        self.app_config = parent  # MainApplication
        self.log_widget = logger
        self.cancel_event = threading.Event()
        self.log_signal.connect(self._handle_log)
        self.progress_setup_signal.connect(self._handle_setup_progress)
        self.progress_update_signal.connect(self._handle_update_progress)
        self.progress_hide_signal.connect(self._handle_hide_progress)
        self.indeterminate_signal.connect(self._handle_indeterminate)
        self.process_finished_signal.connect(self.on_process_finished)
        processor_cls = processor_class or MonthlyFeesProcessor
        self.processor = processor_cls(
            self,
            self.app_config,
            excel_gateway_class=excel_gateway_class,
            word_gateway_class=word_gateway_class,
        )
        current_year = datetime.now().year
        self.anni_giornaliera = [str(y) for y in range(current_year - 5, current_year + 6)]
        self._create_widgets()
        QTimer.singleShot(100, self.populate_printers)
        QTimer.singleShot(150, self._update_paths_from_ui)

    def _create_widgets(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Create a scroll area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        self.scrollable_widget = QWidget()
        self.scrollable_layout = QVBoxLayout(self.scrollable_widget)
        self.scrollable_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scroll_area.setWidget(self.scrollable_widget)

        main_layout.addWidget(self.scroll_area)

        container = self.scrollable_layout

        # --- Description ---
        desc_label = QLabel(
            "Automatizza la stampa dei canoni mensili eseguendo macro VBA su file Excel e stampando documenti Word in sequenza."
        )
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #333333;")
        container.addWidget(desc_label)

        # --- Settings Frame ---
        settings_frame = QGroupBox("1. Impostazioni di Stampa")
        settings_layout = QVBoxLayout(settings_frame)
        container.addWidget(settings_frame)

        # --- Periodo ---
        period_layout = QHBoxLayout()
        period_layout.setContentsMargins(0, 0, 0, 0)

        lbl_periodo = QLabel("Periodo:")
        lbl_periodo.setStyleSheet("font-weight: bold;")
        period_layout.addWidget(lbl_periodo)

        period_layout.addWidget(QLabel("Anno:"))
        self.anno_combo = QComboBox()
        self.anno_combo.addItems(self.anni_giornaliera)
        self.anno_combo.setCurrentText(self.app_config.canoni_selected_year.get())

        def on_anno_changed(text):
            self.app_config.canoni_selected_year.set(text)
            self._update_paths_from_ui()

        self.anno_combo.currentTextChanged.connect(on_anno_changed)
        period_layout.addWidget(self.anno_combo)

        period_layout.addWidget(QLabel("Mese:"))
        self.mese_combo = QComboBox()
        self.mese_combo.addItems(self.app_config.nomi_mesi_italiani)
        self.mese_combo.setCurrentText(self.app_config.canoni_selected_month.get())

        def on_mese_changed(text):
            self.app_config.canoni_selected_month.set(text)
            self._update_paths_from_ui()

        self.mese_combo.currentTextChanged.connect(on_mese_changed)
        period_layout.addWidget(self.mese_combo)

        period_layout.addStretch()
        settings_layout.addLayout(period_layout)

        # --- Numeri Consuntivo ---
        self.consuntivi_frame = QGroupBox("Numeri Canoni Mensili")
        consuntivi_layout = QVBoxLayout(self.consuntivi_frame)
        settings_layout.addWidget(self.consuntivi_frame)

        # Container per la tabella
        self.table_widget = QWidget()
        self.table_layout = QGridLayout(self.table_widget)
        consuntivi_layout.addWidget(self.table_widget)

        # Bottoni di Gestione
        mgmt_layout = QHBoxLayout()
        self.find_numbers_button = QPushButton("🔍 Trova Numeri Automaticamente")
        self.find_numbers_button.clicked.connect(self.find_numbers_and_populate)
        mgmt_layout.addWidget(self.find_numbers_button)
        mgmt_layout.addStretch()
        consuntivi_layout.addLayout(mgmt_layout)

        self._refresh_dynamic_tcl_ui()

        # --- Altri Percorsi ---
        paths_frame = QGroupBox("Percorsi File")
        paths_layout = QVBoxLayout(paths_frame)
        settings_layout.addWidget(paths_frame)

        p1 = create_path_entry(
            paths_frame,
            "File Giornaliera (Auto):",
            self.app_config.canoni_giornaliera_path,
            None,
            readonly=True,
        )
        paths_layout.addLayout(p1)

        p2 = create_path_entry(
            paths_frame,
            "File Foglio Canone (Word):",
            self.app_config.canoni_word_path,
            lambda: select_file_dialog(
                self.app_config.canoni_word_path, "File Word (*.docx *.doc);;Tutti i file (*.*)", self
            ),
            readonly=False,
        )
        paths_layout.addLayout(p2)

        # --- Stampante e Macro ---
        printer_macro_frame = QGroupBox("Dispositivo e Macro")
        printer_macro_layout = QVBoxLayout(printer_macro_frame)
        settings_layout.addWidget(printer_macro_frame)

        p_layout = QHBoxLayout()
        p_layout.setContentsMargins(0, 0, 0, 0)
        lbl_printer = QLabel("Stampante:")
        lbl_printer.setMinimumWidth(150)
        p_layout.addWidget(lbl_printer)

        self.printer_combo = QComboBox()
        self.printer_combo.currentTextChanged.connect(self.app_config.selected_printer.set)
        p_layout.addWidget(self.printer_combo, 1)
        printer_macro_layout.addLayout(p_layout)

        m_layout = create_path_entry(
            printer_macro_frame,
            "Nome Macro VBA:",
            self.app_config.canoni_macro_name,
            None,
            readonly=True,
        )
        printer_macro_layout.addLayout(m_layout)

        # --- Azioni ---
        self.actions_frame = QGroupBox("2. Azione")
        actions_layout = QVBoxLayout(self.actions_frame)
        container.addWidget(self.actions_frame)

        self.run_button = QPushButton("▶ AVVIA PROCESSO STAMPA CANONI")
        self.run_button.setStyleSheet("background-color: #0078D4; color: white; font-weight: bold; padding: 10px;")
        self.run_button.clicked.connect(self.start_printing_process)
        actions_layout.addWidget(self.run_button)

        self.cancel_button = QPushButton("Annulla Processo")
        self.cancel_button.setStyleSheet("padding: 10px;")
        self.cancel_button.clicked.connect(self.cancel_process)
        self.cancel_button.hide()
        actions_layout.addWidget(self.cancel_button)

        self._setup_tcl_traces()
        self.on_process_finished()

    def _setup_tcl_traces(self):
        for ref in self.app_config.canoni_tcl_vars:
            ref["num"].value_changed.connect(self._update_paths_from_ui)
            ref["name"].value_changed.connect(lambda val: self.log_canoni("Nominativo aggiornato", "DEBUG"))

    def _refresh_dynamic_tcl_ui(self):
        # Pulisce tutto il container della tabella
        for i in reversed(range(self.table_layout.count())):
            item = self.table_layout.itemAt(i)
            if item is None:
                continue
            widget_to_remove = item.widget()
            if widget_to_remove is not None:
                widget_to_remove.setParent(None)  # type: ignore
                widget_to_remove.deleteLater()  # type: ignore

        # Header
        lbl1 = QLabel("Stampa")
        lbl1.setStyleSheet("font-weight: bold;")
        lbl1.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.table_layout.addWidget(lbl1, 0, 0)

        lbl2 = QLabel("Nome TCL")
        lbl2.setStyleSheet("font-weight: bold;")
        self.table_layout.addWidget(lbl2, 0, 1)

        lbl3 = QLabel("N° Canone")
        lbl3.setStyleSheet("font-weight: bold;")
        self.table_layout.addWidget(lbl3, 0, 2)

        self.table_layout.setColumnStretch(1, 1)  # Nome espande

        for i, ref in enumerate(self.app_config.canoni_tcl_vars):
            row_idx = i + 1

            cb = QCheckBox()
            cb.setChecked(ref["print"].get())

            # create local scoped function for signal connection
            def make_cb_handler(r):
                return lambda state: r["print"].set(bool(state))

            cb.stateChanged.connect(make_cb_handler(ref))

            # update cb when var changes externally
            def make_cb_updater(c):
                return lambda val: c.setChecked(bool(val))

            ref["print"].value_changed.connect(make_cb_updater(cb))
            self.table_layout.addWidget(cb, row_idx, 0, Qt.AlignmentFlag.AlignCenter)

            name_lbl = QLabel()
            name_lbl.setText(ref["name"].get())

            def make_name_updater(nl):
                return lambda val: nl.setText(str(val))

            ref["name"].value_changed.connect(make_name_updater(name_lbl))
            self.table_layout.addWidget(name_lbl, row_idx, 1)

            num_ent = QLineEdit()
            num_ent.setText(ref["num"].get())

            num_ent.textChanged.connect(ref["num"].set)

            def make_num_updater(ne):
                return lambda val: ne.setText(str(val)) if ne.text() != str(val) else None

            ref["num"].value_changed.connect(make_num_updater(num_ent))
            self.table_layout.addWidget(num_ent, row_idx, 2)

    def populate_printers(self):
        printers, default_printer = self.processor.get_printers()
        self.printer_combo.clear()
        self.printer_combo.addItems(printers)
        saved_printer = self.app_config.selected_printer.get()
        if saved_printer and saved_printer in printers:
            self.printer_combo.setCurrentText(saved_printer)
        elif default_printer in printers:
            self.printer_combo.setCurrentText(default_printer)
        elif printers:
            self.printer_combo.setCurrentText(printers[0])

    def _update_paths_from_ui(self, *args):
        year = self.app_config.canoni_selected_year.get()
        month = self.app_config.canoni_selected_month.get()
        giornaliera_path = self.processor.get_giornaliera_path(year, month)
        self.app_config.canoni_giornaliera_path.set(giornaliera_path)

        for ref in self.app_config.canoni_tcl_vars:
            p = self.processor.get_consuntivo_path(year, ref["num"].get())
            ref["path"].set(p)

    def start_printing_process(self):
        self.cancel_event.clear()
        self.toggle_buttons(is_running=True)
        self.show_progress()

        consuntivi_data = [
            {"path": ref["path"].get(), "print": ref["print"].get(), "name": ref["name"].get()}
            for ref in self.app_config.canoni_tcl_vars
        ]

        paths_to_print = {
            "giornaliera": self.app_config.canoni_giornaliera_path.get(),
            "consuntivi": consuntivi_data,
            "word": self.app_config.canoni_word_path.get(),
        }
        printer = self.app_config.selected_printer.get()
        macro = self.app_config.canoni_macro_name.get()

        def _wrapper():
            try:
                self.processor.run_printing_process(self.cancel_event, paths_to_print, printer, macro)
            finally:
                self.process_finished_signal.emit()

        threading.Thread(target=_wrapper, daemon=True).start()

    def find_numbers_and_populate(self):
        self.cancel_event.clear()
        self.toggle_buttons(is_running=True)
        self.log_canoni("Ricerca automatica dei numeri di canone in corso...", "HEADER")
        threading.Thread(target=self._find_numbers_thread, args=(self.cancel_event,), daemon=True).start()

    def _find_numbers_thread(self, cancel_event):
        try:
            year = self.app_config.canoni_selected_year.get()
            month = self.app_config.canoni_selected_month.get()

            tcls_to_find = {}
            for ref in self.app_config.canoni_tcl_vars:
                tcl_key = ref["tcl"].get().upper()
                if tcl_key:
                    tcls_to_find[tcl_key] = ref["num"]

            for tcl, var in tcls_to_find.items():
                if cancel_event.is_set():
                    self.log_canoni("Ricerca annullata.", "WARNING")
                    break
                number, _ = self.processor.find_consuntivo_for_tcl(year, month, tcl, cancel_event)
                if number:
                    QTimer.singleShot(0, lambda v=var, num=number: v.set(num))
        finally:
            self.process_finished_signal.emit()

    def cancel_process(self):
        self.log_canoni("Annullamento richiesto...", "WARNING")
        self.cancel_event.set()
        self.cancel_button.setEnabled(False)

    def on_process_finished(self):
        self.toggle_buttons(is_running=False)
        self.hide_progress()

    def toggle_buttons(self, is_running):
        self.run_button.setEnabled(not is_running)
        self.find_numbers_button.setEnabled(not is_running)
        if is_running:
            self.run_button.hide()
            self.cancel_button.show()
            self.cancel_button.setEnabled(True)
        else:
            self.cancel_button.hide()
            self.run_button.show()

    def setup_progress(self, max_value, label_text="Progresso:"):
        self.progress_setup_signal.emit(float(max_value), str(label_text))

    def show_indeterminate(self, label_text="Ricerca in corso..."):
        self.indeterminate_signal.emit(str(label_text))

    def update_progress(self, value):
        self.progress_update_signal.emit(float(value))

    def show_progress(self):
        self.show_indeterminate("Ricerca in corso...")

    def hide_progress(self):
        self.progress_hide_signal.emit()

    def log_canoni(self, message, level="INFO"):
        self.log_signal.emit(str(message), str(level))

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
