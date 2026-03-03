import os
import re
import threading
from datetime import datetime

from PySide6.QtCore import Qt, Signal, QTimer, Signal
from PySide6.QtWidgets import (
    QFrame,
    QCheckBox,
    QComboBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.logic.email_handler import EmailHandler
from src.logic.signature import SignatureProcessor
from src.utils.ui_utils import (
    create_path_entry,
    open_folder_in_explorer,
    select_file_dialog,
    select_folder_dialog,
)


class SignatureTab(QWidget):
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
        self.prepared_drafts = []
        self.drafts_lock = threading.Lock()
        self.current_draft_index = 0
        self.cancel_event = threading.Event()
        self.log_signal.connect(self._handle_log)
        self.progress_setup_signal.connect(self._handle_setup_progress)
        self.progress_update_signal.connect(self._handle_update_progress)
        self.progress_hide_signal.connect(self._handle_hide_progress)
        self.indeterminate_signal.connect(self._handle_indeterminate)
        self.process_finished_signal.connect(self.on_process_finished)

        self._create_widgets()

        processor_cls = processor_class or SignatureProcessor
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
        desc_text = (
            "Automatizza il processo di firma: apre file Excel, applica una firma, li converte in PDF e li comprime."
        )
        desc_label = QLabel(desc_text)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #333333;")
        container.addWidget(desc_label)

        # --- Frame Setup ---
        paths_frame = QGroupBox("1. Percorsi e Impostazioni")
        paths_layout = QVBoxLayout(paths_frame)
        container.addWidget(paths_frame)

        mode_frame = QGroupBox("2. Tipo di Documento")
        mode_layout = QVBoxLayout(mode_frame)
        container.addWidget(mode_frame)

        self.actions_frame = QGroupBox("3. Azioni")
        actions_layout = QVBoxLayout(self.actions_frame)
        container.addWidget(self.actions_frame)

        self.email_frame = QGroupBox("4. Crea Bozza Email con PDF Firmati")
        email_layout = QVBoxLayout(self.email_frame)
        container.addWidget(self.email_frame)

        # --- Paths Frame Content ---
        p1 = create_path_entry(
            paths_frame,
            "Cartella Excel:",
            self.app_config.firma_excel_dir,
            lambda: select_folder_dialog(self.app_config.firma_excel_dir, self),
            readonly=True,
        )
        paths_layout.addLayout(p1)

        p2 = create_path_entry(
            paths_frame,
            "Cartella PDF di Output:",
            self.app_config.firma_pdf_dir,
            lambda: open_folder_in_explorer(self.app_config.firma_pdf_dir.get()),
            readonly=True,
            button_text="Apri",
        )
        paths_layout.addLayout(p2)

        p3 = create_path_entry(
            paths_frame,
            "Immagine Firma:",
            self.app_config.firma_image_path,
            lambda: select_file_dialog(self.app_config.firma_image_path, "Immagini (*.png *.jpg *.jpeg)", self),
            readonly=True,
        )
        paths_layout.addLayout(p3)

        p4 = create_path_entry(
            paths_frame,
            "Ghostscript:",
            self.app_config.firma_ghostscript_path,
            lambda: select_file_dialog(self.app_config.firma_ghostscript_path, "Eseguibile (*.exe)", self),
            readonly=False,
        )
        paths_layout.addLayout(p4)

        # --- Mode Frame Content ---
        self.rb_schede = QRadioButton("Schede (Controllo, Manutenzione, etc.)")
        self.rb_preventivi = QRadioButton("Preventivi (Basato su foglio 'Consuntivo')")

        # Set initial state based on variable
        if self.app_config.firma_processing_mode.get() == "preventivi":
            self.rb_preventivi.setChecked(True)
        else:
            self.rb_schede.setChecked(True)

        def on_mode_changed():
            if self.rb_schede.isChecked():
                self.app_config.firma_processing_mode.set("schede")
            else:
                self.app_config.firma_processing_mode.set("preventivi")

        self.rb_schede.toggled.connect(on_mode_changed)
        self.rb_preventivi.toggled.connect(on_mode_changed)

        mode_layout.addWidget(self.rb_schede)
        mode_layout.addWidget(self.rb_preventivi)

        # --- Actions Frame Content ---
        self.run_button = QPushButton("▶  AVVIA PROCESSO FIRMA COMPLETO")
        self.run_button.setStyleSheet("background-color: #0078D4; color: white; font-weight: bold; padding: 10px;")
        self.run_button.clicked.connect(self.start_signature_process)
        actions_layout.addWidget(self.run_button)

        self.cancel_button = QPushButton("Annulla Processo")
        self.cancel_button.setStyleSheet("padding: 10px;")
        self.cancel_button.clicked.connect(self.cancel_process)
        self.cancel_button.hide()
        actions_layout.addWidget(self.cancel_button)

        # --- Email Frame Content ---
        email_settings_layout = QHBoxLayout()
        email_settings_layout.setContentsMargins(0, 0, 0, 0)
        email_layout.addLayout(email_settings_layout)

        email_settings_layout.addWidget(QLabel("Template TCL:"))

        tcl_options = ["", *list(self.app_config.TCL_CONTACTS.keys())]
        self.tcl_combo = QComboBox()
        self.tcl_combo.addItems(tcl_options)
        self.tcl_combo.setMinimumWidth(150)
        self.tcl_combo.setCurrentText(self.app_config.email_tcl.get())

        def on_tcl_changed(text):
            self.app_config.email_tcl.set(text)
            self._update_email_preview()

        self.tcl_combo.currentTextChanged.connect(on_tcl_changed)
        email_settings_layout.addWidget(self.tcl_combo)

        self.style_check = QCheckBox("Usa stile Formale")
        self.style_check.setChecked(self.app_config.email_is_formal.get())

        def on_style_changed(state):
            self.app_config.email_is_formal.set(bool(state))
            self._update_email_preview()

        self.style_check.stateChanged.connect(on_style_changed)
        email_settings_layout.addWidget(self.style_check)

        email_settings_layout.addWidget(QLabel("Limite MB/Email:"))
        self.size_limit_entry = QLineEdit()
        self.size_limit_entry.setMaximumWidth(60)
        self.size_limit_entry.setText(self.app_config.email_size_limit.get())
        self.size_limit_entry.textChanged.connect(self.app_config.email_size_limit.set)
        email_settings_layout.addWidget(self.size_limit_entry)

        email_settings_layout.addStretch()

        e1 = create_path_entry(self.email_frame, "Destinatario(i):", self.app_config.email_to, readonly=False)
        email_layout.addLayout(e1)
        e2 = create_path_entry(self.email_frame, "CC:", self.app_config.email_cc, readonly=False)
        email_layout.addLayout(e2)
        e3 = create_path_entry(self.email_frame, "Oggetto:", self.app_config.email_subject, readonly=False)
        email_layout.addLayout(e3)

        body_layout = QHBoxLayout()
        body_layout.setContentsMargins(0, 0, 0, 0)
        lbl = QLabel("Corpo del Messaggio:")
        lbl.setMinimumWidth(150)
        lbl.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        body_layout.addWidget(lbl)

        self.email_body_text = QTextEdit()
        self.email_body_text.setMinimumHeight(100)
        self.email_body_text.setMaximumHeight(150)
        body_layout.addWidget(self.email_body_text, 1)
        email_layout.addLayout(body_layout)

        action_preview_layout = QHBoxLayout()
        action_preview_layout.setContentsMargins(0, 0, 0, 0)
        email_layout.addLayout(action_preview_layout)

        self.prepare_button = QPushButton("Prepara Bozze")
        self.prepare_button.clicked.connect(self.prepare_email_drafts)
        action_preview_layout.addWidget(self.prepare_button)

        self.preview_widget = QWidget()
        self.preview_layout = QHBoxLayout(self.preview_widget)
        self.preview_layout.setContentsMargins(10, 0, 0, 0)

        self.prev_button = QPushButton("<")
        self.prev_button.setFixedWidth(30)
        self.prev_button.clicked.connect(self.show_prev_draft)
        self.preview_layout.addWidget(self.prev_button)

        self.preview_label = QLabel("Anteprima 0/0")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setMinimumWidth(100)
        self.preview_layout.addWidget(self.preview_label)

        self.next_button = QPushButton(">")
        self.next_button.setFixedWidth(30)
        self.next_button.clicked.connect(self.show_next_draft)
        self.preview_layout.addWidget(self.next_button)

        self.preview_widget.hide()
        action_preview_layout.addWidget(self.preview_widget)
        action_preview_layout.addStretch()

        self.email_button = QPushButton("Crea Bozze in Outlook")
        self.email_button.clicked.connect(self.start_email_creation_process)
        action_preview_layout.addWidget(self.email_button)

        self.on_process_finished()
        self._update_email_preview()

    def start_signature_process(self):
        self.cancel_event.clear()
        self.toggle_buttons(is_running=True)
        self.preview_widget.hide()
        self.prepared_drafts = []
        def _wrapper():
            try:
                self.processor.run_full_signature_process(self.cancel_event)
            finally:
                self.process_finished_signal.emit()
        threading.Thread(target=_wrapper, daemon=True).start()

    def cancel_process(self):
        self.log_firma("Annullamento richiesto...", "WARNING")
        self.cancel_event.set()
        self.cancel_button.setEnabled(False)

    def on_process_finished(self):
        self.toggle_buttons(is_running=False)
        pdf_dir = self.app_config.firma_pdf_dir.get()
        has_pdfs = os.path.isdir(pdf_dir) and any(f.lower().endswith(".pdf") for f in os.listdir(pdf_dir))
        self.prepare_button.setEnabled(has_pdfs)
        self.email_button.setEnabled(False)

    def toggle_buttons(self, is_running):
        if is_running:
            self.run_button.hide()
            self.cancel_button.show()
            self.cancel_button.setEnabled(True)
            self.prepare_button.setEnabled(False)
            self.email_button.setEnabled(False)
        else:
            self.cancel_button.hide()
            self.run_button.show()
            self.run_button.setEnabled(True)

    def prepare_email_drafts(self):
        self.log_firma("Preparazione delle bozze email...", "HEADER")
        try:
            # 1. Validazione Limite Dimensione
            try:
                limit_mb_str = self.app_config.email_size_limit.get()
                limit_mb = float(limit_mb_str)
                if limit_mb <= 0:
                    raise ValueError("Limite <= 0")
                limit_bytes = limit_mb * 1024 * 1024
            except (ValueError, TypeError):
                self.log_firma(
                    f"ERRORE: Limite di dimensione non valido: '{self.app_config.email_size_limit.get()}'.", "ERROR"
                )
                return

            # 2. Controllo Cartella PDF
            pdf_dir = self.app_config.firma_pdf_dir.get()
            if not os.path.isdir(pdf_dir):
                self.log_firma(f"ERRORE: La cartella PDF non esiste: {pdf_dir}", "ERROR")
                return

            # 3. Recupero Metadati TCL (Normalizzando i percorsi per il mapping)
            files_metadata = {}
            if hasattr(self.processor, "prepared_files_data"):
                for item in self.processor.prepared_files_data:
                    norm_p = os.path.normpath(item["path"]).lower()
                    files_metadata[norm_p] = item["tcl"]

            # 4. Raccolta Allegati
            all_attachments = []
            pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith(".pdf")]

            if not pdf_files:
                self.log_firma("Nessun file PDF trovato nella cartella di output.", "WARNING")
                return

            for f in pdf_files:
                full_p = os.path.join(pdf_dir, f)
                norm_full_p = os.path.normpath(full_p).lower()
                all_attachments.append(
                    {"path": full_p, "size": os.path.getsize(full_p), "tcl": files_metadata.get(norm_full_p, "N/D")}
                )

            # 5. Suddivisione in Chunk (Limite MB)
            chunks: list[list[dict]] = []
            current_chunk: list[dict] = []
            current_chunk_size = 0

            for item in all_attachments:
                if current_chunk and current_chunk_size + item["size"] > limit_bytes:
                    chunks.append(current_chunk)
                    current_chunk = []
                    current_chunk_size = 0
                current_chunk.append(item)
                current_chunk_size += item["size"]
            if current_chunk:
                chunks.append(current_chunk)

            # 6. Creazione Bozze Interne
            self.prepared_drafts = []
            num_drafts = len(chunks)
            raw_subject = self.app_config.email_subject.get()
            base_subject = re.sub(r"^\[\d+/\d+\]\s*", "", raw_subject)

            base_template = self.email_body_text.toPlainText().strip()
            
            # Helper per calcolare le email dinamiche dai TCL ("PASSANISI D." -> "dpassanisi@isab.com")
            def get_dynamic_emails(chunk_items):
                emails = []
                tcls = set(item["tcl"] for item in chunk_items)
                for t in tcls:
                    if not t or t == "N/D":
                        continue
                    parts = t.replace('.', '').strip().lower().split()
                    if len(parts) >= 2:
                        # Prende prima lettera del secondo nome/cognome + primo nome/cognome
                        email = f"{parts[1][0]}{parts[0]}@isab.com"
                        emails.append(email)
                return "; ".join(emails)

            is_schede_mode = (self.app_config.email_tcl.get() == "Schede" or not self.app_config.email_tcl.get())

            for i, chunk in enumerate(chunks):
                if self.cancel_event.is_set():
                    break
                    
                draft_to = self.app_config.email_to.get()
                if is_schede_mode:
                    dynamic_to = get_dynamic_emails(chunk)
                    if dynamic_to:
                        draft_to = dynamic_to

                draft = {
                    "to": draft_to,
                    "cc": self.app_config.email_cc.get(),
                    "subject": f"[{i + 1}/{num_drafts}] {base_subject}" if num_drafts > 1 else base_subject,
                    "attachments": [item["path"] for item in chunk],
                    "file_list": [
                        {"name": os.path.splitext(os.path.basename(item["path"]))[0], "tcl": item["tcl"]}
                        for item in chunk
                    ],
                    "intro_text": base_template if i == 0 else "Seguito della mail precedente.",
                }
                with self.drafts_lock:
                    self.prepared_drafts.append(draft)

            # 7. Finalizzazione
            self.log_firma(f"Preparate {len(self.prepared_drafts)} bozze di email.", "SUCCESS")
            self.log_firma("Controlla l'anteprima in alto e premi 'Crea Bozze in Outlook' per aprire le email.", "INFO")
            self.current_draft_index = 0
            self._display_draft_preview()
            self.preview_widget.show()
            self.email_button.setEnabled(True)
            self.email_button.setStyleSheet("background-color: #28a745; color: white; font-weight: bold; padding: 8px;")

        except Exception as e:
            self.log_firma(f"ERRORE IMPREVISTO durante la preparazione bozze: {e}", "ERROR")
            import traceback

            self.log_firma(traceback.format_exc(), "DEBUG")
            self.email_button.setEnabled(False)

    def _display_draft_preview(self):
        if not self.prepared_drafts:
            self.preview_widget.hide()
            return
        draft = self.prepared_drafts[self.current_draft_index]
        self.preview_label.setText(f"Anteprima {self.current_draft_index + 1}/{len(self.prepared_drafts)}")
        self.app_config.email_to.set(draft["to"])
        self.app_config.email_subject.set(draft["subject"])
        file_list_str = "\n".join([os.path.splitext(os.path.basename(p))[0] for p in draft["attachments"]])
        full_body = draft["intro_text"].replace("{file_list}", file_list_str)
        self.email_body_text.clear()
        self.email_body_text.insertPlainText(full_body)
        self.prev_button.setEnabled(self.current_draft_index > 0)
        self.next_button.setEnabled(self.current_draft_index < len(self.prepared_drafts) - 1)

    def show_prev_draft(self):
        with self.drafts_lock:
            if self.current_draft_index > 0:
                self.current_draft_index -= 1
                self._display_draft_preview()

    def show_next_draft(self):
        with self.drafts_lock:
            if self.current_draft_index < len(self.prepared_drafts) - 1:
                self.current_draft_index += 1
                self._display_draft_preview()

    def start_email_creation_process(self):
        self.toggle_buttons(is_running=True)
        threading.Thread(target=self.create_email_drafts_in_outlook, daemon=True).start()

    def create_email_drafts_in_outlook(self):
        try:
            with self.drafts_lock:
                drafts_copy = self.prepared_drafts.copy()

            if not drafts_copy:
                self.log_firma("Nessuna bozza da creare.", "WARNING")
                return
            self.log_firma(f"Avvio creazione di {len(drafts_copy)} bozze in Outlook...", "HEADER")
            email_handler = EmailHandler(self.log_firma)
            for draft_info in drafts_copy:
                if self.cancel_event.is_set():
                    self.log_firma("Creazione bozze annullata.", "WARNING")
                    break
                email_handler.create_outlook_draft(draft_info)
            self.log_firma("Creazione bozze in Outlook completata.", "SUCCESS")
            with self.drafts_lock:
                self.prepared_drafts = []
            QTimer.singleShot(0, self.preview_widget.hide)
        finally:
            self.process_finished_signal.emit()

    def log_firma(self, message, level="INFO"):
        self.log_signal.emit(str(message), str(level))

    def setup_progress(self, max_value, label_text="Progresso:"):
        self.progress_setup_signal.emit(float(max_value), str(label_text))

    def show_indeterminate(self, label_text="Inizializzazione..."):
        self.indeterminate_signal.emit(str(label_text))

    def update_progress(self, value):
        self.progress_update_signal.emit(float(value))

    def hide_progress(self):
        self.progress_hide_signal.emit()

    def _get_date_range_from_filenames(self):
        pdf_dir = self.app_config.firma_pdf_dir.get()
        if not os.path.isdir(pdf_dir):
            return None, None

        date_pattern = re.compile(r"(\d{2}-\d{2}-\d{4})")
        dates = []
        for filename in os.listdir(pdf_dir):
            if filename.lower().endswith(".pdf"):
                match = date_pattern.search(filename)
                if match:
                    try:
                        date_obj = datetime.strptime(match.group(1), "%d-%m-%Y")
                        dates.append(date_obj)
                    except ValueError:
                        continue  # Ignore invalid date formats

        if not dates:
            return None, None

        min_date = min(dates).strftime("%d/%m/%Y")
        max_date = max(dates).strftime("%d/%m/%Y")
        return min_date, max_date

    def _update_email_preview(self, event=None):
        tcl_name = self.app_config.email_tcl.get()
        is_formal = self.app_config.email_is_formal.get()
        body_template = ""
        subject_template = ""

        # Clear fields before populating
        self.app_config.email_to.set("")
        self.app_config.email_cc.set("")
        self.app_config.email_subject.set("")

        if tcl_name == "Schede":
            template_data = self.app_config.EMAIL_TCL_SCHEDE
            self.app_config.email_to.set(template_data["to"])
            self.app_config.email_cc.set(template_data["cc"])

            data_inizio, data_fine = self._get_date_range_from_filenames()
            if data_inizio and data_fine:
                subject_template = template_data["subject"].format(data_inizio=data_inizio, data_fine=data_fine)
                body_template = (
                    template_data["body"].replace("{data_inizio}", data_inizio).replace("{data_fine}", data_fine)
                )
            else:
                # Fallback if no dates found
                subject_template = (
                    template_data["subject"].replace("{data_inizio}", "GG/MM/AAAA").replace("{data_fine}", "GG/MM/AAAA")
                )
                body_template = (
                    template_data["body"].replace("{data_inizio}", "GG/MM/AAAA").replace("{data_fine}", "GG/MM/AAAA")
                )

            self.app_config.email_subject.set(subject_template)

        elif tcl_name and tcl_name in self.app_config.TCL_CONTACTS:
            email = self.app_config.TCL_CONTACTS[tcl_name]
            self.app_config.email_to.set(email)
            first_name = tcl_name.split()[0]
            if is_formal:
                body_template = self.app_config.EMAIL_BODY_FORMAL.format(name=first_name, file_list="{file_list}")
            else:
                body_template = self.app_config.EMAIL_BODY_INFORMAL.format(name=first_name, file_list="{file_list}")
        else:  # Generic or empty selection
            if is_formal:
                body_template = self.app_config.EMAIL_BODY_GENERIC_FORMAL.format(file_list="{file_list}")
            else:
                body_template = self.app_config.EMAIL_BODY_GENERIC_INFORMAL.format(file_list="{file_list}")

        self.email_body_text.clear()
        self.email_body_text.insertPlainText(body_template)

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
