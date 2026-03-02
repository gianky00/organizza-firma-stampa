import os
import re
import threading
import tkinter as tk
from datetime import datetime
from tkinter import ttk

from src.logic.email_handler import EmailHandler
from src.logic.signature import SignatureProcessor
from src.utils.ui_utils import (
    create_path_entry,
    open_folder_in_explorer,
    select_file_dialog,
    select_folder_dialog,
)


class SignatureTab(ttk.Frame):
    def __init__(self, parent, app_config, logger, processor_class=None, excel_gateway_class=None):
        super().__init__(parent)
        self.app_config = app_config
        self.log_widget = logger
        self.prepared_drafts = []
        self.drafts_lock = threading.Lock()
        self.current_draft_index = 0
        self.cancel_event = threading.Event()

        self._create_widgets()

        processor_cls = processor_class or SignatureProcessor
        self.processor = processor_cls(
            self,
            app_config,
            self.setup_progress,
            self.update_progress,
            self.hide_progress,
            excel_gateway_class=excel_gateway_class,
        )

    def _create_widgets(self):
        # Create a canvas and scrollbar for scrolling
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.scroll_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.scroll_window, width=e.width))

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Bind mousewheel to scrolling
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self.scrollable_frame.columnconfigure(0, weight=1)
        container = self.scrollable_frame

        # Update all .pack calls to use 'container' or its children
        # --- Description ---
        desc_text = (
            "Automatizza il processo di firma: apre file Excel, applica una firma, li converte in PDF e li comprime."
        )
        desc_label = ttk.Label(container, text=desc_text, wraplength=800, justify=tk.LEFT, style="info.TLabel")
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor="w")

        # --- Frame Setup ---
        paths_frame = ttk.LabelFrame(container, text="1. Percorsi e Impostazioni", padding=15)
        paths_frame.pack(fill=tk.X, pady=5)
        paths_frame.columnconfigure(0, weight=1)

        mode_frame = ttk.LabelFrame(container, text="2. Tipo di Documento", padding=15)
        mode_frame.pack(fill=tk.X, pady=5)

        self.actions_frame = ttk.LabelFrame(container, text="3. Azioni", padding=15)
        self.actions_frame.pack(fill=tk.X, pady=5)
        self.actions_frame.columnconfigure(0, weight=1)

        self.email_frame = ttk.LabelFrame(container, text="4. Crea Bozza Email con PDF Firmati", padding=15)
        self.email_frame.pack(fill=tk.X, pady=5)
        self.email_frame.columnconfigure(0, weight=1)

        # --- Paths Frame Content ---
        create_path_entry(
            paths_frame,
            "Cartella Excel:",
            self.app_config.firma_excel_dir,
            lambda: select_folder_dialog(self.app_config.firma_excel_dir),
            0,
            readonly=True,
        )
        create_path_entry(
            paths_frame,
            "Cartella PDF di Output:",
            self.app_config.firma_pdf_dir,
            lambda: open_folder_in_explorer(self.app_config.firma_pdf_dir.get()),
            1,
            readonly=True,
            button_text="Apri",
        )
        create_path_entry(
            paths_frame,
            "Immagine Firma:",
            self.app_config.firma_image_path,
            lambda: select_file_dialog(self.app_config.firma_image_path, [("Immagini", "*.png;*.jpg;*.jpeg")]),
            2,
            readonly=True,
        )
        create_path_entry(
            paths_frame,
            "Ghostscript:",
            self.app_config.firma_ghostscript_path,
            lambda: select_file_dialog(self.app_config.firma_ghostscript_path, [("Eseguibile", "*.exe")]),
            3,
            readonly=False,
        )

        # --- Mode Frame Content ---
        ttk.Radiobutton(
            mode_frame,
            text="Schede (Controllo, Manutenzione, etc.)",
            variable=self.app_config.firma_processing_mode,
            value="schede",
        ).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Radiobutton(
            mode_frame,
            text="Preventivi (Basato su foglio 'Consuntivo')",
            variable=self.app_config.firma_processing_mode,
            value="preventivi",
        ).pack(anchor=tk.W, padx=5, pady=2)

        # --- Actions Frame Content ---
        self.run_button = ttk.Button(
            self.actions_frame,
            text="▶  AVVIA PROCESSO FIRMA COMPLETO",
            style="primary.TButton",
            command=self.start_signature_process,
        )
        self.run_button.pack(fill=tk.X, ipady=8, pady=5)
        self.cancel_button = ttk.Button(self.actions_frame, text="Annulla Processo", command=self.cancel_process)
        # self.cancel_button is packed/unpacked dynamically

        # --- Email Frame Content ---
        email_settings_frame = ttk.Frame(self.email_frame)
        email_settings_frame.grid(row=0, column=0, sticky=tk.EW, pady=(0, 10))

        ttk.Label(email_settings_frame, text="Template TCL:", width=25).grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        tcl_options = ["", *list(self.app_config.TCL_CONTACTS.keys())]
        self.tcl_combo = ttk.Combobox(
            email_settings_frame, textvariable=self.app_config.email_tcl, values=tcl_options, state="readonly", width=30
        )
        self.tcl_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))

        self.style_check = ttk.Checkbutton(
            email_settings_frame,
            text="Usa stile Formale",
            variable=self.app_config.email_is_formal,
            onvalue=True,
            offvalue=False,
        )
        self.style_check.grid(row=0, column=2, sticky=tk.W, padx=(0, 10))

        ttk.Label(email_settings_frame, text="Limite MB/Email:", width=15).grid(
            row=0, column=3, sticky=tk.E, padx=(10, 5)
        )
        self.size_limit_entry = ttk.Entry(email_settings_frame, textvariable=self.app_config.email_size_limit, width=8)
        self.size_limit_entry.grid(row=0, column=4, sticky=tk.E)

        create_path_entry(self.email_frame, "Destinatario(i):", self.app_config.email_to, None, 1, readonly=False)
        create_path_entry(self.email_frame, "CC:", self.app_config.email_cc, None, 2, readonly=False)
        create_path_entry(self.email_frame, "Oggetto:", self.app_config.email_subject, None, 3, readonly=False)

        body_frame = ttk.Frame(self.email_frame)
        body_frame.grid(row=4, column=0, sticky="ew", pady=5)
        body_frame.columnconfigure(1, weight=1)
        ttk.Label(body_frame, text="Corpo del Messaggio:", width=25).grid(row=0, column=0, sticky="nw", padx=(0, 5))
        self.email_body_text = tk.Text(body_frame, height=8, font=("Segoe UI", 9), relief=tk.SOLID, borderwidth=1)
        self.email_body_text.grid(row=0, column=1, sticky="ew", padx=5)

        action_preview_frame = ttk.Frame(self.email_frame)
        action_preview_frame.grid(row=5, column=0, sticky=tk.EW, pady=(10, 0))
        self.prepare_button = ttk.Button(action_preview_frame, text="Prepara Bozze", command=self.prepare_email_drafts)
        self.prepare_button.pack(side=tk.LEFT)

        self.preview_frame = ttk.Frame(action_preview_frame)
        # self.preview_frame is packed/unpacked dynamically
        self.prev_button = ttk.Button(self.preview_frame, text="<", command=self.show_prev_draft, width=3)
        self.prev_button.pack(side=tk.LEFT, padx=(10, 0))
        self.preview_label = ttk.Label(self.preview_frame, text="Anteprima 0/0", width=15, anchor="center")
        self.preview_label.pack(side=tk.LEFT)
        self.next_button = ttk.Button(self.preview_frame, text=">", command=self.show_next_draft, width=3)
        self.next_button.pack(side=tk.LEFT)

        self.email_button = ttk.Button(
            action_preview_frame, text="Crea Bozze in Outlook", command=self.start_email_creation_process
        )
        self.email_button.pack(side=tk.RIGHT)

        self.tcl_combo.bind("<<ComboboxSelected>>", self._update_email_preview)
        self.style_check.config(command=self._update_email_preview)
        self.on_process_finished()
        self._update_email_preview()

    def start_signature_process(self):
        self.cancel_event.clear()
        self.toggle_buttons(is_running=True)
        self.preview_frame.pack_forget()
        self.prepared_drafts = []
        threading.Thread(
            target=self.processor.run_full_signature_process, args=(self.cancel_event,), daemon=True
        ).start()

    def cancel_process(self):
        self.log_firma("Annullamento richiesto...", "WARNING")
        self.cancel_event.set()
        self.cancel_button.config(state="disabled")

    def on_process_finished(self):
        self.toggle_buttons(is_running=False)
        pdf_dir = self.app_config.firma_pdf_dir.get()
        has_pdfs = os.path.isdir(pdf_dir) and any(f.lower().endswith(".pdf") for f in os.listdir(pdf_dir))
        self.prepare_button.config(state="normal" if has_pdfs else "disabled")
        self.email_button.config(state="disabled")

    def toggle_buttons(self, is_running):
        if is_running:
            self.run_button.pack_forget()
            self.cancel_button.pack(fill=tk.X, ipady=8, pady=5)
            self.cancel_button.config(state="normal")
            self.prepare_button.config(state="disabled")
            self.email_button.config(state="disabled")
        else:
            self.cancel_button.pack_forget()
            self.run_button.pack(fill=tk.X, ipady=8, pady=5)
            self.run_button.config(state="normal")

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
                self.log_firma(f"ERRORE: Limite di dimensione non valido: '{self.app_config.email_size_limit.get()}'.", "ERROR")
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
                all_attachments.append({
                    "path": full_p,
                    "size": os.path.getsize(full_p),
                    "tcl": files_metadata.get(norm_full_p, "N/D")
                })

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
            
            base_template = self.email_body_text.get("1.0", tk.END).strip()

            for i, chunk in enumerate(chunks):
                if self.cancel_event.is_set(): break
                
                draft = {
                    "to": self.app_config.email_to.get(),
                    "cc": self.app_config.email_cc.get(),
                    "subject": f"[{i + 1}/{num_drafts}] {base_subject}" if num_drafts > 1 else base_subject,
                    "attachments": [item["path"] for item in chunk],
                    "file_list": [
                        {"name": os.path.splitext(os.path.basename(item["path"]))[0], "tcl": item["tcl"]}
                        for item in chunk
                    ],
                    "intro_text": base_template if i == 0 else "Seguito della mail precedente."
                }
                with self.drafts_lock:
                    self.prepared_drafts.append(draft)

            # 7. Finalizzazione
            self.log_firma(f"Preparate {len(self.prepared_drafts)} bozze di email.", "SUCCESS")
            self.current_draft_index = 0
            self._display_draft_preview()
            self.preview_frame.pack(side=tk.LEFT, padx=(20, 0))
            self.email_button.config(state="normal")

        except Exception as e:
            self.log_firma(f"ERRORE IMPREVISTO durante la preparazione bozze: {e}", "ERROR")
            import traceback
            self.log_firma(traceback.format_exc(), "DEBUG")
        self.email_button.config(state="normal")

    def _display_draft_preview(self):
        if not self.prepared_drafts:
            self.preview_frame.pack_forget()
            return
        draft = self.prepared_drafts[self.current_draft_index]
        self.preview_label["text"] = f"Anteprima {self.current_draft_index + 1}/{len(self.prepared_drafts)}"
        self.app_config.email_to.set(draft["to"])
        self.app_config.email_subject.set(draft["subject"])
        file_list_str = "\n".join([os.path.splitext(os.path.basename(p))[0] for p in draft["attachments"]])
        full_body = draft["intro_text"].replace("{file_list}", file_list_str)
        self.email_body_text.delete("1.0", tk.END)
        self.email_body_text.insert("1.0", full_body)
        self.prev_button.config(state="normal" if self.current_draft_index > 0 else "disabled")
        self.next_button.config(
            state="normal" if self.current_draft_index < len(self.prepared_drafts) - 1 else "disabled"
        )

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
            self.master.after(0, self.preview_frame.pack_forget)
        finally:
            self.master.after(0, self.on_process_finished)

    def log_firma(self, message, level="INFO"):
        self.master.after(0, self.log_widget, message, level)

    def setup_progress(self, max_value, label_text="Progresso:"):
        self.app_config.setup_global_progress(max_value, label_text)

    def show_indeterminate(self, label_text="Inizializzazione..."):
        self.app_config.show_global_indeterminate(label_text)

    def update_progress(self, value):
        self.app_config.update_global_progress(value)

    def hide_progress(self):
        self.app_config.hide_global_progress()

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

        self.email_body_text.delete("1.0", tk.END)
        self.email_body_text.insert("1.0", body_template)
