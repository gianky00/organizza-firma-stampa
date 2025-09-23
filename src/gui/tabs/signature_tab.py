import tkinter as tk
from tkinter import ttk
import threading
import os
import math
import re
from src.logic.signature import SignatureProcessor
from src.logic.email_handler import EmailHandler
from src.utils.ui_utils import create_path_entry, select_file_dialog, open_folder_in_explorer

class SignatureTab(ttk.Frame):
    def __init__(self, parent, app_config, logger):
        super().__init__(parent)
        self.app_config = app_config
        self.log_widget = logger
        self.prepared_drafts = []
        self.current_draft_index = 0
        self.cancel_event = threading.Event()

        self._create_widgets()

        self.processor = SignatureProcessor(
            self,
            app_config,
            self.setup_progress,
            self.update_progress,
            self.hide_progress
        )

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        desc_text = "Questa sezione automatizza il processo di firma dei documenti. Prende i file Excel dalla cartella 'FILE EXCEL DA FIRMARE', applica la firma 'TIMBRO.png' in base al tipo di documento selezionato, li converte in PDF nella cartella 'PDF' e infine li comprime."
        desc_label = ttk.Label(main_frame, text=desc_text, wraplength=850, justify=tk.LEFT, style='info.TLabel')
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor='w')

        paths_frame = ttk.LabelFrame(main_frame, text="1. Percorsi (Firma)", padding="15")
        paths_frame.pack(fill=tk.X, pady=(0, 10))
        paths_frame.columnconfigure(1, weight=1)
        create_path_entry(paths_frame, "Cartella Excel:", self.app_config.firma_excel_dir, 0, readonly=True)
        ttk.Label(paths_frame, text="Cartella PDF di Output:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        pdf_entry = ttk.Entry(paths_frame, textvariable=self.app_config.firma_pdf_dir, state='readonly')
        pdf_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        open_button = ttk.Button(paths_frame, text="Apri Cartella", command=lambda: open_folder_in_explorer(self.app_config.firma_pdf_dir.get()))
        open_button.grid(row=1, column=2, sticky=tk.E, padx=(5, 0), pady=5)
        create_path_entry(paths_frame, "Immagine Firma:", self.app_config.firma_image_path, 2, readonly=True)
        create_path_entry(paths_frame, "Ghostscript:", self.app_config.firma_ghostscript_path, 3, readonly=False,
                          browse_command=lambda: select_file_dialog(self.app_config.firma_ghostscript_path, "Seleziona eseguibile Ghostscript", [("Executable", "*.exe")]))

        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=15, padx=5)

        mode_frame = ttk.LabelFrame(main_frame, text="2. Tipo di Documento da Firmare", padding="15")
        mode_frame.pack(fill=tk.X, pady=10)
        ttk.Radiobutton(mode_frame, text="Schede (Controllo, Manutenzione, etc.)", variable=self.app_config.firma_processing_mode, value="schede").pack(anchor=tk.W, padx=5, pady=5)
        ttk.Radiobutton(mode_frame, text="Preventivi (Basato su foglio 'Consuntivo')", variable=self.app_config.firma_processing_mode, value="preventivi").pack(anchor=tk.W, padx=5, pady=5)

        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=15, padx=5)

        actions_frame = ttk.LabelFrame(main_frame, text="3. Azioni Firma", padding="15")
        actions_frame.pack(fill=tk.X, pady=10)
        self.run_button = ttk.Button(actions_frame, text="▶  AVVIA PROCESSO FIRMA COMPLETO", style='primary.TButton', command=self.start_signature_process)
        self.run_button.pack(fill=tk.X, ipady=8, pady=5)
        self.cancel_button = ttk.Button(actions_frame, text="Annulla Processo", command=self.cancel_process)
        self.cancel_button.pack(fill=tk.X, ipady=8, pady=5)

        self.email_frame = ttk.LabelFrame(main_frame, text="4. Crea Bozza Email con PDF Firmati", padding="15")
        self.email_frame.pack(fill=tk.X, pady=10)
        self.email_frame.columnconfigure(1, weight=1)
        settings_row_frame = ttk.Frame(self.email_frame)
        settings_row_frame.grid(row=0, column=0, columnspan=2, sticky=tk.EW, pady=(0, 10))
        ttk.Label(settings_row_frame, text="Template TCL:").pack(side=tk.LEFT, padx=(5, 5))
        tcl_options = [""] + list(self.app_config.TCL_CONTACTS.keys())
        self.tcl_combo = ttk.Combobox(settings_row_frame, textvariable=self.app_config.email_tcl, values=tcl_options, state="readonly", width=25)
        self.tcl_combo.pack(side=tk.LEFT, padx=(0, 15))
        self.style_check = ttk.Checkbutton(settings_row_frame, text="Usa stile Formale", variable=self.app_config.email_is_formal, onvalue=True, offvalue=False)
        self.style_check.pack(side=tk.LEFT, padx=(0, 15))
        ttk.Label(settings_row_frame, text="Limite MB/Email:").pack(side=tk.LEFT, padx=(5, 5))
        self.size_limit_entry = ttk.Entry(settings_row_frame, textvariable=self.app_config.email_size_limit, width=5)
        self.size_limit_entry.pack(side=tk.LEFT)
        ttk.Label(self.email_frame, text="Destinatario(i):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.email_to_entry = ttk.Entry(self.email_frame, textvariable=self.app_config.email_to)
        self.email_to_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        ttk.Label(self.email_frame, text="Oggetto:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.email_subject_entry = ttk.Entry(self.email_frame, textvariable=self.app_config.email_subject)
        self.email_subject_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        ttk.Label(self.email_frame, text="Corpo del Messaggio:").grid(row=3, column=0, sticky=tk.NW, padx=5, pady=5)
        self.email_body_text = tk.Text(self.email_frame, height=8, font=("Segoe UI", 9))
        self.email_body_text.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=2)
        action_preview_frame = ttk.Frame(self.email_frame)
        action_preview_frame.grid(row=4, column=1, sticky=tk.EW, pady=(10, 0))
        self.prepare_button = ttk.Button(action_preview_frame, text="Prepara Bozze", command=self.prepare_email_drafts)
        self.prepare_button.pack(side=tk.LEFT)
        self.preview_frame = ttk.Frame(action_preview_frame)
        self.prev_button = ttk.Button(self.preview_frame, text="<", command=self.show_prev_draft, width=3)
        self.prev_button.pack(side=tk.LEFT, padx=(10, 0))
        self.preview_label = ttk.Label(self.preview_frame, text="Anteprima 0/0", width=15, anchor='center')
        self.preview_label.pack(side=tk.LEFT)
        self.next_button = ttk.Button(self.preview_frame, text=">", command=self.show_next_draft, width=3)
        self.next_button.pack(side=tk.LEFT)
        self.email_button = ttk.Button(action_preview_frame, text="Crea Bozze in Outlook", command=self.start_email_creation_process)
        self.email_button.pack(side=tk.RIGHT)
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_label = ttk.Label(self.progress_frame, text="Progresso:")
        self.progress_label.pack(side=tk.LEFT, padx=(5, 5))
        self.progressbar = ttk.Progressbar(self.progress_frame, orient='horizontal', mode='determinate')
        self.progressbar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.percent_label = ttk.Label(self.progress_frame, text="0%", width=5)
        self.percent_label.pack(side=tk.LEFT)

        self.tcl_combo.bind("<<ComboboxSelected>>", self._update_email_preview)
        self.style_check.config(command=self._update_email_preview)
        self.on_process_finished()
        self._update_email_preview()

    def start_signature_process(self):
        self.cancel_event.clear()
        self.toggle_buttons(is_running=True)
        self.preview_frame.pack_forget()
        self.prepared_drafts = []
        threading.Thread(target=self.processor.run_full_signature_process, args=(self.cancel_event,), daemon=True).start()

    def cancel_process(self):
        self.log_firma("Annullamento richiesto...", "WARNING")
        self.cancel_event.set()
        self.cancel_button.config(state='disabled')

    def on_process_finished(self):
        self.toggle_buttons(is_running=False)
        pdf_dir = self.app_config.firma_pdf_dir.get()
        has_pdfs = os.path.isdir(pdf_dir) and any(f.lower().endswith('.pdf') for f in os.listdir(pdf_dir))
        self.prepare_button.config(state='normal' if has_pdfs else 'disabled')
        self.email_button.config(state='disabled')

    def toggle_buttons(self, is_running):
        if is_running:
            self.run_button.pack_forget()
            self.cancel_button.pack(fill=tk.X, ipady=8, pady=5)
            self.cancel_button.config(state='normal')
            self.prepare_button.config(state='disabled')
            self.email_button.config(state='disabled')
        else:
            self.cancel_button.pack_forget()
            self.run_button.pack(fill=tk.X, ipady=8, pady=5)
            self.run_button.config(state='normal')

    def prepare_email_drafts(self):
        # ... (rest of the file is unchanged)
        self.log_firma("Preparazione delle bozze email...", "HEADER")
        try:
            limit_mb_str = self.app_config.email_size_limit.get()
            limit_mb = float(limit_mb_str)
            limit_bytes = limit_mb * 1024 * 1024
        except (ValueError, TypeError):
            self.log_firma(f"ERRORE: Limite di dimensione non valido: '{limit_mb_str}'. Inserire un numero.", "ERROR")
            return
        pdf_dir = self.app_config.firma_pdf_dir.get()
        if not os.path.isdir(pdf_dir):
            self.log_firma(f"ERRORE: La cartella PDF non esiste: {pdf_dir}", "ERROR")
            return
        all_attachments = [(p, os.path.getsize(p)) for p in [os.path.join(pdf_dir, f) for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]]
        if not all_attachments:
            self.log_firma("Nessun file PDF trovato da allegare.", "WARNING")
            return
        chunks = []
        current_chunk = []
        current_chunk_size = 0
        for path, size in all_attachments:
            if current_chunk and current_chunk_size + size > limit_bytes:
                chunks.append(current_chunk)
                current_chunk = []
                current_chunk_size = 0
            current_chunk.append(path)
            current_chunk_size += size
        if current_chunk:
            chunks.append(current_chunk)
        self.prepared_drafts = []
        num_drafts = len(chunks)
        raw_subject = self.app_config.email_subject.get()
        base_subject = re.sub(r'^\[\d+/\d+\]\s*', '', raw_subject)
        self.app_config.email_subject.set(base_subject)
        for i, chunk in enumerate(chunks):
            draft = {}
            draft['to'] = self.app_config.email_to.get()
            draft['subject'] = f"[{i+1}/{num_drafts}] {base_subject}" if num_drafts > 1 else base_subject
            draft['attachments'] = chunk
            draft['file_list'] = [os.path.splitext(os.path.basename(p))[0] for p in chunk]
            base_template = self.email_body_text.get("1.0", tk.END).strip()
            if i == 0:
                draft['intro_text'] = base_template
            else:
                draft['intro_text'] = f"Seguito della mail precedente.\n\nElenco file:\n{{file_list}}"
            self.prepared_drafts.append(draft)
        self.log_firma(f"Preparate {len(self.prepared_drafts)} bozze di email.", "SUCCESS")
        self.current_draft_index = 0
        self._display_draft_preview()
        self.preview_frame.pack(side=tk.LEFT, padx=(20, 0))
        self.prepare_button.config(state='normal')
        self.email_button.config(state='normal')

    def _display_draft_preview(self):
        if not self.prepared_drafts:
            self.preview_frame.pack_forget()
            return
        draft = self.prepared_drafts[self.current_draft_index]
        self.preview_label['text'] = f"Anteprima {self.current_draft_index + 1}/{len(self.prepared_drafts)}"
        self.app_config.email_to.set(draft['to'])
        self.app_config.email_subject.set(draft['subject'])
        file_list_str = "\n".join([os.path.splitext(os.path.basename(p))[0] for p in draft['attachments']])
        full_body = draft['intro_text'].replace("{file_list}", file_list_str)
        self.email_body_text.delete("1.0", tk.END)
        self.email_body_text.insert("1.0", full_body)
        self.prev_button.config(state='normal' if self.current_draft_index > 0 else 'disabled')
        self.next_button.config(state='normal' if self.current_draft_index < len(self.prepared_drafts) - 1 else 'disabled')

    def show_prev_draft(self):
        if self.current_draft_index > 0:
            self.current_draft_index -= 1
            self._display_draft_preview()

    def show_next_draft(self):
        if self.current_draft_index < len(self.prepared_drafts) - 1:
            self.current_draft_index += 1
            self._display_draft_preview()

    def start_email_creation_process(self):
        self.toggle_buttons(is_running=True)
        threading.Thread(target=self.create_email_drafts_in_outlook, daemon=True).start()

    def create_email_drafts_in_outlook(self):
        try:
            if not self.prepared_drafts:
                self.log_firma("Nessuna bozza da creare.", "WARNING")
                return
            self.log_firma(f"Avvio creazione di {len(self.prepared_drafts)} bozze in Outlook...", "HEADER")
            email_handler = EmailHandler(self.log_firma)
            for draft_info in self.prepared_drafts:
                email_handler.create_outlook_draft(draft_info)
            self.log_firma("Creazione bozze in Outlook completata.", "SUCCESS")
            self.prepared_drafts = []
            self.master.after(0, self.preview_frame.pack_forget)
        finally:
            self.master.after(0, self.on_process_finished)

    def log_firma(self, message, level='INFO'):
        self.master.after(0, self.log_widget, message, level)

    def setup_progress(self, max_value):
        self.progress_frame.pack(fill=tk.X, pady=(10, 5), after=self.email_frame)
        self.progressbar['maximum'] = max_value
        self.progressbar['value'] = 0
        self.percent_label['text'] = "0%"

    def update_progress(self, value):
        self.progressbar['value'] = value
        max_val = self.progressbar['maximum']
        if max_val > 0:
            percent = (value / max_val) * 100
            self.percent_label['text'] = f"{percent:.0f}%"

    def hide_progress(self):
        self.progress_frame.pack_forget()

    def _update_email_preview(self, event=None):
        tcl_name = self.app_config.email_tcl.get()
        is_formal = self.app_config.email_is_formal.get()
        body_template = ""
        if tcl_name and tcl_name in self.app_config.TCL_CONTACTS:
            email = self.app_config.TCL_CONTACTS[tcl_name]
            self.app_config.email_to.set(email)
            first_name = tcl_name.split()[0]
            if is_formal:
                body_template = self.app_config.EMAIL_BODY_FORMAL.format(name=first_name, file_list="{file_list}")
            else:
                body_template = self.app_config.EMAIL_BODY_INFORMAL.format(name=first_name, file_list="{file_list}")
        else:
            self.app_config.email_to.set("")
            if is_formal:
                body_template = self.app_config.EMAIL_BODY_GENERIC_FORMAL.format(file_list="{file_list}")
            else:
                body_template = self.app_config.EMAIL_BODY_GENERIC_INFORMAL.format(file_list="{file_list}")
        self.email_body_text.delete("1.0", tk.END)
        self.email_body_text.insert("1.0", body_template)
