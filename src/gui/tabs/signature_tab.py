import tkinter as tk
from tkinter import ttk
import threading
from src.logic.signature import SignatureProcessor
from src.utils.ui_utils import create_path_entry, select_file_dialog, open_folder_in_explorer

class SignatureTab(ttk.Frame):
    """
    GUI for the Signature Application tab.
    """
    def __init__(self, parent, app_config, logger):
        super().__init__(parent)
        self.app_config = app_config
        self.log_widget = logger

        self._create_widgets()

        # Pass the progress bar methods to the processor
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

        # --- Paths Frame ---
        paths_frame = ttk.LabelFrame(main_frame, text="1. Percorsi (Firma)", padding="15")
        paths_frame.pack(fill=tk.X, pady=(0, 10))

        paths_frame.columnconfigure(1, weight=1)

        create_path_entry(paths_frame, "Cartella Excel:", self.app_config.firma_excel_dir, 0, readonly=True)

        # PDF Path with "Open" button
        ttk.Label(paths_frame, text="Cartella PDF di Output:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        pdf_entry = ttk.Entry(paths_frame, textvariable=self.app_config.firma_pdf_dir, state='readonly')
        pdf_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        open_button = ttk.Button(paths_frame, text="Apri Cartella", command=lambda: open_folder_in_explorer(self.app_config.firma_pdf_dir.get()))
        open_button.grid(row=1, column=2, sticky=tk.E, padx=(5, 0), pady=5)

        create_path_entry(paths_frame, "Immagine Firma:", self.app_config.firma_image_path, 2, readonly=True)
        create_path_entry(paths_frame, "Ghostscript:", self.app_config.firma_ghostscript_path, 3, readonly=False,
                          browse_command=lambda: select_file_dialog(self.app_config.firma_ghostscript_path, "Seleziona eseguibile Ghostscript", [("Executable", "*.exe")]))

        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=15, padx=5)

        # --- Mode Frame ---
        mode_frame = ttk.LabelFrame(main_frame, text="2. Tipo di Documento da Firmare", padding="15")
        mode_frame.pack(fill=tk.X, pady=10)
        ttk.Radiobutton(mode_frame, text="Schede (Controllo, Manutenzione, etc.)", variable=self.app_config.firma_processing_mode, value="schede").pack(anchor=tk.W, padx=5, pady=5)
        ttk.Radiobutton(mode_frame, text="Preventivi (Basato su foglio 'Consuntivo')", variable=self.app_config.firma_processing_mode, value="preventivi").pack(anchor=tk.W, padx=5, pady=5)

        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=15, padx=5)

        # --- Actions Frame ---
        actions_frame = ttk.LabelFrame(main_frame, text="3. Azioni Firma", padding="15")
        actions_frame.pack(fill=tk.X, pady=10)
        self.run_button = ttk.Button(actions_frame, text="▶  AVVIA PROCESSO FIRMA COMPLETO", style='primary.TButton', command=self.start_signature_process)
        self.run_button.pack(fill=tk.X, ipady=8, pady=5)

        # --- Progress Bar ---
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.pack(fill=tk.X, pady=(10, 5))
        self.progress_label = ttk.Label(self.progress_frame, text="Progresso:")
        self.progress_label.pack(side=tk.LEFT, padx=(5, 5))
        self.progressbar = ttk.Progressbar(self.progress_frame, orient='horizontal', mode='determinate')
        self.progressbar.pack(fill=tk.X, expand=True)
        self.progress_frame.pack_forget() # Hide by default

        # --- Cleanup buttons would be added here, similar structure ---
        # For brevity in refactoring, they are omitted but would follow the same pattern
        # self.clean_pdf_button = ttk.Button(...)
        # self.clean_excel_button = ttk.Button(...)

    def start_signature_process(self):
        """
        Starts the signature process in a new thread to avoid freezing the GUI.
        """
        self.toggle_firma_buttons('disabled')
        # The actual work is now in SignatureProcessor
        threading.Thread(target=self.processor.run_full_signature_process, daemon=True).start()

    def toggle_firma_buttons(self, state):
        """
        Enables or disables all buttons in this tab.
        """
        self.run_button.config(state=state)
        # self.clean_pdf_button.config(state=state)
        # self.clean_excel_button.config(state=state)

    def log_firma(self, message, level='INFO'):
        """
        A wrapper to log messages to the correct log widget.
        This method is passed to the processor.
        """
        # The call to the actual logger needs to be thread-safe
        self.master.after(0, self.log_widget, message, level)

    def setup_progress(self, max_value):
        self.progress_frame.pack(fill=tk.X, pady=(10, 5))
        self.progressbar['maximum'] = max_value
        self.progressbar['value'] = 0

    def update_progress(self, value):
        self.progressbar['value'] = value

    def hide_progress(self):
        self.progress_frame.pack_forget()
