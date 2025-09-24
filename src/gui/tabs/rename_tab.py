import tkinter as tk
from tkinter import ttk
import threading
from src.logic.renaming import RenameProcessor
from src.utils.ui_utils import create_path_entry, select_folder_dialog

class RenameTab(ttk.Frame):
    def __init__(self, parent, app_config, logger):
        super().__init__(parent)
        self.app_config = app_config
        self.log_widget = logger
        self.cancel_event = threading.Event()

        self._create_widgets()

        self.processor = RenameProcessor(
            self,
            app_config,
            self.setup_progress,
            self.update_progress,
            self.hide_progress
        )

    def _create_widgets(self):
        self.columnconfigure(0, weight=1)

        # --- Description ---
        desc_text = "Analizza i file Excel in una cartella, trova la data di emissione e li rinomina nel formato NOME (GG-MM-AAAA). Prova ad usare una password per i file protetti."
        desc_label = ttk.Label(self, text=desc_text, wraplength=800, justify=tk.LEFT, style='info.TLabel')
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor='w')

        # --- Settings Frame ---
        settings_frame = ttk.LabelFrame(self, text="1. Impostazioni", padding=15)
        settings_frame.pack(fill=tk.X, pady=5)
        settings_frame.columnconfigure(1, weight=1)

        create_path_entry(settings_frame, "Cartella da Analizzare:", self.app_config.rinomina_path, 0, readonly=False,
                          browse_command=lambda: select_folder_dialog(self.app_config.rinomina_path, "Seleziona cartella con le schede da rinominare"))
        create_path_entry(settings_frame, "Password (opzionale):", self.app_config.rinomina_password, 1, readonly=False)

        # --- Actions Frame ---
        self.actions_frame = ttk.LabelFrame(self, text="2. Azioni", padding=15)
        self.actions_frame.pack(fill=tk.X, pady=5)
        self.actions_frame.columnconfigure(0, weight=1)

        self.run_button = ttk.Button(self.actions_frame, text="▶ AVVIA PROCESSO DI RINOMINA", style='primary.TButton', command=self.start_rename_process)
        self.run_button.pack(fill=tk.X, ipady=8)
        self.cancel_button = ttk.Button(self.actions_frame, text="Annulla Processo", command=self.cancel_process)
        # self.cancel_button is packed dynamically

        # --- Progress Bar ---
        self.progress_frame = ttk.Frame(self)
        self.progress_label = ttk.Label(self.progress_frame, text="Progresso:")
        self.progress_label.pack(side=tk.LEFT, padx=(0, 5))
        self.progressbar = ttk.Progressbar(self.progress_frame, orient='horizontal', mode='determinate')
        self.progressbar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.percent_label = ttk.Label(self.progress_frame, text="0%", width=5)
        self.percent_label.pack(side=tk.LEFT, padx=(5, 0))

        self.on_process_finished()

    def start_rename_process(self):
        self.cancel_event.clear()
        self.toggle_buttons(is_running=True)
        threading.Thread(target=self.processor.run_rename_process, args=(self.cancel_event,), daemon=True).start()

    def cancel_process(self):
        self.log_rinomina("Annullamento richiesto...", "WARNING")
        self.cancel_event.set()
        self.cancel_button.config(state='disabled')

    def on_process_finished(self):
        self.toggle_buttons(is_running=False)

    def toggle_buttons(self, is_running):
        if is_running:
            self.run_button.pack_forget()
            self.cancel_button.pack(fill=tk.X, ipady=8, pady=5)
            self.cancel_button.config(state='normal')
        else:
            self.cancel_button.pack_forget()
            self.run_button.pack(fill=tk.X, ipady=8, pady=5)
            self.run_button.config(state='normal')

    def log_rinomina(self, message, level='INFO'):
        self.master.after(0, self.log_widget, message, level)

    def setup_progress(self, max_value):
        self.progress_frame.pack(fill=tk.X, pady=(10, 5), after=self.actions_frame)
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
