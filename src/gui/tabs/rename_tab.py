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
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        desc_text = "Questa funzione analizza tutti i file Excel in una cartella, cerca la data di emissione al loro interno e rinomina i file aggiungendo la data nel formato (GG-MM-AAAA). Se un file è protetto da password, proverà ad usare la password 'coemi'."
        desc_label = ttk.Label(main_frame, text=desc_text, wraplength=850, justify=tk.LEFT, style='info.TLabel')
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor='w')

        paths_frame = ttk.LabelFrame(main_frame, text="1. Impostazioni di Ridenominazione", padding="15")
        paths_frame.pack(fill=tk.X, pady=(0, 10))
        create_path_entry(paths_frame, "Cartella Schede da Rinominare:", self.app_config.rinomina_path, 0, readonly=False,
                          browse_command=lambda: select_folder_dialog(self.app_config.rinomina_path, "Seleziona cartella con le schede da rinominare"))
        create_path_entry(paths_frame, "Password per file protetti:", self.app_config.rinomina_password, 1, readonly=False)

        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=15, padx=5)

        actions_frame = ttk.LabelFrame(main_frame, text="2. Azioni", padding="15")
        actions_frame.pack(fill=tk.X, pady=10)
        self.run_button = ttk.Button(actions_frame, text="▶  AVVIA PROCESSO DI RINOMINA", style='primary.TButton', command=self.start_rename_process)
        self.run_button.pack(fill=tk.X, ipady=8, pady=5)
        self.cancel_button = ttk.Button(actions_frame, text="Annulla Processo", command=self.cancel_process)
        self.cancel_button.pack(fill=tk.X, ipady=8, pady=5)

        self.on_process_finished()

        self.progress_frame = ttk.Frame(main_frame)
        self.progress_label = ttk.Label(self.progress_frame, text="Progresso:")
        self.progress_label.pack(side=tk.LEFT, padx=(5, 5))
        self.progressbar = ttk.Progressbar(self.progress_frame, orient='horizontal', mode='determinate')
        self.progressbar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.percent_label = ttk.Label(self.progress_frame, text="0%", width=5)
        self.percent_label.pack(side=tk.LEFT)
        self.progress_frame.pack_forget()

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
        self.progress_frame.pack(fill=tk.X, pady=(10, 5), after=self.run_button.master.master)
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
