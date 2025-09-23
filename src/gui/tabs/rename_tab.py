import tkinter as tk
from tkinter import ttk
import threading
from src.logic.renaming import RenameProcessor
from src.utils.ui_utils import create_path_entry, select_folder_dialog

class RenameTab(ttk.Frame):
    """
    GUI for the Rename Files tab.
    """
    def __init__(self, parent, app_config, logger):
        super().__init__(parent)
        self.app_config = app_config
        self.log_widget = logger

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

        # --- Paths Frame ---
        paths_frame = ttk.LabelFrame(main_frame, text="1. Impostazioni di Ridenominazione", padding="15")
        paths_frame.pack(fill=tk.X, pady=(0, 10))
        create_path_entry(paths_frame, "Cartella Schede da Rinominare:", self.app_config.rinomina_path, 0, readonly=False,
                          browse_command=lambda: select_folder_dialog(self.app_config.rinomina_path, "Seleziona cartella con le schede da rinominare"))
        create_path_entry(paths_frame, "Password per file protetti:", self.app_config.rinomina_password, 1, readonly=False)

        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=15, padx=5)

        # --- Actions Frame ---
        actions_frame = ttk.LabelFrame(main_frame, text="2. Azioni", padding="15")
        actions_frame.pack(fill=tk.X, pady=10)
        self.run_button = ttk.Button(actions_frame, text="▶  AVVIA PROCESSO DI RINOMINA", style='primary.TButton', command=self.start_rename_process)
        self.run_button.pack(fill=tk.X, ipady=8, pady=5)

        # --- Progress Bar ---
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.pack(fill=tk.X, pady=(10, 5))
        self.progress_label = ttk.Label(self.progress_frame, text="Progresso:")
        self.progress_label.pack(side=tk.LEFT, padx=(5, 5))
        self.progressbar = ttk.Progressbar(self.progress_frame, orient='horizontal', mode='determinate')
        self.progressbar.pack(fill=tk.X, expand=True)
        self.progress_frame.pack_forget() # Hide by default

    def start_rename_process(self):
        """
        Starts the renaming process in a new thread.
        """
        self.toggle_rinomina_buttons('disabled')
        threading.Thread(target=self.processor.run_rename_process, daemon=True).start()

    def toggle_rinomina_buttons(self, state):
        """
        Enables or disables the buttons in this tab.
        """
        self.run_button.config(state=state)

    def log_rinomina(self, message, level='INFO'):
        """
        Thread-safe logging method for the rename processor.
        """
        self.master.after(0, self.log_widget, message, level)

    def setup_progress(self, max_value):
        self.progress_frame.pack(fill=tk.X, pady=(10, 5))
        self.progressbar['maximum'] = max_value
        self.progressbar['value'] = 0

    def update_progress(self, value):
        self.progressbar['value'] = value

    def hide_progress(self):
        self.progress_frame.pack_forget()
