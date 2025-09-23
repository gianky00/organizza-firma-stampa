import tkinter as tk
from tkinter import ttk
from datetime import datetime
import threading
from src.logic.monthly_fees import MonthlyFeesProcessor
from src.utils.ui_utils import create_path_entry, select_file_dialog

class FeesTab(ttk.Frame):
    """
    GUI for the Monthly Fees Printing tab.
    """
    def __init__(self, parent, app_config, logger):
        super().__init__(parent)
        self.app_config = app_config
        self.log_widget = logger
        self.processor = MonthlyFeesProcessor(self, app_config)

        current_year = datetime.now().year
        self.anni_giornaliera = [str(y) for y in range(current_year - 5, current_year + 6)]

        self._create_widgets()
        self.after(100, self.populate_printers)
        self.after(150, self._update_paths_from_ui)

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        desc_text = "Questa sezione automatizza la stampa dei canoni mensili. Seleziona i file e il periodo, poi avvia il processo per eseguire macro VBA e stampare documenti Word in sequenza."
        desc_label = ttk.Label(main_frame, text=desc_text, wraplength=850, justify=tk.LEFT, style='info.TLabel')
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor='w')

        # --- Paths and Settings Frame ---
        paths_frame = ttk.LabelFrame(main_frame, text="1. File, Periodo e Macro per Stampa Canoni", padding="10")
        paths_frame.pack(fill=tk.X, pady=(0, 5))

        # --- Year/Month Selection ---
        selection_frame = ttk.Frame(paths_frame)
        selection_frame.grid(row=0, column=0, columnspan=3, sticky=tk.EW, pady=3)

        ttk.Label(selection_frame, text="Anno:").pack(side=tk.LEFT, padx=(0,5))
        self.anno_combo = ttk.Combobox(selection_frame, textvariable=self.app_config.canoni_selected_year, values=self.anni_giornaliera, state="readonly", width=10)
        self.anno_combo.pack(side=tk.LEFT, padx=(0,10))

        ttk.Label(selection_frame, text="Mese:").pack(side=tk.LEFT, padx=(5,5))
        self.mese_combo = ttk.Combobox(selection_frame, textvariable=self.app_config.canoni_selected_month, values=self.app_config.nomi_mesi_italiani, state="readonly", width=15)
        self.mese_combo.pack(side=tk.LEFT)

        # --- Path Entries ---
        create_path_entry(paths_frame, "File Giornaliera (Auto):", self.app_config.canoni_giornaliera_path, 1, readonly=True)
        create_path_entry(paths_frame, "N° Canone Messina:", self.app_config.canoni_messina_num, 2, readonly=False)
        create_path_entry(paths_frame, "N° Canone Naselli:", self.app_config.canoni_naselli_num, 3, readonly=False)
        create_path_entry(paths_frame, "N° Canone Caldarella:", self.app_config.canoni_caldarella_num, 4, readonly=False)

        word_ft = [("File Word", "*.docx *.doc"), ("Tutti i file", "*.*")]
        create_path_entry(paths_frame, "File Foglio Canone (Word):", self.app_config.canoni_word_path, 5, readonly=False,
                          browse_command=lambda: select_file_dialog(self.app_config.canoni_word_path, "Seleziona Foglio Canone Word", word_ft))

        # --- Printer Selection ---
        ttk.Label(paths_frame, text="Seleziona Stampante:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=3)
        self.printer_combo = ttk.Combobox(paths_frame, textvariable=self.app_config.selected_printer, state="readonly")
        self.printer_combo.grid(row=6, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=3)

        create_path_entry(paths_frame, "Nome Macro VBA:", self.app_config.canoni_macro_name, 7, readonly=True)

        # --- Actions Frame ---
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)
        actions_frame = ttk.LabelFrame(main_frame, text="2. Azione Stampa Canoni", padding="10")
        actions_frame.pack(fill=tk.X, pady=5)
        self.run_button = ttk.Button(actions_frame, text="▶  AVVIA PROCESSO STAMPA CANONI", style='primary.TButton', command=self.start_printing_process)
        self.run_button.pack(fill=tk.X, ipady=8, pady=5)

        # --- Bindings ---
        self.anno_combo.bind("<<ComboboxSelected>>", self._update_paths_from_ui)
        self.mese_combo.bind("<<ComboboxSelected>>", self._update_paths_from_ui)
        self.app_config.canoni_messina_num.trace_add("write", self._update_paths_from_ui)
        self.app_config.canoni_naselli_num.trace_add("write", self._update_paths_from_ui)
        self.app_config.canoni_caldarella_num.trace_add("write", self._update_paths_from_ui)

    def populate_printers(self):
        printers, default_printer = self.processor.get_printers()
        self.printer_combo['values'] = printers

        saved_printer = self.app_config.selected_printer.get()
        if saved_printer and saved_printer in printers:
            self.printer_combo.set(saved_printer)
        elif default_printer in printers:
            self.printer_combo.set(default_printer)
        elif printers:
            self.printer_combo.set(printers[0])

    def _update_paths_from_ui(self, *args):
        year = self.app_config.canoni_selected_year.get()
        month = self.app_config.canoni_selected_month.get()

        # Update Giornaliera path
        giornaliera_path = self.processor.get_giornaliera_path(year, month)
        self.app_config.canoni_giornaliera_path.set(giornaliera_path)

        # Update Consuntivo paths
        c1_path = self.processor.get_consuntivo_path(year, self.app_config.canoni_messina_num.get())
        self.app_config.canoni_cons1_path.set(c1_path)
        c2_path = self.processor.get_consuntivo_path(year, self.app_config.canoni_naselli_num.get())
        self.app_config.canoni_cons2_path.set(c2_path)
        c3_path = self.processor.get_consuntivo_path(year, self.app_config.canoni_caldarella_num.get())
        self.app_config.canoni_cons3_path.set(c3_path)

    def start_printing_process(self):
        self.toggle_stampa_canoni_buttons('disabled')

        # Gather all paths and settings for the processor
        paths_to_print = {
            "giornaliera": self.app_config.canoni_giornaliera_path.get(),
            "consuntivi": [
                self.app_config.canoni_cons1_path.get(),
                self.app_config.canoni_cons2_path.get(),
                self.app_config.canoni_cons3_path.get()
            ],
            "word": self.app_config.canoni_word_path.get()
        }
        printer = self.app_config.selected_printer.get()
        macro = self.app_config.canoni_macro_name.get()

        threading.Thread(target=self.processor.run_printing_process,
                         args=(paths_to_print, printer, macro),
                         daemon=True).start()

    def toggle_stampa_canoni_buttons(self, state):
        self.run_button.config(state=state)

    def log_canoni(self, message, level='INFO'):
        self.master.after(0, self.log_widget, message, level)
