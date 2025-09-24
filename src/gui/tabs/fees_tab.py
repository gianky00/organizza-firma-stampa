import tkinter as tk
from tkinter import ttk
from datetime import datetime
import threading
from src.logic.monthly_fees import MonthlyFeesProcessor
from src.utils.ui_utils import create_path_entry, select_file_dialog

class FeesTab(ttk.Frame):
    def __init__(self, parent, app_config, logger):
        super().__init__(parent)
        self.app_config = app_config
        self.log_widget = logger
        self.cancel_event = threading.Event()
        self.processor = MonthlyFeesProcessor(self, app_config)
        current_year = datetime.now().year
        self.anni_giornaliera = [str(y) for y in range(current_year - 5, current_year + 6)]
        self._create_widgets()
        self.after(100, self.populate_printers)
        self.after(150, self._update_paths_from_ui)

    def _create_widgets(self):
        self.columnconfigure(0, weight=1)

        # --- Description ---
        desc_label = ttk.Label(self, text="Automatizza la stampa dei canoni mensili eseguendo macro VBA su file Excel e stampando documenti Word in sequenza.", wraplength=800, justify=tk.LEFT, style='info.TLabel')
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor='w')

        # --- Settings Frame ---
        settings_frame = ttk.LabelFrame(self, text="1. Impostazioni di Stampa", padding=15)
        settings_frame.pack(fill=tk.X, pady=5)
        settings_frame.columnconfigure(1, weight=1)

        # --- Periodo ---
        period_frame = ttk.Frame(settings_frame)
        period_frame.grid(row=0, column=0, columnspan=2, sticky=tk.EW, pady=(0, 10))
        ttk.Label(period_frame, text="Periodo:", font=self.app_config.font_bold).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(period_frame, text="Anno:").pack(side=tk.LEFT, padx=(5,5))
        self.anno_combo = ttk.Combobox(period_frame, textvariable=self.app_config.canoni_selected_year, values=self.anni_giornaliera, state="readonly", width=10)
        self.anno_combo.pack(side=tk.LEFT, padx=(0,15))
        ttk.Label(period_frame, text="Mese:").pack(side=tk.LEFT, padx=(5,5))
        self.mese_combo = ttk.Combobox(period_frame, textvariable=self.app_config.canoni_selected_month, values=self.app_config.nomi_mesi_italiani, state="readonly", width=15)
        self.mese_combo.pack(side=tk.LEFT, padx=(0, 5))

        # --- Numeri Consuntivo ---
        consuntivi_frame = ttk.LabelFrame(settings_frame, text="Numeri Consuntivo", padding=10)
        consuntivi_frame.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=5)
        consuntivi_frame.columnconfigure(1, weight=1)
        create_path_entry(consuntivi_frame, "N° Canone Messina:", self.app_config.canoni_messina_num, 0, readonly=False)
        create_path_entry(consuntivi_frame, "N° Canone Naselli:", self.app_config.canoni_naselli_num, 1, readonly=False)
        create_path_entry(consuntivi_frame, "N° Canone Caldarella:", self.app_config.canoni_caldarella_num, 2, readonly=False)
        self.find_numbers_button = ttk.Button(consuntivi_frame, text="Trova Numeri Automaticamente", command=self.find_numbers_and_populate)
        self.find_numbers_button.grid(row=3, column=0, columnspan=2, sticky="we", pady=(10, 0))

        # --- Altri Percorsi ---
        paths_frame = ttk.LabelFrame(settings_frame, text="Percorsi File", padding=10)
        paths_frame.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=5)
        paths_frame.columnconfigure(1, weight=1)
        create_path_entry(paths_frame, "File Giornaliera (Auto):", self.app_config.canoni_giornaliera_path, 0, readonly=True)
        word_ft = [("File Word", "*.docx *.doc"), ("Tutti i file", "*.*")]
        create_path_entry(paths_frame, "File Foglio Canone (Word):", self.app_config.canoni_word_path, 1, readonly=False, browse_command=lambda: select_file_dialog(self.app_config.canoni_word_path, "Seleziona Foglio Canone Word", word_ft))

        # --- Stampante e Macro ---
        printer_macro_frame = ttk.LabelFrame(settings_frame, text="Dispositivo e Macro", padding=10)
        printer_macro_frame.grid(row=3, column=0, columnspan=2, sticky=tk.EW, pady=5)
        printer_macro_frame.columnconfigure(1, weight=1)
        ttk.Label(printer_macro_frame, text="Stampante:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.printer_combo = ttk.Combobox(printer_macro_frame, textvariable=self.app_config.selected_printer, state="readonly")
        self.printer_combo.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        create_path_entry(printer_macro_frame, "Nome Macro VBA:", self.app_config.canoni_macro_name, 1, readonly=True)

        # --- Azioni ---
        self.actions_frame = ttk.LabelFrame(self, text="2. Azione", padding=15)
        self.actions_frame.pack(fill=tk.X, pady=5)
        self.actions_frame.columnconfigure(0, weight=1)
        self.run_button = ttk.Button(self.actions_frame, text="▶ AVVIA PROCESSO STAMPA CANONI", style='primary.TButton', command=self.start_printing_process)
        self.run_button.pack(fill=tk.X, ipady=8)
        self.cancel_button = ttk.Button(self.actions_frame, text="Annulla Processo", command=self.cancel_process)
        # self.cancel_button is packed dynamically

        # --- Progress Bar ---
        self.progress_frame = ttk.Frame(self)
        self.progress_label = ttk.Label(self.progress_frame, text="Elaborazione in corso...")
        self.progress_label.pack(side=tk.LEFT, padx=(0, 5))
        self.progressbar = ttk.Progressbar(self.progress_frame, orient='horizontal', mode='indeterminate')
        self.progressbar.pack(fill=tk.X, expand=True)

        self.anno_combo.bind("<<ComboboxSelected>>", self._update_paths_from_ui)
        self.mese_combo.bind("<<ComboboxSelected>>", self._update_paths_from_ui)
        self.app_config.canoni_messina_num.trace_add("write", self._update_paths_from_ui)
        self.app_config.canoni_naselli_num.trace_add("write", self._update_paths_from_ui)
        self.app_config.canoni_caldarella_num.trace_add("write", self._update_paths_from_ui)

        self.on_process_finished()

    def populate_printers(self):
        printers, default_printer = self.processor.get_printers()
        self.printer_combo['values'] = printers
        saved_printer = self.app_config.selected_printer.get()
        if saved_printer and saved_printer in printers: self.printer_combo.set(saved_printer)
        elif default_printer in printers: self.printer_combo.set(default_printer)
        elif printers: self.printer_combo.set(printers[0])

    def _update_paths_from_ui(self, *args):
        year = self.app_config.canoni_selected_year.get()
        month = self.app_config.canoni_selected_month.get()
        giornaliera_path = self.processor.get_giornaliera_path(year, month)
        self.app_config.canoni_giornaliera_path.set(giornaliera_path)
        c1_path = self.processor.get_consuntivo_path(year, self.app_config.canoni_messina_num.get())
        self.app_config.canoni_cons1_path.set(c1_path)
        c2_path = self.processor.get_consuntivo_path(year, self.app_config.canoni_naselli_num.get())
        self.app_config.canoni_cons2_path.set(c2_path)
        c3_path = self.processor.get_consuntivo_path(year, self.app_config.canoni_caldarella_num.get())
        self.app_config.canoni_cons3_path.set(c3_path)

    def start_printing_process(self):
        self.cancel_event.clear()
        self.toggle_buttons(is_running=True)
        self.show_progress()
        paths_to_print = {
            "giornaliera": self.app_config.canoni_giornaliera_path.get(),
            "consuntivi": [self.app_config.canoni_cons1_path.get(), self.app_config.canoni_cons2_path.get(), self.app_config.canoni_cons3_path.get()],
            "word": self.app_config.canoni_word_path.get()
        }
        printer = self.app_config.selected_printer.get()
        macro = self.app_config.canoni_macro_name.get()
        threading.Thread(target=self.processor.run_printing_process, args=(self.cancel_event, paths_to_print, printer, macro), daemon=True).start()

    def find_numbers_and_populate(self):
        self.cancel_event.clear()
        self.toggle_buttons(is_running=True)
        self.log_canoni("Ricerca automatica dei numeri di canone in corso...", "HEADER")
        threading.Thread(target=self._find_numbers_thread, args=(self.cancel_event,), daemon=True).start()

    def _find_numbers_thread(self, cancel_event):
        try:
            year = self.app_config.canoni_selected_year.get()
            month = self.app_config.canoni_selected_month.get()
            tcls_to_find = {"MESSINA": self.app_config.canoni_messina_num, "NASELLI": self.app_config.canoni_naselli_num, "CALDARELLA": self.app_config.canoni_caldarella_num}
            for tcl, var in tcls_to_find.items():
                if cancel_event.is_set():
                    self.log_canoni("Ricerca annullata.", "WARNING")
                    break
                number, _ = self.processor.find_consuntivo_for_tcl(year, month, tcl, cancel_event)
                if number:
                    self.master.after(0, var.set, number)
        finally:
            self.master.after(0, self.on_process_finished)

    def cancel_process(self):
        self.log_canoni("Annullamento richiesto...", "WARNING")
        self.cancel_event.set()
        self.cancel_button.config(state='disabled')

    def on_process_finished(self):
        self.toggle_buttons(is_running=False)
        self.hide_progress()

    def toggle_buttons(self, is_running):
        state = 'disabled' if is_running else 'normal'
        self.run_button.config(state=state)
        self.find_numbers_button.config(state=state)
        if is_running:
            self.run_button.pack_forget()
            self.cancel_button.pack(fill=tk.X, ipady=8, pady=5)
            self.cancel_button.config(state='normal')
        else:
            self.cancel_button.pack_forget()
            self.run_button.pack(fill=tk.X, ipady=8, pady=5)

    def show_progress(self):
        self.progress_frame.pack(fill=tk.X, pady=(10, 5), after=self.actions_frame)
        self.progressbar.start(10)

    def hide_progress(self):
        self.progressbar.stop()
        self.progress_frame.pack_forget()

    def log_canoni(self, message, level='INFO'):
        self.master.after(0, self.log_widget, message, level)
