import threading
import tkinter as tk
from datetime import datetime
from tkinter import ttk

from src.logic.monthly_fees import MonthlyFeesProcessor
from src.utils.ui_utils import ProgressWithETA, create_path_entry, select_file_dialog


class FeesTab(ttk.Frame):
    def __init__(
        self, parent, app_config, logger, processor_class=None, excel_gateway_class=None, word_gateway_class=None
    ):
        super().__init__(parent)
        self.app_config = app_config
        self.log_widget = logger
        self.cancel_event = threading.Event()
        processor_cls = processor_class or MonthlyFeesProcessor
        self.processor = processor_cls(
            self,
            app_config,
            excel_gateway_class=excel_gateway_class,
            word_gateway_class=word_gateway_class,
        )
        current_year = datetime.now().year
        self.anni_giornaliera = [str(y) for y in range(current_year - 5, current_year + 6)]
        self._create_widgets()
        self.after(100, self.populate_printers)
        self.after(150, self._update_paths_from_ui)

    def _create_widgets(self):
        self.columnconfigure(0, weight=1)

        # --- Description ---
        desc_label = ttk.Label(
            self,
            text="Automatizza la stampa dei canoni mensili eseguendo macro VBA su file Excel e stampando documenti Word in sequenza.",
            wraplength=800,
            justify=tk.LEFT,
            style="info.TLabel",
        )
        desc_label.pack(fill=tk.X, pady=(0, 15), anchor="w")

        # --- Settings Frame ---
        settings_frame = ttk.LabelFrame(self, text="1. Impostazioni di Stampa", padding=15)
        settings_frame.pack(fill=tk.X, pady=5)
        settings_frame.columnconfigure(0, weight=1)

        # --- Periodo ---
        period_frame = ttk.Frame(settings_frame)
        period_frame.grid(row=0, column=0, sticky=tk.EW, pady=(0, 10))
        ttk.Label(period_frame, text="Periodo:", font=self.app_config.font_bold).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(period_frame, text="Anno:").pack(side=tk.LEFT, padx=(5, 5))
        self.anno_combo = ttk.Combobox(
            period_frame,
            textvariable=self.app_config.canoni_selected_year,
            values=self.anni_giornaliera,
            state="readonly",
            width=10,
        )
        self.anno_combo.pack(side=tk.LEFT, padx=(0, 15))
        ttk.Label(period_frame, text="Mese:").pack(side=tk.LEFT, padx=(5, 5))
        self.mese_combo = ttk.Combobox(
            period_frame,
            textvariable=self.app_config.canoni_selected_month,
            values=self.app_config.nomi_mesi_italiani,
            state="readonly",
            width=15,
        )
        self.mese_combo.pack(side=tk.LEFT, padx=(0, 5))

        # --- Numeri Consuntivo ---
        self.consuntivi_frame = ttk.LabelFrame(settings_frame, text="Numeri Canoni Mensili", padding=10)
        self.consuntivi_frame.grid(row=1, column=0, sticky=tk.EW, pady=5)
        self.consuntivi_frame.columnconfigure(1, weight=1)

        # Container per la tabella (Header + Righe)
        self.table_container = ttk.Frame(self.consuntivi_frame)
        self.table_container.pack(fill=tk.X)
        self.table_container.columnconfigure(1, weight=1)  # Nome TCL

        # Bottoni di Gestione
        mgmt_f = ttk.Frame(self.consuntivi_frame)
        mgmt_f.pack(fill=tk.X, pady=(10, 0))

        self.find_numbers_button = ttk.Button(
            mgmt_f, text="🔍 Trova Numeri Automaticamente", command=self.find_numbers_and_populate
        )
        self.find_numbers_button.pack(side=tk.LEFT, padx=5)

        self._refresh_dynamic_tcl_ui()

        # --- Altri Percorsi ---
        paths_frame = ttk.LabelFrame(settings_frame, text="Percorsi File", padding=10)
        paths_frame.grid(row=2, column=0, sticky=tk.EW, pady=5)
        paths_frame.columnconfigure(0, weight=1)
        create_path_entry(
            paths_frame,
            "File Giornaliera (Auto):",
            self.app_config.canoni_giornaliera_path,
            None,
            0,
            readonly=True,
        )
        word_ft = [("File Word", "*.docx *.doc"), ("Tutti i file", "*.*")]
        create_path_entry(
            paths_frame,
            "File Foglio Canone (Word):",
            self.app_config.canoni_word_path,
            lambda: select_file_dialog(self.app_config.canoni_word_path, word_ft),
            1,
            readonly=False,
        )

        # --- Stampante e Macro ---
        printer_macro_frame = ttk.LabelFrame(settings_frame, text="Dispositivo e Macro", padding=10)
        printer_macro_frame.grid(row=3, column=0, sticky=tk.EW, pady=5)
        printer_macro_frame.columnconfigure(0, weight=1)

        p_frame = ttk.Frame(printer_macro_frame)
        p_frame.grid(row=0, column=0, sticky="ew", pady=5)
        p_frame.columnconfigure(1, weight=1)
        ttk.Label(p_frame, text="Stampante:", width=25).grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.printer_combo = ttk.Combobox(p_frame, textvariable=self.app_config.selected_printer, state="readonly")
        self.printer_combo.grid(row=0, column=1, sticky=tk.EW, padx=5)
        create_path_entry(
            printer_macro_frame,
            "Nome Macro VBA:",
            self.app_config.canoni_macro_name,
            None,
            1,
            readonly=True,
        )

        # --- Azioni ---
        self.actions_frame = ttk.LabelFrame(self, text="2. Azione", padding=15)
        self.actions_frame.pack(fill=tk.X, pady=5)
        self.actions_frame.columnconfigure(0, weight=1)
        self.run_button = ttk.Button(
            self.actions_frame,
            text="▶ AVVIA PROCESSO STAMPA CANONI",
            style="primary.TButton",
            command=self.start_printing_process,
        )
        self.run_button.pack(fill=tk.X, ipady=8)
        self.cancel_button = ttk.Button(self.actions_frame, text="Annulla Processo", command=self.cancel_process)
        # self.cancel_button is packed dynamically

        # --- Progress Bar ---
        self.progress_frame = ProgressWithETA(self)

        self.anno_combo.bind("<<ComboboxSelected>>", self._update_paths_from_ui)
        self.mese_combo.bind("<<ComboboxSelected>>", self._update_paths_from_ui)
        self._setup_tcl_traces()

        self.on_process_finished()

    def _setup_tcl_traces(self):
        from contextlib import suppress

        for ref in self.app_config.canoni_tcl_vars:
            # Rimuoviamo eventuali trace vecchie per evitare duplicati
            with suppress(Exception):
                for mode, callbacks in ref["num"].trace_info() + ref["name"].trace_info():
                    ref["num"].trace_remove(mode, callbacks[0])
                    ref["name"].trace_remove(mode, callbacks[0])

            ref["num"].trace_add("write", self._update_paths_from_ui)
            # Se cambia il nome, aggiorniamo il log o la vista se necessario
            ref["name"].trace_add("write", lambda *args: self.log_canoni("Nominativo aggiornato", "DEBUG"))

    def _refresh_dynamic_tcl_ui(self):
        # Pulisce tutto il container della tabella
        for widget in self.table_container.winfo_children():
            widget.destroy()

        # Header con larghezze fisse per allineamento
        ttk.Label(self.table_container, text="Stampa", font=self.app_config.font_bold, width=8, anchor="center").grid(
            row=0, column=0, padx=5, pady=5
        )
        ttk.Label(self.table_container, text="Nome TCL", font=self.app_config.font_bold, width=35).grid(
            row=0, column=1, padx=5, pady=5, sticky="w"
        )
        ttk.Label(self.table_container, text="N° Canone", font=self.app_config.font_bold, width=15).grid(
            row=0, column=2, padx=5, pady=5, sticky="w"
        )

        for i, ref in enumerate(self.app_config.canoni_tcl_vars):
            row_idx = i + 1
            # Checkbutton centrato
            cb = ttk.Checkbutton(self.table_container, variable=ref["print"])
            cb.grid(row=row_idx, column=0, padx=5, pady=2)

            # Label per il nome (Fisso, non Entry per risparmiare spazio e confusione)
            name_lbl = ttk.Label(self.table_container, textvariable=ref["name"], width=35)
            name_lbl.grid(row=row_idx, column=1, padx=5, pady=2, sticky="w")

            # Entry per il numero (unica cosa modificabile qui)
            num_ent = ttk.Entry(self.table_container, textvariable=ref["num"], width=15)
            num_ent.grid(row=row_idx, column=2, padx=5, pady=2, sticky="w")

    # Rimosse le funzioni di gestione da questa tab per pulizia
    def _add_tcl(self):
        pass

    def _remove_tcl(self, index):
        pass

    def populate_printers(self):
        printers, default_printer = self.processor.get_printers()
        self.printer_combo["values"] = printers
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
        threading.Thread(
            target=self.processor.run_printing_process,
            args=(self.cancel_event, paths_to_print, printer, macro),
            daemon=True,
        ).start()

    def find_numbers_and_populate(self):
        self.cancel_event.clear()
        self.toggle_buttons(is_running=True)
        self.log_canoni("Ricerca automatica dei numeri di canone in corso...", "HEADER")
        threading.Thread(target=self._find_numbers_thread, args=(self.cancel_event,), daemon=True).start()

    def _find_numbers_thread(self, cancel_event):
        try:
            year = self.app_config.canoni_selected_year.get()
            month = self.app_config.canoni_selected_month.get()

            # Mappa TCL -> StringVar per il popolamento
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
                    self.master.after(0, var.set, number)
        finally:
            self.master.after(0, self.on_process_finished)

    def cancel_process(self):
        self.log_canoni("Annullamento richiesto...", "WARNING")
        self.cancel_event.set()
        self.cancel_button.config(state="disabled")

    def on_process_finished(self):
        self.toggle_buttons(is_running=False)
        self.hide_progress()

    def toggle_buttons(self, is_running):
        state = "disabled" if is_running else "normal"
        self.run_button.config(state=state)
        self.find_numbers_button.config(state=state)
        if is_running:
            self.run_button.pack_forget()
            self.cancel_button.pack(fill=tk.X, ipady=8, pady=5)
            self.cancel_button.config(state="normal")
        else:
            self.cancel_button.pack_forget()
            self.run_button.pack(fill=tk.X, ipady=8, pady=5)

    def show_progress(self):
        self.app_config.show_global_indeterminate("Ricerca in corso...")

    def hide_progress(self):
        self.app_config.hide_global_progress()

    def log_canoni(self, message, level="INFO"):
        self.master.after(0, self.log_widget, message, level)
