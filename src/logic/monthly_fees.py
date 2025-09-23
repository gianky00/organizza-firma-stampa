import os
import re
import win32print
import traceback
from datetime import datetime
from src.utils.excel_handler import ExcelHandler
from src.utils.word_handler import WordHandler

class MonthlyFeesProcessor:
    def __init__(self, gui, config):
        self.gui = gui
        self.config = config
        self.logger = gui.log_canoni

    def get_printers(self):
        try:
            printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            default_printer = win32print.GetDefaultPrinter()
            return printers, default_printer
        except Exception as e:
            self.logger(f"Errore nel caricamento delle stampanti: {e}", "ERROR")
            return [], None

    def get_giornaliera_path(self, year, month_name):
        if not year or not month_name: return "Seleziona Anno e Mese"
        month_number = self.config.mesi_giornaliera_map.get(month_name)
        if not month_number: return "Mese non valido"
        year_folder_name = f"Giornaliere {year}"
        file_name = f"Giornaliera {month_number}-{year}.xlsm"
        return os.path.join(self.config.CANONI_GIORNALIERA_BASE_DIR, year_folder_name, file_name)

    def get_consuntivo_path(self, year, consuntivo_num):
        if not year: return "Anno non selezionato"
        if not consuntivo_num.strip().isdigit(): return "Inserire un numero valido"
        cons_dir = os.path.join(self.config.CANONI_CONSUNTIVI_BASE_DIR, year, "CONSUNTIVI", year)
        if not os.path.isdir(cons_dir): return f"ERRORE: Cartella non trovata"
        try:
            for filename in os.listdir(cons_dir):
                if filename.startswith(f"{consuntivo_num}-") or filename.startswith(f"{consuntivo_num} "):
                    return os.path.join(cons_dir, filename)
            return f"File non trovato per il n° {consuntivo_num}"
        except Exception as e:
            self.logger(f"Errore ricerca consuntivo n°{consuntivo_num}: {e}", "ERROR")
            return "Errore ricerca file"

    def find_consuntivo_for_tcl(self, year, month_name, tcl_name, cancel_event):
        if not year or not month_name: return None, "Periodo non selezionato"
        cons_dir = os.path.join(self.config.CANONI_CONSUNTIVI_BASE_DIR, year, "CONSUNTIVI", year)
        if not os.path.isdir(cons_dir): return None, f"Cartella non trovata: {cons_dir}"
        try:
            files_in_dir = os.listdir(cons_dir)
            month_norm = month_name.upper()
            tcl_norm = tcl_name.upper()
            for filename in files_in_dir:
                if cancel_event.is_set(): return None, "Annullato"
                filename_norm = filename.upper()
                if all(keyword in filename_norm for keyword in ["CANONE", month_norm, tcl_norm]):
                    match = re.match(r'^(\d+)', filename)
                    if match:
                        number = match.group(1)
                        self.logger(f"Trovato file '{filename}' per {tcl_name}, numero: {number}", "SUCCESS")
                        return number, os.path.join(cons_dir, filename)
            self.logger(f"Nessun file consuntivo trovato per {tcl_name} nel periodo {month_name} {year}", "WARNING")
            return None, "File non trovato"
        except Exception as e:
            self.logger(f"Errore durante la ricerca del file per {tcl_name}: {e}", "ERROR")
            return None, "Errore di sistema"

    def run_printing_process(self, cancel_event, paths_to_print, printer_name, macro_name):
        self.logger("Avvio del processo di stampa canoni...", "HEADER")
        try:
            if not self._validate_paths(paths_to_print, printer_name, macro_name): return
            if cancel_event.is_set(): return

            with ExcelHandler(self.logger) as excel_app, WordHandler(self.logger) as word_app:
                if not excel_app or not word_app: return
                if cancel_event.is_set(): return

                word_app.ActivePrinter = printer_name
                self.logger(f"Stampante attiva impostata su: '{printer_name}'", "SUCCESS")

                giornaliera_path = paths_to_print["giornaliera"]
                cons_paths = paths_to_print["consuntivi"]
                word_path = paths_to_print["word"]

                wb_giornaliera = excel_app.Workbooks.Open(giornaliera_path)
                wb_cons_list = [excel_app.Workbooks.Open(p) for p in cons_paths]
                doc_word = word_app.Documents.Open(word_path)
                self.logger("Documenti aperti.", 'INFO')

                for i, cons_wb in enumerate(wb_cons_list):
                    if cancel_event.is_set(): break

                    leaf_name = cons_wb.Name
                    self.logger(f"Esecuzione macro '{macro_name}' su {leaf_name}...", 'INFO')
                    excel_app.Run(f"'{leaf_name}'!{macro_name}")
                    self.logger(f"Macro su Consuntivo {i+1} completata.", 'SUCCESS')

                    if i < len(wb_cons_list) - 1:
                        if cancel_event.is_set(): break
                        self.logger(f"Stampa file Word: {doc_word.Name}...", 'INFO')
                        doc_word.PrintOut()
                        self.logger("Comando di stampa Word inviato.", 'SUCCESS')

                doc_word.Close(SaveChanges=0)
                for wb in wb_cons_list: wb.Close(SaveChanges=False)
                wb_giornaliera.Close(SaveChanges=False)

            if not cancel_event.is_set():
                self.logger("--- PROCESSO STAMPA CANONI COMPLETATO ---", 'SUCCESS')
        except Exception as e:
            self.logger(f"ERRORE CRITICO nel processo: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            if cancel_event.is_set(): self.logger("Processo di stampa annullato.", "WARNING")
            self.gui.after(0, self.gui.on_process_finished)

    def _validate_paths(self, paths, printer, macro):
        all_paths = {"File Giornaliera": paths["giornaliera"], "File Foglio Canone": paths["word"]}
        for i, p in enumerate(paths["consuntivi"]): all_paths[f"Canone {i+1}"] = p
        for name, path in all_paths.items():
            if not path or not os.path.isfile(path):
                self.logger(f"ERRORE: Percorso per '{name}' non valido o file non trovato: '{path}'", 'ERROR')
                return False
        if not macro.strip(): self.logger("ERRORE: Nome della macro VBA non specificato.", 'ERROR'); return False
        if not printer: self.logger("ERRORE: Nessuna stampante selezionata.", 'ERROR'); return False
        return True
