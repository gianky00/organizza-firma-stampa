import os
import re
import traceback
from pathlib import Path

from src.utils.excel_gateway import ExcelGateway
from src.utils.word_gateway import WordGateway


class MonthlyFeesProcessor:
    def __init__(
        self,
        gui,
        app_config,
        excel_gateway_class=None,
        word_gateway_class=None,
    ):
        self.gui = gui
        self.app_config = app_config
        self.logger = gui.log_canoni
        self.excel_gateway = (excel_gateway_class or ExcelGateway)(self.logger)
        self.word_gateway = (word_gateway_class or WordGateway)(self.logger)

    def get_printers(self):
        try:
            import win32print

            printers = [
                p[2]
                for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            ]
            default_printer = win32print.GetDefaultPrinter()
            return printers, default_printer
        except Exception as e:
            self.logger(f"Errore nel caricamento delle stampanti: {e}", "ERROR")
            return [], None

    def get_giornaliera_path(self, year, month_name):
        if not year or not month_name:
            return ""
        month_number = self.app_config.mesi_giornaliera_map.get(month_name)
        if not month_number:
            return ""
        year_folder_name = f"Giornaliere {year}"
        file_name = f"Giornaliera {month_number}-{year}.xlsm"
        base_dir = self.app_config.canoni_giornaliera_base_dir.get()
        return os.path.join(base_dir, year_folder_name, file_name)

    def get_consuntivo_path(self, year, consuntivo_num):
        if not year:
            return ""
        if not consuntivo_num.strip().isdigit():
            return ""
        base_dir = self.app_config.canoni_consuntivi_base_dir.get()
        cons_dir = os.path.join(base_dir, year, "CONSUNTIVI", year)
        if not Path(cons_dir).is_dir():
            return ""
        try:
            for filename in os.listdir(cons_dir):
                if filename.startswith((f"{consuntivo_num}-", f"{consuntivo_num} ")):
                    return os.path.join(cons_dir, filename)
            return ""
        except Exception as e:
            self.logger(f"Errore ricerca consuntivo n°{consuntivo_num}: {e}", "ERROR")
            return ""

    def find_consuntivo_for_tcl(self, year, month_name, tcl_name, cancel_event):
        if not year or not month_name:
            return None, "Periodo non selezionato"
        base_dir = self.app_config.canoni_consuntivi_base_dir.get()
        cons_dir = os.path.join(base_dir, year, "CONSUNTIVI", year)
        if not Path(cons_dir).is_dir():
            return None, f"Cartella non trovata: {cons_dir}"
        try:
            files_in_dir = os.listdir(cons_dir)
            month_norm = month_name.upper()
            tcl_norm = tcl_name.upper()
            for filename in files_in_dir:
                if cancel_event.is_set():
                    return None, "Annullato"
                filename_norm = filename.upper()
                if all(keyword in filename_norm for keyword in ("CANONE", month_norm, tcl_norm)):
                    match = re.match(r"^(\d+)", filename)
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
            printers, _ = self.get_printers()
            if printer_name not in printers:
                self.logger(f"ERRORE: Stampante '{printer_name}' non trovata nel sistema.", "ERROR")
                return

            if not self._validate_paths(paths_to_print, printer_name, macro_name) or cancel_event.is_set():
                return

            # Batch execution using raw handlers via Gateways for optimization
            from src.utils.excel_handler import ExcelHandler
            from src.utils.word_handler import WordHandler

            excel_h_class = getattr(self.excel_gateway, "excel_handler_class", ExcelHandler)
            word_h_class = getattr(self.word_gateway, "word_handler_class", WordHandler)

            with excel_h_class(self.logger) as excel_app, word_h_class(self.logger) as word_app:
                if not excel_app or not word_app or cancel_event.is_set():
                    return

                word_app.ActivePrinter = printer_name
                self.logger(f"Stampante attiva impostata su: '{printer_name}'", "SUCCESS")

                enabled_consuntivi = [c for c in paths_to_print["consuntivi"] if c["print"]]
                if not enabled_consuntivi:
                    self.logger("Nessun canone selezionato per la stampa.", "WARNING")
                    return

                self.logger(f"Canoni da stampare: {', '.join(c['name'] for c in enabled_consuntivi)}", "INFO")
                self._execute_batch_print(
                    excel_app, word_app, paths_to_print, enabled_consuntivi, macro_name, cancel_event
                )

            if not cancel_event.is_set():
                self.logger("--- PROCESSO STAMPA CANONI COMPLETATO ---", "SUCCESS")
        except Exception as e:
            self.logger(f"ERRORE CRITICO: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            if cancel_event.is_set():
                self.logger("Processo annullato.", "WARNING")

    def _execute_batch_print(self, excel, word, paths, consuntivi, macro, cancel_event):
        wb_giornaliera = excel.Workbooks.Open(paths["giornaliera"])
        if wb_giornaliera is None:
            self.logger("Impossibile aprire il file Giornaliera.", "ERROR")
            return

        doc_word = word.Documents.Open(paths["word"])
        if doc_word is None:
            self.logger("Impossibile aprire il file Word.", "ERROR")
            wb_giornaliera.Close(SaveChanges=False)
            return

        try:
            for i, cons_info in enumerate(consuntivi):
                if cancel_event.is_set():
                    break

                wb_cons = excel.Workbooks.Open(cons_info["path"])
                if wb_cons is None:
                    self.logger(f"Impossibile aprire il consuntivo: {cons_info['name']}", "ERROR")
                    continue
                try:
                    leaf_name = wb_cons.Name
                    self.logger(f"Esecuzione macro '{macro}' su {leaf_name} ({cons_info['name']})...", "INFO")
                    excel.Run(f"'{leaf_name}'!{macro}")

                    if i < len(consuntivi) - 1 and not cancel_event.is_set():
                        self.logger(f"Stampa file Word: {doc_word.Name}...", "INFO")
                        doc_word.PrintOut()
                finally:
                    wb_cons.Close(SaveChanges=False)
        finally:
            try:
                doc_word.Close(SaveChanges=0)
            except Exception as e:
                self.logger(f"Errore chiusura documento Word: {e}", "WARNING")
            wb_giornaliera.Close(SaveChanges=False)

    def _validate_paths(self, paths, printer, macro):
        if not printer or not macro:
            self.logger("ERRORE: Stampante o Macro non specificata.", "ERROR")
            return False
        if not os.path.isfile(paths["giornaliera"]) or not os.path.isfile(paths["word"]):
            self.logger("ERRORE: File Giornaliera o Word non trovato.", "ERROR")
            return False
        return True
