import os
import pythoncom
import win32com.client
import win32print
import traceback
from datetime import datetime

class MonthlyFeesProcessor:
    """
    Handles the complex workflow for printing monthly fee documents.
    """
    def __init__(self, gui, config):
        self.gui = gui
        self.config = config
        self.logger = gui.log_canoni

    def get_printers(self):
        """Returns a list of available printer names."""
        try:
            printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            default_printer = win32print.GetDefaultPrinter()
            return printers, default_printer
        except Exception as e:
            self.logger(f"Errore nel caricamento delle stampanti: {e}", "ERROR")
            return [], None

    def get_giornaliera_path(self, year, month_name):
        """Constructs and returns the path for the 'Giornaliera' file."""
        if not year or not month_name:
            return "Seleziona Anno e Mese"

        month_number = self.config.mesi_giornaliera_map.get(month_name)
        if not month_number:
            return "Mese non valido"

        year_folder_name = f"Giornaliere {year}"
        file_name = f"Giornaliera {month_number}-{year}.xlsm"
        return os.path.join(self.config.CANONI_GIORNALIERA_BASE_DIR, year_folder_name, file_name)

    def get_consuntivo_path(self, year, consuntivo_num):
        """Finds and returns the path for a 'Consuntivo' file based on its number."""
        if not year:
            return "Anno non selezionato"
        if not consuntivo_num.strip().isdigit():
            return "Inserire un numero valido"

        cons_dir = os.path.join(self.config.CANONI_CONSUNTIVI_BASE_DIR, year, "CONSUNTIVI", year)
        if not os.path.isdir(cons_dir):
            return f"ERRORE: Cartella non trovata"

        try:
            files_in_dir = os.listdir(cons_dir)
            for filename in files_in_dir:
                # Check for formats like "183-..." or "183 ..."
                if filename.startswith(f"{consuntivo_num}-") or filename.startswith(f"{consuntivo_num} "):
                    return os.path.join(cons_dir, filename)
            return f"File non trovato per il n° {consuntivo_num}"
        except Exception as e:
            self.logger(f"Errore ricerca consuntivo n°{consuntivo_num}: {e}", "ERROR")
            return "Errore ricerca file"

    def run_printing_process(self, paths_to_print, printer_name, macro_name):
        """
        Main entry point for the printing process.
        """
        pythoncom.CoInitialize()
        excel_app, word_app = None, None
        docs_to_close = []
        try:
            if not self._validate_paths(paths_to_print, printer_name, macro_name):
                return

            self.logger("Avvio applicazioni Excel e Word...", 'INFO')
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.DisplayAlerts = False
            word_app = win32com.client.Dispatch("Word.Application")

            word_app.ActivePrinter = printer_name
            self.logger(f"Stampante impostata su: '{printer_name}'", "SUCCESS")

            self.logger("--- Apertura file sorgente ---", 'HEADER')
            giornaliera_path = paths_to_print["giornaliera"]
            wb_giornaliera = excel_app.Workbooks.Open(giornaliera_path)
            docs_to_close.append(wb_giornaliera)
            self.logger(f"File Giornaliera aperto: {os.path.basename(giornaliera_path)}", 'SUCCESS')

            cons_paths = paths_to_print["consuntivi"]
            word_path = paths_to_print["word"]
            wb_cons_list = []

            for i, p in enumerate(cons_paths):
                wb = excel_app.Workbooks.Open(p)
                wb_cons_list.append(wb)
                docs_to_close.append(wb)
                self.logger(f"Aperto Consuntivo {i+1}: {os.path.basename(p)}", 'INFO')

            doc_word = word_app.Documents.Open(word_path)
            docs_to_close.append(doc_word)
            self.logger(f"Aperto documento Word: {os.path.basename(word_path)}", 'INFO')

            self.logger("--- Inizio sequenza operazioni ---", 'HEADER')
            for i, cons_wb in enumerate(wb_cons_list):
                leaf_name = cons_wb.Name
                self.logger(f"Esecuzione macro '{macro_name}' su {leaf_name}...", 'INFO')
                excel_app.Run(f"'{leaf_name}'!{macro_name}")
                self.logger(f"Macro su Consuntivo {i+1} completata.", 'SUCCESS')

                # Print after every macro run except the last one
                if i < len(wb_cons_list) - 1:
                    self.logger(f"Stampa file Word: {doc_word.Name}...", 'INFO')
                    doc_word.PrintOut()
                    self.logger("Comando di stampa Word inviato.", 'SUCCESS')

            self.logger("--- PROCESSO STAMPA CANONI COMPLETATO ---", 'SUCCESS')

        except Exception as e:
            self.logger(f"ERRORE CRITICO nel processo: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
            self.logger("Il processo è stato interrotto.", "WARNING")
        finally:
            self.logger("Chiusura file e applicazioni...", 'INFO')
            for doc in reversed(docs_to_close):
                try:
                    doc.Close(SaveChanges=0)
                except Exception as e_cl:
                    self.logger(f"Errore durante la chiusura di un documento: {e_cl}", "WARNING")
            if word_app:
                word_app.Quit()
                self.logger("Applicazione Word chiusa.", 'INFO')
            if excel_app:
                excel_app.Quit()
                self.logger("Applicazione Excel chiusa.", 'INFO')

            pythoncom.CoUninitialize()
            self.gui.after(0, self.gui.toggle_stampa_canoni_buttons, 'normal')


    def _validate_paths(self, paths, printer, macro):
        """Validates all paths required for the process."""
        all_paths = {
            "File Giornaliera": paths["giornaliera"],
            "File Foglio Canone": paths["word"],
        }
        for i, p in enumerate(paths["consuntivi"]):
            all_paths[f"Canone {i+1}"] = p

        for name, path in all_paths.items():
            if not path or not os.path.isfile(path):
                self.logger(f"ERRORE: Percorso per '{name}' non valido o file non trovato: '{path}'", 'ERROR')
                return False

        if not macro.strip():
            self.logger("ERRORE: Nome della macro VBA non specificato.", 'ERROR')
            return False

        if not printer:
            self.logger("ERRORE: Nessuna stampante selezionata.", 'ERROR')
            return False

        return True
