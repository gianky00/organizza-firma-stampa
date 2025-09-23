import os
import re
import shutil
import traceback
from src.utils.excel_handler import ExcelHandler
from src.logic.monthly_fees import MonthlyFeesProcessor

class OrganizationProcessor:
    """
    Handles organizing Excel files by ODC and batch printing them.
    """
    def __init__(self, gui, config, setup_progress_cb, update_progress_cb, hide_progress_cb):
        self.gui = gui
        self.config = config
        self.logger = gui.log_organizza
        self.setup_progress = setup_progress_cb
        self.update_progress = update_progress_cb
        self.hide_progress = hide_progress_cb

        self.stampa_processing_data = {
            "schedacontrolloSTRUMENTIANALOGICI": {"PrintArea": "A2:N55"},
            "schedacontrolloSTRUMENTIDIGITALI": {"PrintArea": "A2:N50"},
            "SchedacontrolloREPORTMANUTENZIONECORRETTIVA": {"PrintArea": "A2:N55"},
            "SCHEDAMANUTENZIONE": {"PrintArea": "A1:FV106"}
        }

    def run_organization_process(self):
        """
        Main entry point for organizing files.
        """
        self.logger("Avvio del processo di organizzazione...", "HEADER")
        try:
            self.logger("Pulizia della cartella di destinazione...", "INFO")
            self._clear_folder_content(
                self.config.organizza_dest_dir.get(),
                self.config.ORGANIZZA_DEST_DIR,
                self.logger
            )
            self._organize_files()
            self.logger("Organizzazione completata. Aggiornamento della lista file...", "SUCCESS")
            self.gui.after(0, self.gui.populate_stampa_list)
        except Exception as e:
            self.logger(f"ERRORE CRITICO E IMPREVISTO durante l'organizzazione: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            self.gui.after(0, self.hide_progress)
            self.gui.after(0, self.gui.toggle_organizza_buttons, 'normal')

    def run_printing_process(self, folders_to_print):
        """
        Main entry point for printing files from selected folders.
        """
        if not folders_to_print:
            self.logger("Nessuna cartella selezionata per la stampa.", "WARNING")
            self.gui.after(0, self.gui.toggle_organizza_buttons, 'normal')
            return

        try:
            self.logger(f"--- Avvio Stampa per {len(folders_to_print)} cartelle ---", "HEADER")
            self._print_files_in_folders(folders_to_print)
            self.logger("--- Stampa Completata ---", "SUCCESS")
        except Exception as e:
            self.logger(f"ERRORE CRITICO durante la stampa: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            self.gui.after(0, self.hide_progress)
            self.gui.after(0, self.gui.toggle_organizza_buttons, 'normal')

    def _organize_files(self):
        source_dir = self.config.organizza_source_dir.get()
        dest_dir = self.config.organizza_dest_dir.get()

        self.logger("--- Inizio Organizzazione ---", "HEADER")
        if not os.path.isdir(source_dir):
            self.logger(f"ERRORE: La cartella di origine '{source_dir}' non esiste.", "ERROR")
            return

        excel_ext = ('.xls', '.xlsx', '.xlsm', '.xlsb')
        try:
            excel_files = [os.path.join(r, f) for r, _, fs in os.walk(source_dir) for f in fs if f.lower().endswith(excel_ext) and not f.startswith('~')]
        except Exception as e:
            self.logger(f"ERRORE durante l'accesso alla cartella di origine '{source_dir}': {e}", "ERROR")
            return

        if not excel_files:
            self.logger(f"Nessun file Excel trovato in: {source_dir}", "WARNING")
            return

        num_files = len(excel_files)
        self.logger(f"Trovati {num_files} file Excel da analizzare.")
        self.gui.after(0, self.setup_progress, num_files, "Organizzazione in corso:")
        summary = {"processed": 0, "errors": []}

        with ExcelHandler(self.logger) as excel:
            if not excel:
                return

            for i, fp in enumerate(excel_files):
                self.gui.after(0, self.update_progress, i + 1)
                self.logger(f"Processando: {os.path.basename(fp)}")
                wb = None
                try:
                    wb = excel.Workbooks.Open(fp)
                    ws = wb.Worksheets(1)

                    odc_v = next((ws.Range(c).Value for c in ["L50", "L45", "DB14", "DB17"] if ws.Range(c).Value is not None and str(ws.Range(c).Value).strip() != ""), None)
                    odc_s = str(int(odc_v)) if isinstance(odc_v, (int, float)) else (str(odc_v).strip() if isinstance(odc_v, str) else "")

                    wb.Close(SaveChanges=False)
                    wb = None

                    dest_folder_name = re.sub(r'[\\/:*?"<>|]', '', odc_s) if odc_s and odc_s.upper() != "NA" else "Schede senza ODC"
                    dest_folder_path = os.path.join(dest_dir, dest_folder_name)

                    os.makedirs(dest_folder_path, exist_ok=True)
                    shutil.copy2(fp, dest_folder_path)
                    self.logger(f"  -> Copiato in: {dest_folder_name}", "SUCCESS")
                    summary["processed"] += 1

                except Exception as e:
                    error_msg = f"Impossibile analizzare o copiare il file. Dettagli: {e}"
                    self.logger(f"ERRORE: {error_msg}", "ERROR")
                    summary["errors"].append((os.path.basename(fp), error_msg))
                finally:
                    if wb:
                        wb.Close(SaveChanges=False)

        self.logger(f"\n--- RIEPILOGO ORGANIZZAZIONE ---", "HEADER")
        self.logger(f"File processati con successo: {summary['processed']}/{num_files}", "SUCCESS")
        if summary['errors']:
            self.logger(f"File con errori: {len(summary['errors'])}", "ERROR")
            self.logger("--- DETTAGLIO ERRORI ---", "HEADER")
            for file_name, error_msg in summary['errors']:
                self.logger(f"- {file_name}: {error_msg}", "ERROR")

    def _print_files_in_folders(self, folder_list):
        excel_ext = ('.xls', '.xlsx', '.xlsm', '.xlsb')
        num_folders = len(folder_list)
        self.gui.after(0, self.setup_progress, num_folders, "Stampa in corso:")

        with ExcelHandler(self.logger) as excel:
            if not excel:
                return

            errors = []
            for i, folder_p in enumerate(folder_list):
                self.gui.after(0, self.update_progress, i + 1)
                self.logger(f"Stampa cartella: {os.path.basename(folder_p)}")
                try:
                    excel_fs = [os.path.join(folder_p, f) for f in os.listdir(folder_p) if f.lower().endswith(excel_ext) and not f.startswith('~')]
                    if not excel_fs:
                        self.logger("  -> Nessun file Excel trovato in questa cartella.", "WARNING")
                        continue

                    for fp in excel_fs:
                        wb = None
                        try:
                            wb = excel.Workbooks.Open(fp)
                            ws = wb.Worksheets(1)

                            m_val = next((str(ws.Cells(r, c).Value).strip() for r, c in [(2, 5), (2, 20), (5, 20)] if ws.Cells(r, c).Value and str(ws.Cells(r, c).Value).strip()), "")
                            cleaned_model = re.sub(r'\W', '', m_val)

                            if cleaned_model in self.stampa_processing_data:
                                ws.PageSetup.PrintArea = self.stampa_processing_data[cleaned_model]["PrintArea"]
                                wb.PrintOut()
                                self.logger(f"  -> Stampa inviata per: {os.path.basename(fp)}", "SUCCESS")
                            else:
                                self.logger(f"  -> Ignorato (modello non trovato '{cleaned_model}'): {os.path.basename(fp)}", "WARNING")
                        except Exception as e_file:
                            error_msg = f"Impossibile stampare il file. Dettagli: {e_file}"
                            self.logger(f"ERRORE: {error_msg}", "ERROR")
                            errors.append((os.path.basename(fp), error_msg))
                        finally:
                            if wb:
                                wb.Close(SaveChanges=False)
                except Exception as e_folder:
                    error_msg = f"Impossibile elaborare la cartella. Dettagli: {e_folder}"
                    self.logger(f"ERRORE: {error_msg}", "ERROR")
                    errors.append((os.path.basename(folder_p), error_msg))

            if errors:
                self.logger("\n--- RIEPILOGO ERRORI DI STAMPA ---", "HEADER")
                for item_name, error_msg in errors:
                    self.logger(f"- {item_name}: {error_msg}", "ERROR")

    def get_odc_to_canone_map(self, year, month):
        """
        Reads the 'Giornaliera' file for the given period and returns a dictionary
        mapping ODC numbers to canone names (e.g., {'5400...': 'canone messina'}).
        """
        self.logger(f"Lettura del file Giornaliera per {month} {year} per mappare gli ODC...", "INFO")

        # Build the path to the 'Giornaliera' file directly to avoid incorrect dependencies.
        if not year or not month:
            giornaliera_path = ""
        else:
            month_number = self.config.mesi_giornaliera_map.get(month)
            if not month_number:
                giornaliera_path = ""
            else:
                year_folder_name = f"Giornaliere {year}"
                file_name = f"Giornaliera {month_number}-{year}.xlsm"
                giornaliera_path = os.path.join(self.config.CANONI_GIORNALIERA_BASE_DIR, year_folder_name, file_name)

        if not os.path.isfile(giornaliera_path):
            self.logger(f"File Giornaliera non trovato al percorso: {giornaliera_path}", "WARNING")
            return {}

        mapping = {}
        with ExcelHandler(self.logger) as excel:
            if not excel:
                return {}

            wb = None
            try:
                wb = excel.Workbooks.Open(giornaliera_path, ReadOnly=True)
                ws = wb.Worksheets("RIEPILOGO")

                cells_to_check = [("S16", "S17"), ("U16", "U17"), ("V16", "V17")]
                for header_cell, value_cell in cells_to_check:
                    header = ws.Range(header_cell).Value
                    value_raw = ws.Range(value_cell).Value

                    if header and value_raw:
                        # Clean the value, e.g., "5400261236\ncanone" -> "5400261236"
                        odc_num = str(value_raw).split('\n')[0].strip()
                        if odc_num.isdigit():
                            mapping[odc_num] = str(header).lower()

            except Exception as e:
                self.logger(f"Errore durante la lettura del file Giornaliera: {e}", "ERROR")
            finally:
                if wb:
                    wb.Close(SaveChanges=False)

        self.logger(f"Mappa ODC creata con {len(mapping)} voci.", "INFO")
        return mapping

