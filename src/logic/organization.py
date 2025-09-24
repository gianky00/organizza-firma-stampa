import os
import re
import shutil
import traceback
import time
from src.utils.excel_handler import ExcelHandler
from src.utils.file_utils import clear_folder_content

class OrganizationProcessor:
    def __init__(self, gui, config, fees_processor, setup_progress_cb, update_progress_cb, hide_progress_cb):
        self.gui = gui
        self.config = config
        self.fees_processor = fees_processor
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

    def run_organization_process(self, cancel_event):
        self.logger("Avvio del processo di organizzazione...", "HEADER")
        dest_dir = self.config.organizza_dest_dir.get()
        backup_dir = ""
        operation_successful = False
        try:
            try:
                if os.path.isdir(dest_dir) and os.listdir(dest_dir):
                    timestamp = time.strftime("%Y%m%d-%H%M%S")
                    backup_dir = f"{dest_dir}_backup_{timestamp}"
                    self.logger(f"Creazione backup: {os.path.basename(backup_dir)}", "INFO")
                    shutil.copytree(dest_dir, backup_dir)
            except Exception as e:
                self.logger(f"ERRORE CRITICO durante la creazione del backup: {e}", "ERROR")
                self.logger("L'operazione di organizzazione è stata interrotta per prevenire la perdita di dati.", "ERROR")
                return # Abort the entire operation if backup fails

            clear_folder_content(dest_dir, self.logger, folder_display_name=self.config.ORGANIZZA_DEST_DIR)
            os.makedirs(dest_dir, exist_ok=True)
            if cancel_event.is_set(): return

            self._organize_files(cancel_event)

            if not cancel_event.is_set():
                operation_successful = True
                self.logger("Organizzazione completata.", "SUCCESS")
                self.gui.after(0, self.gui.populate_stampa_list)
        except Exception as e:
            self.logger(f"ERRORE CRITICO: {e}", "ERROR"); self.logger(traceback.format_exc(), "ERROR")
            operation_successful = False
        finally:
            if operation_successful:
                if backup_dir: shutil.rmtree(backup_dir)
            else:
                self.logger("ANNULLAMENTO/ERRORE: Ripristino cartella dal backup.", "WARNING")
                if backup_dir and os.path.isdir(backup_dir):
                    clear_folder_content(dest_dir, self.logger)
                    shutil.rmtree(dest_dir)
                    os.rename(backup_dir, dest_dir)
                    self.logger("Ripristino completato.", "SUCCESS")

            if cancel_event.is_set(): self.logger("Processo annullato.", "WARNING")
            self.gui.after(0, self.hide_progress)
            self.gui.after(0, self.gui.on_process_finished)

    def run_printing_process(self, cancel_event, folders_to_print):
        if not folders_to_print:
            self.logger("Nessuna cartella selezionata.", "WARNING")
            self.gui.after(0, self.gui.on_process_finished)
            return
        try:
            self.logger(f"--- Avvio Stampa per {len(folders_to_print)} cartelle ---", "HEADER")
            self._print_files_in_folders(cancel_event, folders_to_print)
            if not cancel_event.is_set(): self.logger("--- Stampa Completata ---", "SUCCESS")
        except Exception as e:
            self.logger(f"ERRORE CRITICO: {e}", "ERROR"); self.logger(traceback.format_exc(), "ERROR")
        finally:
            if cancel_event.is_set(): self.logger("Processo di stampa annullato.", "WARNING")
            self.gui.after(0, self.hide_progress)
            self.gui.after(0, self.gui.on_process_finished)

    def _organize_files(self, cancel_event):
        source_dir = self.config.organizza_source_dir.get(); dest_dir = self.config.organizza_dest_dir.get()
        if not os.path.isdir(source_dir): self.logger(f"ERRORE: Cartella di origine non trovata.", "ERROR"); return
        try:
            excel_files = [os.path.join(r, f) for r, _, fs in os.walk(source_dir) for f in fs if f.lower().endswith(('.xls', '.xlsx', '.xlsm', '.xlsb')) and not f.startswith('~')]
        except Exception as e:
            self.logger(f"ERRORE accesso cartella di origine: {e}", "ERROR"); return
        if not excel_files: self.logger(f"Nessun file Excel trovato.", "WARNING"); return

        self.gui.after(0, self.setup_progress, len(excel_files), "Organizzazione in corso:")
        summary = {"processed": 0, "errors": []}
        with ExcelHandler(self.logger) as excel:
            if not excel: return
            for i, fp in enumerate(excel_files):
                if cancel_event.is_set(): return
                self.gui.after(0, self.update_progress, i + 1)
                self.logger(f"Processando: {os.path.basename(fp)}...")
                wb = None
                try:
                    wb = excel.Workbooks.Open(fp)
                    ws = wb.Worksheets(1)
                    odc_v = next((ws.Range(c).Value for c in ["L50", "L45", "DB14", "DB17"] if ws.Range(c).Value is not None and str(ws.Range(c).Value).strip() != ""), None)
                    odc_s = str(int(odc_v)) if isinstance(odc_v, (int, float)) else (str(odc_v).strip() if isinstance(odc_v, str) else "")
                    wb.Close(SaveChanges=False); wb = None
                    dest_folder_name = re.sub(r'[\\/:*?"<>|]', '', odc_s) if odc_s and odc_s.upper() != "NA" else "Schede senza ODC"
                    dest_folder_path = os.path.join(dest_dir, dest_folder_name)
                    os.makedirs(dest_folder_path, exist_ok=True)
                    shutil.copy2(fp, dest_folder_path)
                    summary["processed"] += 1
                except Exception as e:
                    summary["errors"].append((os.path.basename(fp), f"Dettagli: {e}"))
                finally:
                    if wb: wb.Close(SaveChanges=False)
        # ... (summary logging)

    def _print_files_in_folders(self, cancel_event, folder_list):
        self.gui.after(0, self.setup_progress, len(folder_list), "Stampa in corso:")
        with ExcelHandler(self.logger) as excel:
            if not excel: return
            errors = []
            for i, folder_p in enumerate(folder_list):
                if cancel_event.is_set(): return
                self.gui.after(0, self.update_progress, i + 1)
                self.logger(f"Stampa cartella: {os.path.basename(folder_p)}")
                try:
                    excel_fs = [os.path.join(folder_p, f) for f in os.listdir(folder_p) if f.lower().endswith(('.xls', '.xlsx', '.xlsm', '.xlsb')) and not f.startswith('~')]
                    if not excel_fs: self.logger("  -> Nessun file Excel trovato.", "WARNING"); continue
                    for fp in excel_fs:
                        if cancel_event.is_set(): return
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
                            else: self.logger(f"  -> Ignorato (modello non trovato): {os.path.basename(fp)}", "WARNING")
                        except Exception as e_file: errors.append((os.path.basename(fp), f"Dettagli: {e_file}"))
                        finally:
                            if wb: wb.Close(SaveChanges=False)
                except Exception as e_folder: errors.append((os.path.basename(folder_p), f"Dettagli: {e_folder}"))
        # ... (error summary logging)

    def get_odc_to_canone_map(self, year, month):
        self.logger(f"Lettura del file Giornaliera per {month} {year}...", "INFO")
        giornaliera_path = self.fees_processor.get_giornaliera_path(year, month)

        if not os.path.isfile(giornaliera_path):
            self.logger(f"File Giornaliera non trovato: {giornaliera_path}", "WARNING")
            return {}

        mapping = {}
        with ExcelHandler(self.logger) as excel:
            if not excel: return {}
            wb = None
            try:
                wb = excel.Workbooks.Open(giornaliera_path, ReadOnly=True)
                ws = wb.Worksheets("RIEPILOGO")
                cells_to_check = [("S16", "S17"), ("U16", "U17"), ("V16", "V17")]
                for header_cell, value_cell in cells_to_check:
                    header = ws.Range(header_cell).Value; value_raw = ws.Range(value_cell).Value
                    if header and value_raw:
                        odc_num = str(value_raw).split('\n')[0].strip()
                        if odc_num.isdigit(): mapping[odc_num] = str(header).lower()
            except Exception as e: self.logger(f"Errore lettura Giornaliera: {e}", "ERROR")
            finally:
                if wb: wb.Close(SaveChanges=False)
        self.logger(f"Mappa ODC creata con {len(mapping)} voci.", "INFO")
        return mapping
