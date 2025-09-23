import os
import re
import shutil
import pythoncom
import win32com.client
import traceback

class OrganizationProcessor:
    """
    Handles organizing Excel files by ODC and batch printing them.
    """
    def __init__(self, gui, config):
        self.gui = gui
        self.config = config
        self.logger = gui.log_organizza  # Assumes a logger for the organization tab

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
        pythoncom.CoInitialize()
        try:
            # The destination folder is cleared before organizing
            self._clear_folder_content(
                self.config.organizza_dest_dir.get(),
                self.config.ORGANIZZA_DEST_DIR,
                self.logger
            )
            self._organize_files()
            # After organizing, refresh the list in the GUI
            self.gui.after(0, self.gui.populate_stampa_list)
        except Exception as e:
            self.logger(f"ERRORE CRITICO durante l'organizzazione: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            pythoncom.CoUninitialize()
            self.gui.after(0, self.gui.toggle_organizza_buttons, 'normal')

    def run_printing_process(self, folders_to_print):
        """
        Main entry point for printing files from selected folders.
        """
        if not folders_to_print:
            self.logger("Nessuna cartella selezionata per la stampa.", "WARNING")
            self.gui.after(0, self.gui.toggle_organizza_buttons, 'normal')
            return

        pythoncom.CoInitialize()
        try:
            self.logger(f"--- Avvio Stampa per {len(folders_to_print)} cartelle ---", "HEADER")
            self._print_files_in_folders(folders_to_print)
            self.logger("--- Stampa Completata ---", "SUCCESS")
        except Exception as e:
            self.logger(f"ERRORE CRITICO durante la stampa: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            pythoncom.CoUninitialize()
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

        self.logger(f"Trovati {len(excel_files)} file Excel da analizzare.")
        excel, proc_count = None, 0
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            for fp in excel_files:
                self.logger(f"Processando: {os.path.basename(fp)}")
                wb = None
                try:
                    wb = excel.Workbooks.Open(fp)
                    ws = wb.Worksheets(1)

                    # Search for ODC value in specified cells
                    odc_v = next((ws.Range(c).Value for c in ["L50", "L45", "DB14", "DB17"] if ws.Range(c).Value is not None and str(ws.Range(c).Value).strip() != ""), None)
                    odc_s = str(int(odc_v)) if isinstance(odc_v, (int, float)) else (str(odc_v).strip() if isinstance(odc_v, str) else "")

                    wb.Close(SaveChanges=False)
                    wb = None

                    # Determine destination folder name
                    dest_folder_name = re.sub(r'[\\/:*?"<>|]', '', odc_s) if odc_s and odc_s.upper() != "NA" else "Schede senza ODC"
                    dest_folder_path = os.path.join(dest_dir, dest_folder_name)

                    os.makedirs(dest_folder_path, exist_ok=True)
                    shutil.copy2(fp, dest_folder_path)
                    self.logger(f"  -> Copiato in: {dest_folder_name}", "SUCCESS")
                    proc_count += 1

                except Exception as e:
                    self.logger(f"ERRORE durante l'analisi del file {os.path.basename(fp)}: {e}", "ERROR")
                finally:
                    if wb:
                        wb.Close(SaveChanges=False)
        finally:
            if excel:
                excel.Quit()
            self.logger(f"--- Organizzazione Completata ({proc_count}/{len(excel_files)}) ---", "HEADER")

    def _print_files_in_folders(self, folder_list):
        excel_ext = ('.xls', '.xlsx', '.xlsm', '.xlsb')
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            for folder_p in folder_list:
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

                            # Find model to determine print area
                            m_val = next((str(ws.Cells(r, c).Value).strip() for r, c in [(2, 5), (2, 20), (5, 20)] if ws.Cells(r, c).Value and str(ws.Cells(r, c).Value).strip()), "")
                            cleaned_model = re.sub(r'\W', '', m_val)

                            if cleaned_model in self.stampa_processing_data:
                                ws.PageSetup.PrintArea = self.stampa_processing_data[cleaned_model]["PrintArea"]
                                wb.PrintOut()
                                self.logger(f"  -> Stampa inviata per: {os.path.basename(fp)}", "SUCCESS")
                            else:
                                self.logger(f"  -> Ignorato (modello non trovato '{cleaned_model}'): {os.path.basename(fp)}", "WARNING")
                        except Exception as e_file:
                            self.logger(f"ERRORE durante la stampa del file {os.path.basename(fp)}: {e_file}", "ERROR")
                        finally:
                            if wb:
                                wb.Close(SaveChanges=False)
                except Exception as e_folder:
                    self.logger(f"ERRORE durante l'elaborazione della cartella {os.path.basename(folder_p)}: {e_folder}", "ERROR")
        finally:
            if excel:
                excel.Quit()

    def _clear_folder_content(self, folder_path, folder_display_name, logger):
        """Utility to clear the contents of a folder."""
        logger(f"--- Pulizia della cartella '{folder_display_name}' in corso... ---", 'HEADER')
        if os.path.isdir(folder_path):
            for item_name in os.listdir(folder_path):
                item_path = os.path.join(folder_path, item_name)
                try:
                    if os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                    else:
                        os.remove(item_path)
                except Exception as e:
                    logger(f"Impossibile eliminare '{item_name}': {e}", 'ERROR')
        logger(f"--- Pulizia di '{folder_display_name}' completata. ---", 'SUCCESS')
