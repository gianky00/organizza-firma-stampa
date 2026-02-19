import os
import re
import shutil
from typing import TypedDict

from src.utils.excel_gateway import ExcelGateway
from src.utils.file_utils import create_backup


class OrgSummary(TypedDict):
    processed: int
    errors: list[tuple[str, str]]


class OrganizationProcessor:
    def __init__(
        self,
        gui,
        app_config,
        fees_processor,
        setup_progress_cb,
        update_progress_cb,
        hide_progress_cb,
        excel_gateway_class=None,
    ):
        self.gui = gui
        self.app_config = app_config
        self.fees_processor = fees_processor
        self.logger = gui.log_organizza
        self.setup_progress = setup_progress_cb
        self.update_progress = update_progress_cb
        self.hide_progress = hide_progress_cb
        self.excel_gateway = (excel_gateway_class or ExcelGateway)(self.logger)
        self.stampa_processing_data = {
            "schedacontrolloSTRUMENTIANALOGICI": {"PrintArea": "A2:N55"},
            "schedacontrolloSTRUMENTIDIGITALI": {"PrintArea": "A2:N50"},
            "SchedacontrolloREPORTMANUTENZIONECORRETTIVA": {"PrintArea": "A2:N55"},
            "schedacontrolloBILANCE": {"PrintArea": "A2:N50"},
            "schedacontrolloCALIBRI": {"PrintArea": "A2:N45"},
            "schedacontrolloMICROMETRI": {"PrintArea": "A2:N45"},
            "schedacontrolloCOMPARATORI": {"PrintArea": "A2:N45"},
            "schedacontrolloALESAMETRI": {"PrintArea": "A2:N45"},
            "schedacontrolloDISCOCALIBRO": {"PrintArea": "A2:N45"},
            "schedacontrolloSQUADRE": {"PrintArea": "A2:N45"},
            "schedacontrolloGONIOMETRI": {"PrintArea": "A2:N45"},
            "schedacontrolloPRISMI": {"PrintArea": "A2:N45"},
            "schedacontrolloRIGHE": {"PrintArea": "A2:N45"},
            "schedacontrolloLIVELLADIGITALE": {"PrintArea": "A2:N45"},
        }

    def run_organization_process(self, cancel_event):
        dest_dir = self.app_config.organizza_dest_dir.get()
        if os.path.isdir(dest_dir) and any(os.scandir(dest_dir)):
            self.logger("Creazione backup cartella di destinazione...")
            if not create_backup(dest_dir):
                self.logger("ERRORE: Impossibile creare il backup. Operazione annullata.", "ERROR")
                return

        self._organize_files(cancel_event)
        if cancel_event.is_set():
            self.logger("Operazione annullata dall'utente.", "WARNING")
        else:
            self.logger("Organizzazione completata!", "SUCCESS")

    def run_printing_process(self, cancel_event):
        dest_dir = self.app_config.organizza_dest_dir.get()
        if not os.path.isdir(dest_dir):
            self.logger("ERRORE: Cartella organizzata non trovata.", "ERROR")
            return
        
        folder_list = [
            os.path.join(dest_dir, d)
            for d in os.listdir(dest_dir)
            if os.path.isdir(os.path.join(dest_dir, d))
        ]
        
        if not folder_list:
            self.logger("Nessuna cartella trovata nella destinazione.", "WARNING")
            return

        self._print_files_in_folders(cancel_event, folder_list)
        if cancel_event.is_set():
            self.logger("Stampa annullata.", "WARNING")
        else:
            self.logger("Processo di stampa completato!", "SUCCESS")
        
        self.gui.after(0, self.hide_progress)
        self.gui.after(0, self.gui.on_process_finished)

    def _organize_files(self, cancel_event):
        source_dir = self.app_config.organizza_source_dir.get()
        dest_dir = self.app_config.organizza_dest_dir.get()
        if not os.path.isdir(source_dir):
            self.logger("ERRORE: Cartella di origine non trovata.", "ERROR")
            return
        try:
            excel_files = [
                os.path.join(r, f)
                for r, _, fs in os.walk(source_dir)
                for f in fs
                if f.lower().endswith((".xls", ".xlsx", ".xlsm", ".xlsb")) and not f.startswith("~")
            ]
        except Exception as e:
            self.logger(f"ERRORE accesso cartella di origine: {e}", "ERROR")
            return
        if not excel_files:
            self.logger("Nessun file Excel trovato.", "WARNING")
            return

        self.gui.after(0, self.setup_progress, len(excel_files), "Organizzazione in corso:")
        summary: OrgSummary = {"processed": 0, "errors": []}

        for i, fp in enumerate(excel_files):
            if cancel_event.is_set():
                return
            self.gui.after(0, self.update_progress, i + 1)
            self.logger(f"Processando: {os.path.basename(fp)}...")
            
            try:
                odc_s = self.excel_gateway.get_odc_value(fp)
                dest_folder_name = (
                    re.sub(r'[\\/:*?"<>|]', "", odc_s) if odc_s and odc_s.upper() != "NA" else "Schede senza ODC"
                )
                dest_folder_path = os.path.join(dest_dir, dest_folder_name)
                os.makedirs(dest_folder_path, exist_ok=True)
                shutil.copy2(fp, dest_folder_path)
                summary["processed"] += 1
            except Exception as e:
                summary["errors"].append((os.path.basename(fp), f"Dettagli: {e}"))

        self.logger(f"Processati {summary['processed']} file su {len(excel_files)}.")
        if summary["errors"]:
            self.logger(f"Si sono verificati {len(summary['errors'])} errori:", "ERROR")
            for f, err in summary["errors"]:
                self.logger(f" - {f}: {err}", "ERROR")

    def _print_files_in_folders(self, cancel_event, folder_list):
        self.gui.after(0, self.setup_progress, len(folder_list), "Stampa in corso:")
        # Per la stampa usiamo ancora ExcelHandler direttamente o espandiamo il Gateway.
        # Per ora manteniamo la logica di stampa qui per brevità, ma iniettiamo l'handler.
        from src.utils.excel_handler import ExcelHandler
        excel_h_class = getattr(self.excel_gateway, 'excel_handler_class', ExcelHandler)
        
        with excel_h_class(self.logger) as excel:
            if not excel:
                return
            errors = []
            for i, folder_p in enumerate(folder_list):
                if cancel_event.is_set():
                    return
                self.gui.after(0, self.update_progress, i + 1)
                self.logger(f"Stampa cartella: {os.path.basename(folder_p)}")
                try:
                    excel_fs = [
                        os.path.join(folder_p, f)
                        for f in os.listdir(folder_p)
                        if f.lower().endswith((".xls", ".xlsx", ".xlsm", ".xlsb")) and not f.startswith("~")
                    ]
                    if not excel_fs:
                        self.logger("  -> Nessun file Excel trovato.", "WARNING")
                        continue
                    for fp in excel_fs:
                        if cancel_event.is_set():
                            return
                        wb = None
                        try:
                            wb = excel.Workbooks.Open(fp)
                            ws = wb.Worksheets(1)
                            m_val = next(
                                (
                                    str(ws.Cells(r, c).Value).strip()
                                    for r, c in [(2, 5), (2, 20), (5, 20)]
                                    if ws.Cells(r, c).Value and str(ws.Cells(r, c).Value).strip()
                                ),
                                "",
                            )
                            cleaned_model = re.sub(r"\W", "", m_val)
                            if cleaned_model in self.stampa_processing_data:
                                ws.PageSetup.PrintArea = self.stampa_processing_data[cleaned_model]["PrintArea"]
                                wb.PrintOut()
                                self.logger(f"  -> Stampa inviata per: {os.path.basename(fp)}", "SUCCESS")
                            else:
                                self.logger(f"  -> Ignorato (modello non trovato): {os.path.basename(fp)}", "WARNING")
                        except Exception as e_file:
                            errors.append((os.path.basename(fp), f"Dettagli: {e_file}"))
                        finally:
                            if wb:
                                wb.Close(SaveChanges=False)
                except Exception as e_folder:
                    self.logger(f"ERRORE cartella {os.path.basename(folder_p)}: {e_folder}", "ERROR")

            if errors:
                self.logger(f"Errori durante la stampa di {len(errors)} file.", "ERROR")

    def get_odc_to_canone_map(self, year, month):
        self.logger(f"Lettura del file Giornaliera per {month} {year}...", "INFO")
        giornaliera_path = self.fees_processor.get_giornaliera_path(year, month)

        if not os.path.isfile(giornaliera_path):
            self.logger(f"File Giornaliera non trovato: {giornaliera_path}", "WARNING")
            return {}

        mapping = {}
        # Usiamo l'handler iniettato tramite il gateway
        from src.utils.excel_handler import ExcelHandler
        excel_h_class = getattr(self.excel_gateway, "excel_handler_class", ExcelHandler)

        with excel_h_class(self.logger) as excel:
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
                        odc_num = str(value_raw).split("\n")[0].strip()
                        if odc_num.isdigit():
                            mapping[odc_num] = str(header).lower()
            except Exception as e:
                self.logger(f"Errore lettura Giornaliera: {e}", "ERROR")
            finally:
                if wb:
                    wb.Close(SaveChanges=False)
        self.logger(f"Mappa ODC creata con {len(mapping)} voci.", "INFO")
        return mapping
