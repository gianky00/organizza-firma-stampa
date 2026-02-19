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
            os.path.join(dest_dir, d) for d in os.listdir(dest_dir) if os.path.isdir(os.path.join(dest_dir, d))
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

        excel_files = self._get_excel_files(source_dir)
        if not excel_files:
            return

        self.gui.after(0, self.setup_progress, len(excel_files), "Organizzazione in corso:")
        summary: OrgSummary = {"processed": 0, "errors": []}

        for i, fp in enumerate(excel_files):
            if cancel_event.is_set():
                break
            self.gui.after(0, self.update_progress, i + 1)
            self.logger(f"Processando: {os.path.basename(fp)}...")

            success, error = self._process_single_file(fp, dest_dir)
            if success:
                summary["processed"] += 1
            else:
                summary["errors"].append((os.path.basename(fp), error or "Errore sconosciuto"))

        self._log_org_summary(summary, len(excel_files))

    def _get_excel_files(self, source_dir):
        if not os.path.isdir(source_dir):
            self.logger("ERRORE: Cartella di origine non trovata.", "ERROR")
            return []
        try:
            files = [
                os.path.join(r, f)
                for r, _, fs in os.walk(source_dir)
                for f in fs
                if f.lower().endswith((".xls", ".xlsx", ".xlsm", ".xlsb")) and not f.startswith("~")
            ]
            if not files:
                self.logger("Nessun file Excel trovato.", "WARNING")
            return files
        except Exception as e:
            self.logger(f"ERRORE accesso cartella di origine: {e}", "ERROR")
            return []

    def _process_single_file(self, file_path, dest_dir):
        try:
            odc_s = self.excel_gateway.get_odc_value(file_path)
            dest_folder_name = (
                re.sub(r'[\\/:*?"<>|]', "", odc_s) if odc_s and odc_s.upper() != "NA" else "Schede senza ODC"
            )
            dest_folder_path = os.path.join(dest_dir, dest_folder_name)
            os.makedirs(dest_folder_path, exist_ok=True)
            shutil.copy2(file_path, dest_folder_path)
            return True, None
        except Exception as e:
            return False, str(e)

    def _log_org_summary(self, summary, total):
        self.logger(f"Processati {summary['processed']} file su {total}.")
        if summary["errors"]:
            self.logger(f"Si sono verificati {len(summary['errors'])} errori:", "ERROR")
            for f, err in summary["errors"]:
                self.logger(f" - {f}: {err}", "ERROR")

    def _print_files_in_folders(self, cancel_event, folder_list):
        self.gui.after(0, self.setup_progress, len(folder_list), "Stampa in corso:")
        from src.utils.excel_handler import ExcelHandler

        excel_h_class = getattr(self.excel_gateway, "excel_handler_class", ExcelHandler)

        with excel_h_class(self.logger) as excel:
            if not excel:
                return
            errors = []
            for i, folder_p in enumerate(folder_list):
                if cancel_event.is_set():
                    break
                self.gui.after(0, self.update_progress, i + 1)
                self.logger(f"Stampa cartella: {os.path.basename(folder_p)}")

                folder_errors = self._print_folder_content(excel, folder_p, cancel_event)
                errors.extend(folder_errors)

            if errors:
                self.logger(f"Errori durante la stampa di {len(errors)} file.", "ERROR")

    def _print_folder_content(self, excel, folder_path, cancel_event):
        errors = []
        try:
            excel_fs = [
                os.path.join(folder_path, f)
                for f in os.listdir(folder_path)
                if f.lower().endswith((".xls", ".xlsx", ".xlsm", ".xlsb")) and not f.startswith("~")
            ]
            if not excel_fs:
                self.logger("  -> Nessun file Excel trovato.", "WARNING")
                return []

            for fp in excel_fs:
                if cancel_event.is_set():
                    break
                success, err = self._print_single_excel_file(excel, fp)
                if not success:
                    errors.append((os.path.basename(fp), err))
        except Exception as e:
            self.logger(f"ERRORE cartella {os.path.basename(folder_path)}: {e}", "ERROR")
        return errors

    def _print_single_excel_file(self, excel, file_path):
        wb = None
        try:
            wb = excel.Workbooks.Open(file_path)
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
                self.logger(f"  -> Stampa inviata per: {os.path.basename(file_path)}", "SUCCESS")
                return True, None
            else:
                self.logger(f"  -> Ignorato (modello non trovato): {os.path.basename(file_path)}", "WARNING")
                return True, None
        except Exception as e:
            return False, str(e)
        finally:
            if wb:
                wb.Close(SaveChanges=False)

    def get_odc_to_canone_map(self, year, month):
        self.logger(f"Lettura del file Giornaliera per {month} {year}...", "INFO")
        giornaliera_path = self.fees_processor.get_giornaliera_path(year, month)

        if not os.path.isfile(giornaliera_path):
            self.logger(f"File Giornaliera non trovato: {giornaliera_path}", "WARNING")
            return {}

        from src.utils.excel_handler import ExcelHandler

        excel_h_class = getattr(self.excel_gateway, "excel_handler_class", ExcelHandler)

        mapping = {}
        with excel_h_class(self.logger) as excel:
            if not excel:
                return {}
            wb = None
            try:
                wb = excel.Workbooks.Open(giornaliera_path, ReadOnly=True)
                ws = wb.Worksheets("RIEPILOGO")
                mapping = self._extract_mapping_from_riepilogo(ws)
            except Exception as e:
                self.logger(f"Errore lettura Giornaliera: {e}", "ERROR")
            finally:
                if wb:
                    wb.Close(SaveChanges=False)
        self.logger(f"Mappa ODC creata con {len(mapping)} voci.", "INFO")
        return mapping

    def _extract_mapping_from_riepilogo(self, worksheet):
        mapping = {}
        cells_to_check = [("S16", "S17"), ("U16", "U17"), ("V16", "V17")]
        for header_cell, value_cell in cells_to_check:
            header = worksheet.Range(header_cell).Value
            value_raw = worksheet.Range(value_cell).Value
            if header and value_raw:
                odc_num = str(value_raw).split("\n")[0].strip()
                if odc_num.isdigit():
                    mapping[odc_num] = str(header).lower()
        return mapping
