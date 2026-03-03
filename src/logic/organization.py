import os
import re
import shutil
from pathlib import Path
from typing import TypedDict

from src.utils import constants as const
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
        self._load_stampa_mappings()

    def _load_stampa_mappings(self):
        """Carica le mappature di stampa dalla configurazione unificata come lista."""
        self.stampa_models_list = []
        config_data = self.app_config.config_manager.get("rename_models_config")
        if config_data:
            for item in config_data:
                # Normalizziamo il match_value per il confronto
                mv = re.sub(r"[\W_]+", "", str(item.get("match_value", "")).lower())
                if mv:
                    self.stampa_models_list.append(
                        {
                            "match_value": mv,
                            "PrintArea": item.get("print_area", "A1:N50"),
                            "id_cells": item.get("id_cells", ["F2", "E2", "T2"]),
                        }
                    )

    def run_organization_process(self, cancel_event):
        try:
            source_dir = self.app_config.organizza_source_dir.get()
            dest_dir = self.app_config.organizza_dest_dir.get()
            local_source_path = os.path.join(const.APPLICATION_PATH, const.ORGANIZZA_SOURCE_DIR)

            # 1. Logica di Importazione per Organizzazione
            if os.path.normpath(source_dir) != os.path.normpath(local_source_path):
                self.logger(f"Importazione schede da sorgente: {source_dir}", "INFO")
                from src.utils.file_utils import clear_folder_content

                clear_folder_content(local_source_path, self.logger, folder_display_name="Area Sorgente Locale")

                files_to_import = self._get_excel_files(source_dir)
                if not files_to_import:
                    return

                import shutil

                for f in files_to_import:
                    shutil.copy2(f, local_source_path)

                self.logger(f"Importate {len(files_to_import)} schede nell'area locale.", "SUCCESS")
                active_source = local_source_path
            else:
                active_source = source_dir

            # 2. Backup Destinazione
            if Path(dest_dir).is_dir() and any(os.scandir(dest_dir)):
                self.logger("Creazione backup cartella di destinazione...")
                backup_parent = os.path.join(const.APPLICATION_PATH, const.BACKUP_DIR, "Storico_Organizzate")
                if not create_backup(dest_dir, backup_parent_dir=backup_parent):
                    self.logger("ERRORE: Impossibile creare il backup. Operazione annullata.", "ERROR")
                    return

            # 3. Elaborazione
            self._organize_files_at_path(active_source, cancel_event)

            if not cancel_event.is_set():
                self.logger("--- PULIZIA: Spostamento originali in backup... ---", "INFO")
                backup_parent_source = os.path.join(const.APPLICATION_PATH, const.BACKUP_DIR, "Originali_Organizzati")
                if create_backup(active_source, backup_parent_dir=backup_parent_source):
                    from src.utils.file_utils import clear_folder_content

                    clear_folder_content(active_source, self.logger, folder_display_name="Schede Lavorate")

                self.logger("Organizzazione completata!", "SUCCESS")
            else:
                self.logger("Operazione annullata dall'utente.", "WARNING")
        finally:
            self.hide_progress()

    def run_printing_process(self, cancel_event, folder_list=None):
        self.logger("Avvio del processo di stampa schede...", "HEADER")
        self._load_stampa_mappings()  # Ricarica le mappature dai settings
        try:
            dest_dir = self.app_config.organizza_dest_dir.get()
            if not Path(dest_dir).is_dir():
                self.logger("ERRORE: Cartella organizzata non trovata.", "ERROR")
                return

            # Se folder_list non è fornito, usa tutte le cartelle nella destinazione (fallback)
            if folder_list is None:
                folder_list = [
                    os.path.join(dest_dir, d) for d in os.listdir(dest_dir) if Path(os.path.join(dest_dir, d)).is_dir()
                ]

            if not folder_list:
                self.logger("Nessuna cartella selezionata o trovata per la stampa.", "WARNING")
                return

            self._print_files_in_folders(cancel_event, folder_list)
            if cancel_event.is_set():
                self.logger("Stampa annullata.", "WARNING")
            else:
                self.logger("Processo di stampa completato!", "SUCCESS")
        except Exception as e:
            self.logger(f"ERRORE FATALE durante il processo di stampa: {e}", "ERROR")
        finally:
            self.hide_progress()

    def _organize_files_at_path(self, source_dir, cancel_event):
        dest_dir = self.app_config.organizza_dest_dir.get()

        excel_files = self._get_excel_files(source_dir)
        if not excel_files:
            self.logger("Nessun file Excel trovato da organizzare.", "WARNING")
            return

        self.setup_progress(len(excel_files), "Organizzazione in corso:")
        summary: OrgSummary = {"processed": 0, "errors": []}

        for i, fp in enumerate(excel_files):
            if cancel_event.is_set():
                break
            self.update_progress(i + 1)
            self.logger(f"Processando: {os.path.basename(fp)}...")

            success, error = self._process_single_file(fp, dest_dir)
            if success:
                summary["processed"] += 1
            else:
                summary["errors"].append((os.path.basename(fp), error or "Errore sconosciuto"))

        self._log_org_summary(summary, len(excel_files))

    def _get_excel_files(self, source_dir):
        if not Path(source_dir).is_dir():
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
            # Prevenzione Path Traversal: puliamo ulteriormente il nome
            dest_folder_name = os.path.basename(dest_folder_name)
            dest_folder_path = Path(dest_dir) / dest_folder_name

            # Validazione che il path sia effettivamente interno alla directory di destinazione
            if str(Path(dest_dir).resolve()) not in str(dest_folder_path.resolve()):
                return False, "Tentativo di path traversal bloccato."

            dest_folder_path.mkdir(parents=True, exist_ok=True)
            shutil.copy2(file_path, str(dest_folder_path))
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
        self.setup_progress(len(folder_list), "Stampa in corso:")
        from src.utils.excel_handler import ExcelHandler

        excel_h_class = getattr(self.excel_gateway, "excel_handler_class", ExcelHandler)

        with excel_h_class(self.logger) as excel:
            if not excel:
                self.logger("Impossibile avviare il gestore Excel.", "ERROR")
                return
            errors = []
            for i, folder_p in enumerate(folder_list):
                if cancel_event.is_set():
                    break
                self.update_progress(i + 1)
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
            if wb is None:
                return False, "Impossibile aprire il file Excel (Workbook è None)."
            ws = wb.Worksheets(1)

            # 1. Identificazione del modello tramite Gateway (Dinamico dai settings)
            cfg = self.excel_gateway.identify_model(ws)

            # 2. Esecuzione stampa
            if cfg:
                ws.PageSetup.PrintArea = cfg.get("print_area", "A1:N50")
                wb.PrintOut()
                self.logger(f"  -> Stampa inviata ({cfg.get('match_value')}): {os.path.basename(file_path)}", "SUCCESS")
                return True, None
            else:
                # Fallback dinamico: scansiona TUTTE le celle ID che l'utente ha configurato nella tabella
                # per mostrare cosa contengono e facilitare la correzione del Match ID.
                hints = []
                config_data = self.app_config.config_manager.get("rename_models_config") or []

                # Raccogliamo tutte le celle ID uniche presenti nelle configurazioni
                target_cells = set()
                for m in config_data:
                    cells = m.get("id_cells", [])
                    if not cells and m.get("id_cell"):
                        cells = [m.get("id_cell")]
                    for c in cells:
                        if c:
                            target_cells.add(c.strip().upper())

                # Scansione delle celle effettivamente in uso
                for cell_ref in sorted(list(target_cells)):
                    try:
                        val = ws.Range(cell_ref).Value
                        if val:
                            clean_v = re.sub(r"[\W_]+", "", str(val).strip().lower())
                            hints.append(f"{cell_ref}:'{clean_v}'")
                    except Exception:
                        continue

                hint_str = " | ".join(hints) if hints else "nessun valore trovato nelle celle ID configurate"
                self.logger(
                    f"  -> Modello NON riconosciuto per {os.path.basename(file_path)}. Contenuto celle ID: {hint_str}",
                    "WARNING",
                )
                return True, None
        except Exception as e:
            return False, str(e)
        finally:
            if wb:
                wb.Close(SaveChanges=False)

    def get_odc_to_canone_map(self, year, month):
        self.logger(f"Lettura del file Giornaliera per {month} {year}...", "INFO")
        giornaliera_path = self.fees_processor.get_giornaliera_path(year, month)

        if not Path(giornaliera_path).is_file():
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
                if wb is None:
                    self.logger("Apertura file Giornaliera fallita: Workbook è None.", "ERROR")
                    return {}
                try:
                    ws = wb.Worksheets("RIEPILOGO")
                except Exception:
                    self.logger("Foglio 'RIEPILOGO' non trovato nel file Giornaliera.", "WARNING")
                    return {}
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
        cells_to_check = (("S16", "S17"), ("U16", "U17"), ("V16", "V17"))
        for header_cell, value_cell in cells_to_check:
            header = worksheet.Range(header_cell).Value
            value_raw = worksheet.Range(value_cell).Value
            if header and value_raw:
                odc_num = str(value_raw).split("\n")[0].strip()
                if odc_num.isdigit():
                    mapping[odc_num] = str(header).lower()
        return mapping
