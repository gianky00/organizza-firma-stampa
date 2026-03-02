import os
import re
import traceback
from pathlib import Path
from typing import TypedDict

from src.utils import constants as const
from src.utils.excel_gateway import ExcelGateway
from src.utils.file_utils import create_backup


class RenameSummary(TypedDict):
    corrected: int
    already_ok: int
    no_date: int
    errors: list[tuple[str, str]]


class RenameProcessor:
    def __init__(
        self,
        gui,
        app_config,
        setup_progress_cb,
        update_progress_cb,
        hide_progress_cb,
        excel_gateway_class=None,
    ):
        self.gui = gui
        self.app_config = app_config
        self.logger = gui.log_rinomina
        self.setup_progress = setup_progress_cb
        self.update_progress = update_progress_cb
        self.hide_progress = hide_progress_cb
        self.excel_gateway = (excel_gateway_class or ExcelGateway)(self.logger)

    def run_rename_process(self, cancel_event):
        self.logger("Avvio del processo di ridenominazione...", "HEADER")
        try:
            # 0. Mostra barra di caricamento immediata
            if hasattr(self.gui, "show_indeterminate"):
                self.gui.after(0, self.gui.show_indeterminate, "Inizializzazione ambiente...")

            # 1. Identificazione Area di Lavoro Locale e Sorgente
            local_work_path = os.path.join(const.APPLICATION_PATH, const.RINOMINA_DEFAULT_DIR)
            source_path = self.app_config.rinomina_path.get()

            # 2. Logica di Importazione
            if os.path.normpath(source_path) != os.path.normpath(local_work_path):
                if hasattr(self.gui, "show_indeterminate"):
                    self.gui.after(0, self.gui.show_indeterminate, "Importazione file da rete...")
                self.logger(f"Importazione file da sorgente: {source_path}", "INFO")
                # Pulizia locale preventiva
                from src.utils.file_utils import clear_folder_content

                clear_folder_content(local_work_path, self.logger, folder_display_name="Area di Lavoro Locale")

                excel_files_to_import = self._get_excel_files(source_path, cancel_event)
                if not excel_files_to_import:
                    return

                import shutil

                for f in excel_files_to_import:
                    shutil.copy2(f, local_work_path)

                self.logger(f"Importati {len(excel_files_to_import)} file Excel.", "SUCCESS")
                active_path = local_work_path
            else:
                active_path = source_path

            if not Path(active_path).is_dir():
                self.logger(f"ERRORE: La cartella specificata non è valida o non esiste: '{active_path}'", "ERROR")
                return

            # 3. Elaborazione
            self._rename_excel_files_in_place(active_path, cancel_event)

            # 4. Auto-Clean Finale (solo se sorgente era esterna)
            if not cancel_event.is_set() and os.path.normpath(source_path) != os.path.normpath(local_work_path):
                self.logger("--- PULIZIA: Spostamento file rinominati in backup... ---", "INFO")
                backup_parent = os.path.join(const.APPLICATION_PATH, const.BACKUP_DIR, "Originali_Rinominati")
                if create_backup(active_path, backup_parent_dir=backup_parent):
                    from src.utils.file_utils import clear_folder_content

                    clear_folder_content(active_path, self.logger, folder_display_name="Area Rinominati")

        except Exception as e:
            self.logger(f"ERRORE CRITICO E IMPREVISTO durante la ridenominazione: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            if cancel_event.is_set():
                self.logger("Processo di ridenominazione annullato.", "WARNING")
            self.gui.after(0, self.hide_progress)
            self.gui.after(0, self.gui.on_process_finished)

    def _rename_excel_files_in_place(self, root_path, cancel_event):
        self.logger("[FASE 1/2] Raccolta file Excel...", "HEADER")
        excel_files = self._get_excel_files(root_path, cancel_event)

        if not excel_files:
            self.logger("Nessun file Excel trovato.", "WARNING")
            return
        if cancel_event.is_set():
            return

        num_files = len(excel_files)
        self.logger(f"Trovati {num_files} file Excel. Inizio analisi.", "INFO")
        self.gui.after(0, self.setup_progress, num_files, "Analisi e ridenominazione:")

        date_in_filename_regex = re.compile(r"\s*\(\d{2}-\d{2}-\d{4}\)")
        summary: RenameSummary = {"corrected": 0, "already_ok": 0, "no_date": 0, "errors": []}

        password = self.app_config.rinomina_password.get()

        for i, file_path in enumerate(excel_files):
            if cancel_event.is_set():
                return
            self.gui.after(0, self.update_progress, i + 1)
            self.logger(f"Analisi: {os.path.basename(file_path)}...")

            try:
                emission_date = self.excel_gateway.get_workbook_date(file_path, password=password)

                if emission_date:
                    original_path = Path(file_path)
                    original_dir = original_path.parent
                    original_filename = original_path.name
                    base_name, ext = original_path.stem, original_path.suffix

                    cleaned_base_name = date_in_filename_regex.sub("", base_name).strip()
                    cleaned_base_name = self._clean_windows_duplicate_marker(cleaned_base_name)
                    # Rimuove spazi e normalizza
                    cleaned_base_name = cleaned_base_name.replace(" ", "")
                    new_filename = f"{cleaned_base_name} ({emission_date.strftime('%d-%m-%Y')}){ext}"

                    if new_filename.lower() != original_filename.lower():
                        new_filepath = original_dir / new_filename
                        final_path = self._get_unique_filepath(new_filepath)
                        os.rename(file_path, str(final_path))
                        self.logger(f"  -> RINOMINATO in: {final_path.name}", "SUCCESS")
                        summary["corrected"] += 1
                    else:
                        self.logger("  -> Già corretto.", "INFO")
                        summary["already_ok"] += 1
                else:
                    self.logger("  -> Data non trovata.", "WARNING")
                    summary["no_date"] += 1
            except Exception as e:
                error_msg = f"Dettagli: {e}"
                self.logger(f"--- ERRORE FILE: {os.path.basename(file_path)} ---", "ERROR")
                self.logger(error_msg, "ERROR")
                summary["errors"].append((os.path.basename(file_path), error_msg))

        self.logger("\n--- RIEPILOGO PROCESSO RINOMINA ---", "HEADER")
        self.logger(f"File rinominati o corretti: {summary['corrected']}", "SUCCESS")
        self.logger(f"File già corretti: {summary['already_ok']}", "INFO")
        self.logger(f"File con data non trovata: {summary['no_date']}", "WARNING")
        self.logger(f"File con errori: {len(summary['errors'])}", "ERROR")
        if summary["errors"]:
            self.logger("\n--- DETTAGLIO ERRORI ---", "HEADER")
            for file_name, err in summary["errors"]:
                self.logger(f"- {file_name}: {err}", "ERROR")
        self.logger("--- COMPLETATO ---", "HEADER")

    def _get_excel_files(self, root_path, cancel_event):
        files = []
        for r, _, fs in os.walk(root_path):
            if cancel_event.is_set():
                break
            for f in fs:
                if f.lower().endswith((".xls", ".xlsx", ".xlsm", ".xlsb")) and not f.startswith("~"):
                    files.append(os.path.join(r, f))
        return files

    def _get_unique_filepath(self, filepath: Path) -> Path:
        if not filepath.exists():
            return filepath
        base, ext = filepath.stem, filepath.suffix
        original_dir = filepath.parent
        counter = 1
        while True:
            new_path = original_dir / f"{base} ({counter}){ext}"
            if not new_path.exists():
                return new_path
            counter += 1

    def _clean_windows_duplicate_marker(self, name: str) -> str:
        return re.sub(r"\s*\(\d+\)$", "", name.strip())
