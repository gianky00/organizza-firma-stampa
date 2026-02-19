import os
import re
import traceback
from datetime import datetime
from typing import TypedDict

from src.utils.excel_gateway import ExcelGateway


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
        root_path = self.app_config.rinomina_path.get()
        if not os.path.isdir(root_path):
            self.logger(f"ERRORE: La cartella specificata non è valida o non esiste: '{root_path}'", "ERROR")
            self.gui.after(0, self.gui.on_process_finished)
            return

        try:
            self._rename_excel_files_in_place(root_path, cancel_event)
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

        DATE_IN_FILENAME_REGEX = re.compile(r"\s*\(\d{2}-\d{2}-\d{4}\)")
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
                    original_dir, original_filename = os.path.split(file_path)
                    base_name, ext = os.path.splitext(original_filename)
                    cleaned_base_name = DATE_IN_FILENAME_REGEX.sub("", base_name).strip()
                    cleaned_base_name = self._clean_windows_duplicate_marker(cleaned_base_name)
                    # Rimuove spazi e normalizza
                    cleaned_base_name = cleaned_base_name.replace(" ", "")
                    new_filename = f"{cleaned_base_name} ({emission_date.strftime('%d-%m-%Y')}){ext}"
                    
                    if new_filename.lower() != original_filename.lower():
                        new_filepath = os.path.join(original_dir, new_filename)
                        final_path = self._get_unique_filepath(new_filepath)
                        os.rename(file_path, final_path)
                        self.logger(f"  -> RINOMINATO in: {os.path.basename(final_path)}", "SUCCESS")
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

    def _get_unique_filepath(self, filepath: str) -> str:
        if not os.path.exists(filepath):
            return filepath
        base, ext = os.path.splitext(filepath)
        counter = 1
        while True:
            new_path = f"{base} ({counter}){ext}"
            if not os.path.exists(new_path):
                return new_path
            counter += 1

    def _clean_windows_duplicate_marker(self, name: str) -> str:
        return re.sub(r"\s*\(\d+\)$", "", name.strip())
