import os
import re
import traceback
from datetime import datetime
from pathlib import Path
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
        import pythoncom

        pythoncom.CoInitialize()  # Inizializzazione COM per il thread corrente

        self.logger("Avvio del processo di ridenominazione...", "HEADER")
        try:
            # 0. Mostra barra di caricamento immediata
            if hasattr(self.gui, "show_indeterminate"):
                self.gui.show_indeterminate("Inizializzazione ambiente...")

            # 1. Identificazione Sorgente
            source_path = self.app_config.rinomina_path.get()

            if not source_path or not Path(source_path).is_dir():
                self.logger(f"ERRORE: La cartella specificata non è valida o non esiste: '{source_path}'", "ERROR")
                return

            # 2. Elaborazione DIRETTA sulla sorgente
            self.logger(f"Elaborazione in corso direttamente su: {source_path}", "INFO")
            self._rename_excel_files_in_place(source_path, cancel_event)

        except Exception as e:
            self.logger(f"ERRORE CRITICO E IMPREVISTO durante la ridenominazione: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            if cancel_event.is_set():
                self.logger("Processo di ridenominazione annullato.", "WARNING")
            self.hide_progress()
            import pythoncom

            pythoncom.CoUninitialize()  # Rilascio risorse COM per questo thread

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
        self.setup_progress(num_files, "Analisi e ridenominazione:")

        date_in_filename_regex = re.compile(r"\s*\(\d{2}-\d{2}-\d{4}\)")
        summary: RenameSummary = {"corrected": 0, "already_ok": 0, "no_date": 0, "errors": []}

        password = self.app_config.rinomina_password.get()

        from src.utils.excel_handler import ExcelHandler

        excel_h_class = getattr(self.excel_gateway, "excel_handler_class", ExcelHandler)

        # TURBO: Apriamo Excel una volta sola per leggere tutte le date
        with excel_h_class(self.logger) as excel:
            if not excel:
                return

            for i, file_path in enumerate(excel_files):
                if cancel_event.is_set():
                    return
                self.update_progress(i + 1)

                try:
                    # Estrazione data ultra-veloce tramite istanza condivisa
                    emission_date = self._get_date_with_instance(excel, file_path, password)

                    if emission_date:
                        original_path = Path(file_path)
                        original_dir = original_path.parent
                        original_filename = original_path.name
                        base_name, ext = original_path.stem, original_path.suffix

                        cleaned_base_name = date_in_filename_regex.sub("", base_name).strip()
                        cleaned_base_name = self._clean_windows_duplicate_marker(cleaned_base_name)

                        # Separazione universale di 'NEW' e 'OLD' se attaccati al testo precedente
                        # Esempio: 10P152SHNEW -> 10P152SH - NEW, 10P152SHOLD -> 10P152SH - OLD
                        cleaned_base_name = re.sub(r"([a-zA-Z0-9])NEW\b", r"\1 - NEW", cleaned_base_name)
                        cleaned_base_name = re.sub(r"([a-zA-Z0-9])OLD\b", r"\1 - OLD", cleaned_base_name)

                        # Normalizzazione: rimuove eventuali doppi spazi e pulisce i bordi
                        cleaned_base_name = re.sub(r"\s+", " ", cleaned_base_name).strip()

                        new_filename = f"{cleaned_base_name} ({emission_date.strftime('%d-%m-%Y')}){ext}"

                        if new_filename.lower() != original_filename.lower():
                            new_filepath = original_dir / new_filename
                            final_path = self._get_unique_filepath(new_filepath)
                            os.rename(file_path, str(final_path))
                            self.logger(f"Ridenominato: {original_filename} -> {final_path.name}", "SUCCESS")
                            summary["corrected"] += 1
                        else:
                            summary["already_ok"] += 1
                    else:
                        self.logger(
                            f"Data non trovata nel file (cella non riconosciuta): {os.path.basename(file_path)}",
                            "WARNING",
                        )
                        summary["no_date"] += 1
                except Exception as e:
                    self.logger(f"Errore durante l'analisi di {os.path.basename(file_path)}: {e}", "ERROR")
                    summary["errors"].append((os.path.basename(file_path), str(e)))

        self._log_rename_summary(summary, num_files)

    def _get_date_with_instance(self, excel, file_path, password) -> datetime | None:
        try:
            # Apertura silenziosa e rapida
            wb = excel.Workbooks.Open(file_path, 0, True, None, password)
            try:
                ws = wb.Worksheets(1)
                # Usiamo la logica robusta definita nel gateway
                return self.excel_gateway.extract_date_from_worksheet(ws)
            finally:
                wb.Close(False)
        except Exception as e:
            self.logger(f"Impossibile aprire il file {os.path.basename(file_path)}: {e}", "DEBUG")
            return None

    def _log_rename_summary(self, summary, total):
        self.logger("\n--- RIEPILOGO PROCESSO RINOMINA ---", "HEADER")
        self.logger(f"File rinominati: {summary['corrected']} / {total}", "SUCCESS")
        if summary["errors"]:
            self.logger(f"Errori rilevati: {len(summary['errors'])}", "ERROR")

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
