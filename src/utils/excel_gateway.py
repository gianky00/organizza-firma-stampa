from __future__ import annotations

import os
import re
from contextlib import suppress
from datetime import datetime, timedelta
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from datetime import datetime

from src.utils.excel_handler import ExcelHandler


class ExcelGateway:
    """
    Astrae le operazioni complesse su Excel (COM) fornendo un'interfaccia pulita.
    Segue il Gateway Pattern per isolare la logica di business dalle dipendenze esterne instabili.
    """

    def __init__(self, logger, excel_handler_class=None):
        self.logger = logger
        self.excel_handler_class = excel_handler_class or ExcelHandler

    def get_odc_value(self, file_path: str) -> str | None:
        """Estrae il codice ODC da un file Excel provando diverse celle note."""
        with self.excel_handler_class(self.logger) as excel:
            if not excel:
                return None
            wb = None
            try:
                wb = excel.Workbooks.Open(file_path)
                ws = wb.Worksheets(1)
                # Celle candidate per l'ODC in base ai modelli noti
                candidates = ["L50", "L45", "DB14", "DB17"]
                for cell in candidates:
                    try:
                        val = ws.Range(cell).Value
                        if val is not None and str(val).strip() != "":
                            if isinstance(val, (int, float)):
                                return str(int(val))
                            return str(val).strip()
                    except Exception:
                        continue
                return None
            except Exception as e:
                self.logger(f"Errore lettura ODC da {os.path.basename(file_path)}: {e}", "ERROR")
                return None
            finally:
                if wb:
                    wb.Close(SaveChanges=False)

    def apply_signature_and_export_pdf(self, excel_path: str, pdf_path: str, image_path: str, config: dict) -> bool:
        """Applica l'immagine della firma in una posizione specifica e esporta in PDF."""
        with self.excel_handler_class(self.logger) as excel:
            if not excel:
                return False
            wb = None
            try:
                wb = excel.Workbooks.Open(excel_path)
                ws = wb.Worksheets(1)

                ws.Shapes.AddPicture(
                    image_path,
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Left=config.get("left", 0),
                    Top=config.get("top", 0),
                    Width=config.get("width", -1),
                    Height=config.get("height", -1),
                )

                wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
                return True
            except Exception as e:
                self.logger(f"Errore firma/export PDF per {os.path.basename(excel_path)}: {e}", "ERROR")
                return False
            finally:
                if wb:
                    wb.Close(SaveChanges=False)

    def run_macro_and_print(self, excel_path: str, macro_name: str, printer_name: str) -> bool:
        """Esegue una macro VBA e invia il documento alla stampante specifica."""
        with self.excel_handler_class(self.logger) as excel:
            if not excel:
                return False
            wb = None
            try:
                wb = excel.Workbooks.Open(excel_path)
                excel.ActivePrinter = printer_name
                excel.Run(macro_name)
                return True
            except Exception as e:
                self.logger(f"Errore macro/stampa per {os.path.basename(excel_path)}: {e}", "ERROR")
                return False
            finally:
                if wb:
                    wb.Close(SaveChanges=False)

    def get_workbook_date(self, file_path: str, password: str = "") -> datetime | None:
        """Apre un workbook (anche protetto) e tenta di estrarne la data di emissione."""
        with self.excel_handler_class(self.logger) as excel:
            if not excel:
                return None
            wb = None
            try:
                wb = self._open_workbook(excel, file_path, password)
                if not wb:
                    return None

                ws = wb.Worksheets(1)
                return self.extract_date_from_worksheet(ws)
            except Exception as e:
                self.logger(f"Errore estrazione data da {os.path.basename(file_path)}: {e}", "ERROR")
                return None
            finally:
                if wb:
                    wb.Close(SaveChanges=False)

    def extract_date_from_worksheet(self, ws: Any) -> datetime | None:
        """Logica core per estrarre la data da un foglio di lavoro usando modelli e candidati."""
        from src.domain.models import DEFAULT_DATE_CANDIDATES

        # Carica configurazione dinamica dei modelli
        models_config = self._get_dynamic_models_config()

        # Prova matching modelli specifici
        dt = self._find_date_by_model(ws, models_config)
        if dt:
            return dt

        # Prova candidati di default
        dt = self._find_date_in_cells(ws, DEFAULT_DATE_CANDIDATES)
        if dt:
            return dt

        return None

    def extract_tcl_from_worksheet(self, ws: Any) -> str | None:
        """Estrae il nome del TCL (referente) basandosi sul modello riconosciuto."""
        models_config = self._get_dynamic_models_config()

        for cfg in models_config:
            if cfg.get("id_cell"):
                try:
                    val = self._normalize_model_string(ws.Range(cfg["id_cell"]).Value)
                    if val == cfg.get("match_value"):
                        tcl_cell = cfg.get("tcl_cell")
                        if tcl_cell:
                            tcl_val = ws.Range(tcl_cell).Value
                            if tcl_val:
                                return str(tcl_val).strip()
                except Exception:
                    continue
        return None

    def _get_dynamic_models_config(self) -> list[dict]:
        """Recupera i modelli dalla configurazione dell'app."""
        from src.utils.config_manager import ConfigManager

        config = ConfigManager()
        models = config.get("rename_models_config")
        return list(models) if isinstance(models, list) else []

    def _open_workbook(self, excel, file_path, password):
        try:
            return excel.Workbooks.Open(file_path, ReadOnly=True)
        except Exception:
            if password:
                return excel.Workbooks.Open(file_path, ReadOnly=True, Password=password)
            raise

    def _find_date_by_model(self, worksheet, models_config) -> datetime | None:
        for cfg in models_config:
            # Gestisce sia oggetti (RENAME_MODELS) che dizionari (da JSON config)
            id_cell = getattr(cfg, "id_cell", cfg.get("id_cell") if isinstance(cfg, dict) else None)
            match_value = getattr(cfg, "match_value", cfg.get("match_value") if isinstance(cfg, dict) else None)
            date_cells = getattr(cfg, "date_cells", cfg.get("date_cells") if isinstance(cfg, dict) else [])

            if id_cell:
                val = self._normalize_model_string(worksheet.Range(id_cell).Value)
                if val == match_value:
                    return self._find_date_in_cells(worksheet, date_cells)
        return None

    def _find_date_in_cells(self, worksheet, cell_list) -> datetime | None:
        for cell in cell_list:
            dt = self._extract_date_from_val(worksheet.Range(cell).Value)
            if dt:
                return dt
        return None

    def _normalize_model_string(self, s: Any) -> str:
        if s is None:
            return ""
        s_str = str(s)
        cleaned = re.sub(r"\s+", " ", s_str).strip()
        return re.sub(r"[\W_]+", "", cleaned).lower()

    def _extract_date_from_val(self, value: Any) -> datetime | None:
        if value is None or (isinstance(value, str) and not value.strip()):
            return None
        if isinstance(value, datetime):
            return value
        if isinstance(value, (int, float)):
            return datetime(1899, 12, 30) + timedelta(days=value)
        if isinstance(value, str):
            return self._parse_date_string(value)
        return None

    def _parse_date_string(self, value: str) -> datetime | None:
        # Formati diretti
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
            with suppress(ValueError):
                return datetime.strptime(value.strip(), fmt)
        # Estrazione regex
        match = re.search(r"(\d{2})[-/](\d{2})[-/](\d{4})", value)
        if match:
            with suppress(ValueError):
                return datetime.strptime(match.group(0).replace("/", "-"), "%d-%m-%Y")
        return None
