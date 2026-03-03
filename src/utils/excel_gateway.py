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
        """Estrae la data identificando prima il modello corretto dalle configurazioni."""
        cfg = self.identify_model(ws)
        if cfg:
            date_cells = cfg.get("date_cells", [])
            return self._find_date_in_cells(ws, date_cells)

        # Fallback su candidati globali se nessun modello identificato
        from src.domain.models import DEFAULT_DATE_CANDIDATES

        return self._find_date_in_cells(ws, DEFAULT_DATE_CANDIDATES)

    def extract_tcl_from_worksheet(self, ws: Any) -> str | None:
        """Estrae il TCL identificando prima il modello corretto dalle configurazioni."""
        cfg = self.identify_model(ws)
        if cfg:
            tcl_cells = cfg.get("tcl_cells", [])
            for t_cell in tcl_cells:
                try:
                    tcl_val = ws.Range(t_cell).Value
                    if tcl_val:
                        return str(tcl_val).strip()
                except Exception:
                    continue
        return None

    def identify_model(self, ws: Any) -> dict | None:
        """Scansiona i modelli configurati e restituisce quello che corrisponde al foglio corrente."""
        models_config = self._get_dynamic_models_config()
        for cfg in models_config:
            mv = self._normalize_model_string(cfg.get("match_value", ""))
            if not mv:
                continue

            id_cells = cfg.get("id_cells", [])
            # Supporto per id_cell singola (legacy)
            if not id_cells and cfg.get("id_cell"):
                id_cells = [cfg.get("id_cell")]

            for cell_ref in id_cells:
                try:
                    raw_val = ws.Range(cell_ref).Value
                    if raw_val:
                        clean_val = self._normalize_model_string(raw_val)
                        if clean_val == mv:
                            return cfg
                except Exception:
                    continue
        return None

    def _get_dynamic_models_config(self) -> list[dict]:
        """Recupera i modelli dalla configurazione dell'app (file JSON)."""
        from src.utils.config_manager import ConfigManager

        models = ConfigManager().get("rename_models_config")
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
            # Supporta sia oggetti che dizionari
            is_dict = isinstance(cfg, dict)
            id_cells = cfg.get("id_cells") if is_dict else getattr(cfg, "id_cells", None)
            if not id_cells:
                id_cell = cfg.get("id_cell") if is_dict else getattr(cfg, "id_cell", None)
                id_cells = [id_cell] if id_cell else []

            match_value = cfg.get("match_value") if is_dict else getattr(cfg, "match_value", None)
            date_cells = cfg.get("date_cells") if is_dict else getattr(cfg, "date_cells", [])

            for id_cell in id_cells:
                try:
                    val = self._normalize_model_string(worksheet.Range(id_cell).Value)
                    if val == match_value:
                        return self._find_date_in_cells(worksheet, date_cells)
                except Exception:
                    continue
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
