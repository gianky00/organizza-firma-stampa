from __future__ import annotations
import os
import re
from datetime import datetime, timedelta
from typing import Optional, Any, TYPE_CHECKING

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

    def get_odc_value(self, file_path: str) -> Optional[str]:
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
                    Height=config.get("height", -1)
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

    def get_workbook_date(self, file_path: str, password: str = "") -> Optional[datetime]:
        """Apre un workbook (anche protetto) e tenta di estrarne la data di emissione."""
        from src.domain.models import DEFAULT_DATE_CANDIDATES, RENAME_MODELS

        with self.excel_handler_class(self.logger) as excel:
            if not excel:
                return None
            wb = None
            try:
                try:
                    wb = excel.Workbooks.Open(file_path, ReadOnly=True)
                except Exception:
                    if password:
                        wb = excel.Workbooks.Open(file_path, ReadOnly=True, Password=password)
                    else:
                        raise

                ws = wb.Worksheets(1)
                n1_val = self._normalize_model_string(ws.Range("N1").Value)
                model_config = RENAME_MODELS.get(n1_val)
                
                if model_config and "date_cells" in model_config:
                    for cell in model_config["date_cells"]:
                        dt = self._extract_date_from_val(ws.Range(cell).Value)
                        if dt:
                            return dt

                for cell in DEFAULT_DATE_CANDIDATES:
                    dt = self._extract_date_from_val(ws.Range(cell).Value)
                    if dt:
                        return dt
                
                return None
            except Exception as e:
                self.logger(f"Errore estrazione data da {os.path.basename(file_path)}: {e}", "ERROR")
                return None
            finally:
                if wb:
                    wb.Close(SaveChanges=False)

    def _normalize_model_string(self, s: Any) -> str:
        if s is None: return ""
        s_str = str(s)
        cleaned = re.sub(r"\s+", " ", s_str).strip()
        return re.sub(r"[\W_]+", "", cleaned).lower()

    def _extract_date_from_val(self, value: Any) -> Optional[datetime]:
        if value is None or (isinstance(value, str) and not value.strip()):
            return None
        if isinstance(value, datetime):
            return value
        if isinstance(value, (int, float)):
            return datetime(1899, 12, 30) + timedelta(days=value)
        if isinstance(value, str):
            for fmt in ["%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"]:
                try:
                    return datetime.strptime(value.strip(), fmt)
                except ValueError:
                    continue
            match = re.search(r"(\d{2})[-/](\d{2})[-/](\d{4})", value)
            if match:
                try:
                    return datetime.strptime(match.group(0).replace("/", "-"), "%d-%m-%Y")
                except ValueError:
                    pass
        return None
