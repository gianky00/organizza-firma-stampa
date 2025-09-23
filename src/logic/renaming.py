import os
import re
import pythoncom
import win32com.client
from datetime import datetime
import traceback

class RenameProcessor:
    """
    Handles the logic for renaming Excel files based on a date found within them.
    """
    def __init__(self, gui, config):
        self.gui = gui
        self.config = config
        self.logger = gui.log_rinomina  # Assumes the GUI provides a logging method for this tab

    def run_rename_process(self):
        """
        Main entry point for the renaming process.
        """
        root_path = self.config.rinomina_path.get()
        if not os.path.isdir(root_path):
            self.logger(f"ERRORE: Percorso non valido o inesistente: '{root_path}'", "ERROR")
            self.gui.after(0, self.gui.toggle_rinomina_buttons, 'normal')
            return

        pythoncom.CoInitialize()
        try:
            self._rename_excel_files_in_place(root_path)
        except Exception as e:
            self.logger(f"ERRORE CRITICO INASPETTATO: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            pythoncom.CoUninitialize()
            self.gui.after(0, self.gui.toggle_rinomina_buttons, 'normal')

    def _rename_excel_files_in_place(self, root_path):
        self.logger("[FASE 1/2] Raccolta file Excel...", "HEADER")
        excel_files = []
        for root, _, filenames in os.walk(root_path):
            for filename in filenames:
                if filename.lower().endswith(('.xlsx', '.xlsm', '.xls')) and not filename.startswith('~'):
                    excel_files.append(os.path.join(root, filename))

        if not excel_files:
            self.logger("Nessun file Excel trovato.", "WARNING")
            return

        self.logger(f"Trovati {len(excel_files)} file Excel. Inizio analisi.", "INFO")
        self.logger("[FASE 2/2] Analisi e ridenominazione...", "HEADER")

        DATE_IN_FILENAME_REGEX = re.compile(r'\s*\(\d{2}-\d{2}-\d{4}\)')
        corrected, already_ok, no_date_found, errors = 0, 0, 0, 0

        excel_app = None
        try:
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.DisplayAlerts = False

            for file_path in excel_files:
                self.logger(f"Analisi: {os.path.basename(file_path)}...")
                wb = None
                try:
                    # Attempt to open with and without password
                    try:
                        wb = excel_app.Workbooks.Open(file_path, ReadOnly=True)
                    except Exception:
                        self.logger(f"  -> File protetto. Tentativo con password 'coemi'...", "WARNING")
                        wb = excel_app.Workbooks.Open(file_path, ReadOnly=True, Password="coemi")

                    ws = wb.Worksheets(1)

                    # Cell extraction and date finding logic remains the same
                    all_cell_refs = ['F3', 'G3', 'C54', 'T6', 'AY3', 'B95', 'N1', 'AK2', 'Q3', 'S3', 'E2', 'F2', 'T2', 'T5', 'F56', 'F44', 'F49', 'B99', 'L46', 'B46', 'B108', 'B45', 'B50', 'B105', 'L52']
                    cell_values = {ref: ws.Range(ref).Value for ref in all_cell_refs}

                    model_n1 = self._normalize_model_string(cell_values.get('N1'))
                    model_f3 = self._normalize_model_string(cell_values.get('F3'))
                    model_g3 = self._normalize_model_string(cell_values.get('G3'))
                    model_ay3 = self._normalize_model_string(cell_values.get('AY3'))
                    model_q3 = self._normalize_model_string(cell_values.get('Q3'))
                    model_t6 = self._normalize_model_string(cell_values.get('T6'))
                    model_f2 = self._normalize_model_string(cell_values.get('F2'))
                    model_s3 = self._normalize_model_string(cell_values.get('S3'))
                    model_e2 = self._normalize_model_string(cell_values.get('E2'))
                    combined_e2_t2_t5 = model_e2 or self._normalize_model_string(cell_values.get('T2')) or self._normalize_model_string(cell_values.get('T5'))

                    date_candidates = []
                    if model_n1 == "schedatecnicaverificadiscocalibro": date_candidates = ["AK2"]
                    elif model_f3 == "schedavalvole" or model_g3 == "schedavalvole": date_candidates = ["C54"]
                    elif "valvole" in model_ay3: date_candidates = ["B95"]
                    elif model_q3 == "schedataraturastrumentidigitali": date_candidates = ["B50"]
                    elif model_t6 == "valvolediregolazione": date_candidates = ["B108"]
                    elif model_f2 == "schedacontrollovalvole": date_candidates = ["F56"]
                    elif model_f2 == "schedacontrollostrumentidigitali": date_candidates = ["F44"]
                    elif model_f2 == "schedacontrollostrumentianalogici": date_candidates = ["F49"]
                    elif model_f2 == "schedacontrollostrumenti": date_candidates = ["F49", "F44"]
                    elif model_s3 == "schedataraturastrumentodiprocesso": date_candidates = ["B99"]
                    elif combined_e2_t2_t5 == "schedacontrollovalvole": date_candidates = ["L46", "B46", "B108"]
                    elif combined_e2_t2_t5 == "schedacontrollostrumentidigitali": date_candidates = ["B45"]
                    elif model_e2 == "schedacontrollostrumentianalogici": date_candidates = ["L52", "B45", "B50", "B108", "B99", "B105"]
                    elif model_e2 == "schedacontrolloreportmanutenzionecorrettiva": date_candidates = ["B50"]
                    else: date_candidates = ["B45", "B50", "B108", "B99", "B105", "L52", "C54"]

                    emission_date = None
                    for cell_ref in date_candidates:
                        status, date_found = self._extract_date_from_val(cell_values.get(cell_ref))
                        if status == 'VALID':
                            emission_date = date_found
                            break

                    if emission_date:
                        original_dir, original_filename = os.path.split(file_path)
                        base_name, ext = os.path.splitext(original_filename)
                        cleaned_base_name = DATE_IN_FILENAME_REGEX.sub('', base_name).strip()
                        cleaned_base_name = self._clean_windows_duplicate_marker(cleaned_base_name)
                        new_filename = f"{cleaned_base_name} ({emission_date.strftime('%d-%m-%Y')}){ext}"

                        wb.Close(SaveChanges=False)
                        wb = None

                        if new_filename.lower() != original_filename.lower():
                            new_filepath = os.path.join(original_dir, new_filename)
                            final_path = self._get_unique_filepath(new_filepath)
                            os.rename(file_path, final_path)
                            self.logger(f"  -> RINOMINATO in: {os.path.basename(final_path)}", "SUCCESS")
                            corrected += 1
                        else:
                            self.logger("  -> Già corretto.", "INFO")
                            already_ok += 1
                    else:
                        self.logger("  -> Data non trovata. File non modificato.", "WARNING")
                        no_date_found += 1

                except Exception as e:
                    self.logger(f"--- ERRORE FILE: {os.path.basename(file_path)} ---", "ERROR")
                    self.logger(f"Tipo errore: {type(e).__name__} - Messaggio: {e}", "ERROR")
                    if "password" not in str(e).lower():
                        self.logger(traceback.format_exc(), "ERROR")
                    errors += 1
                finally:
                    if wb:
                        wb.Close(SaveChanges=False)
        finally:
            if excel_app:
                excel_app.Quit()

        self.logger("\n--- RIEPILOGO PROCESSO RINOMINA ---", "HEADER")
        self.logger(f"File rinominati o corretti: {corrected}", "SUCCESS")
        self.logger(f"File già corretti: {already_ok}", "INFO")
        self.logger(f"File con data non trovata: {no_date_found}", "WARNING")
        self.logger(f"File con errori di lettura: {errors}", "ERROR")
        self.logger("--- COMPLETATO ---", "HEADER")

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
        return re.sub(r'\s*\(\d+\)$', '', name.strip())

    def _normalize_model_string(self, s: any) -> str:
        if s is None:
            return ""
        s_str = str(s)
        cleaned = re.sub(r'\s+', ' ', s_str).strip()
        return re.sub(r'[\W_]+', '', cleaned).lower()

    def _extract_date_from_val(self, value: any) -> tuple[str, datetime | None]:
        if value is None or (isinstance(value, str) and not value.strip()):
            return 'EMPTY', None

        if hasattr(value, 'year') and hasattr(value, 'month') and hasattr(value, 'day'):
            try:
                return 'VALID', datetime(value.year, value.month, value.day)
            except Exception:
                pass

        if isinstance(value, str):
            date_str = value.strip()
            if not date_str:
                return 'EMPTY', None

            range_match = re.match(r'^\d{1,2}\s*-\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', date_str)
            if range_match:
                date_str = range_match.group(1)

            date_str = date_str.split('&')[0].strip()
            if not date_str:
                return 'EMPTY', None

            date_formats = (
                "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y",
                "%d/%m/%y", "%m/%d/%y", "%y-%m-%d", "%d-%m-%y", "%d.%m.%y"
            )
            for fmt in date_formats:
                try:
                    dt_obj = datetime.strptime(date_str, fmt)
                    if dt_obj.year < 100:
                        current_year_base = datetime.now().year // 100 * 100
                        year_adjusted = current_year_base + dt_obj.year
                        if year_adjusted > datetime.now().year + 20:
                            year_adjusted -= 100
                        dt_obj = dt_obj.replace(year=year_adjusted)
                    return 'VALID', dt_obj
                except ValueError:
                    continue
            return 'TYPO', None
        return 'EMPTY', None
