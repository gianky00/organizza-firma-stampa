import os
import re
import subprocess
import traceback
from src.utils import constants as const
from src.utils.excel_handler import ExcelHandler
from src.utils.file_utils import clear_folder_content


class SignatureProcessor:
    def __init__(self, gui, app_config, setup_progress_cb, update_progress_cb, hide_progress_cb):
        self.gui = gui
        self.app_config = app_config
        self.logger = gui.log_firma
        self.setup_progress = setup_progress_cb
        self.update_progress = update_progress_cb
        self.hide_progress = hide_progress_cb
        self.firma_processing_data = {
            "schedacontrolloSTRUMENTIANALOGICI": {"PrintArea": "A2:N55", "FirmaCella": "G54"},
            "schedacontrolloSTRUMENTIDIGITALI": {"PrintArea": "A2:N50", "FirmaCella": "G49"},
            "SchedacontrolloREPORTMANUTENZIONECORRETTIVA": {"PrintArea": "A2:N55", "FirmaCella": "G55"},
            "SCHEDAMANUTENZIONE": {"PrintArea": "A1:FV106", "FirmaCella": "BZ105"},
        }

    def run_full_signature_process(self, cancel_event):
        self.logger("Avvio del processo di firma...", 'HEADER')
        try:
            clear_folder_content(
                self.app_config.firma_pdf_dir.get(),
                self.logger,
                folder_display_name=const.FIRMA_PDF_OUTPUT_DIR
            )
            if not self._validate_paths():
                self.logger("Processo interrotto a causa di percorsi non validi.", 'ERROR')
                return

            excel_path = self.app_config.firma_excel_dir.get()
            excel_files = [f for f in os.listdir(excel_path) if f.lower().endswith(('.xlsx', '.xls', '.xlsm')) and not f.startswith('~')]
            total_steps = len(excel_files) * 2
            self.gui.after(0, self.setup_progress, total_steps)

            if cancel_event.is_set(): return

            self.logger("--- FASE 1: Elaborazione Excel e Conversione PDF ---", 'HEADER')
            processed_ok = self._process_excel_files(excel_files, cancel_event)

            if cancel_event.is_set(): return
            if not processed_ok:
                self.logger("Fase 1 terminata con errori. Processo interrotto.", 'ERROR')
                return

            self.logger("--- FASE 2: Compressione dei file PDF ---", 'HEADER')
            self._compress_pdfs(cancel_event, len(excel_files))

            if not cancel_event.is_set():
                self.logger("--- PROCESSO DI FIRMA COMPLETATO ---", 'SUCCESS')

        except Exception as e:
            self.logger(f"ERRORE CRITICO E IMPREVISTO: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            if cancel_event.is_set(): self.logger("Processo di firma annullato.", "WARNING")
            self.gui.after(0, self.hide_progress)
            self.gui.after(0, self.gui.on_process_finished)

    def _validate_paths(self):
        paths_to_check = {
            "Immagine Firma": self.app_config.firma_image_path.get(),
            "Eseguibile Ghostscript": self.app_config.firma_ghostscript_path.get()
        }
        for name, path in paths_to_check.items():
            if not path or not os.path.isfile(path):
                self.logger(f"ERRORE: '{name}' non trovato: {path}", 'ERROR')
                return False
        return True

    def _process_excel_files(self, excel_files, cancel_event):
        excel_path = self.app_config.firma_excel_dir.get()
        if not excel_files:
            self.logger(f"Nessun file Excel da elaborare in: {const.FIRMA_EXCEL_INPUT_DIR}", 'WARNING')
            return True
        self.logger(f"Inizio elaborazione di {len(excel_files)} file Excel...")
        errors = []
        with ExcelHandler(self.logger) as excel:
            if not excel: return False
            mode = self.app_config.firma_processing_mode.get()
            for i, file_name in enumerate(excel_files):
                if cancel_event.is_set(): return False
                self.gui.after(0, self.update_progress, i + 1)
                file_path = os.path.join(excel_path, file_name)
                self.logger("-" * 50)
                self.logger(f"Elaborazione: {file_name}", 'INFO')
                workbook = None
                try:
                    workbook = excel.Workbooks.Open(file_path, 0, True)
                    self.logger(f"  -> File '{file_name}' aperto con successo.", 'INFO')
                    if mode == "schede": self._apply_signature_schede(workbook, file_name)
                    elif mode == "preventivi": self._apply_signature_preventivi(workbook, file_name)
                except Exception as e:
                    errors.append((file_name, f"Impossibile aprire o elaborare il file. Dettagli: {e}"))
                finally:
                    if workbook: workbook.Close(SaveChanges=False)
        if errors:
            self.logger("\n--- RIEPILOGO ERRORI ---", "HEADER")
            for file_name, error_msg in errors: self.logger(f"- {file_name}: {error_msg}", "ERROR")
        return not errors

    def _apply_signature_schede(self, workbook, file_name):
        try:
            ws = workbook.Worksheets(1)
            valE2 = ws.Cells(2, 5).Text.strip(); valT2 = ws.Cells(2, 20).Text.strip(); valT5 = ws.Cells(5, 20).Text.strip()
            model_value = valE2 or valT2 or valT5
            cleaned_model = ''.join(filter(str.isalnum, model_value))
            if cleaned_model in self.firma_processing_data:
                data = self.firma_processing_data[cleaned_model]
                ws.PageSetup.PrintArea = data["PrintArea"]
                img_width, img_height = (105, 35) if cleaned_model == "SCHEDAMANUTENZIONE" else (150, 50)
                cell_address = data["FirmaCella"]
                col_str = ''.join(re.findall("[A-Z]+", cell_address)); row_str = ''.join(re.findall(r"\d+", cell_address))
                target_cell = ws.Cells(int(row_str), self._col_to_num(col_str))
                points_per_cm = 28.35; offset_1cm = 1.0 * points_per_cm; offset_03cm = 0.3 * points_per_cm
                top_pos = max(0, target_cell.Top - (offset_03cm if cleaned_model == "SCHEDAMANUTENZIONE" else offset_1cm))
                left_pos = max(0, target_cell.Left - offset_1cm)
                ws.Shapes.AddPicture(self.app_config.firma_image_path.get(), True, True, left_pos, top_pos, img_width, img_height)
                pdf_file_path = os.path.join(self.app_config.firma_pdf_dir.get(), f"{os.path.splitext(file_name)[0]}.pdf")
                workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_file_path)
                self.logger("Firma applicata e PDF esportato.", 'SUCCESS')
            else: self.logger(f"Modello non gestito: '{cleaned_model}'. File ignorato.", 'WARNING')
        except Exception as e: self.logger(f"ERRORE in _apply_signature_schede: {e}", 'ERROR')

    def _apply_signature_preventivi(self, workbook, file_name):
        try:
            ws = next((s for s in workbook.Worksheets if s.Name == "Consuntivo"), None)
            if ws is None: self.logger("Foglio 'Consuntivo' non trovato.", 'WARNING'); return
            ws.Activate(); ws.PageSetup.PrintArea = "A3:L63"; target_cell = ws.Cells(59, 3); top_position = target_cell.Top + 10
            ws.Shapes.AddPicture(self.app_config.firma_image_path.get(), True, True, target_cell.Left, top_position, 150, 50)
            pdf_file_path = os.path.join(self.app_config.firma_pdf_dir.get(), f"{os.path.splitext(file_name)[0]}.pdf")
            ws.ExportAsFixedFormat(0, pdf_file_path)
            self.logger("Firma applicata e PDF esportato.", 'SUCCESS')
        except Exception as e: self.logger(f"ERRORE in _apply_signature_preventivi: {e}", 'ERROR')

    def _compress_pdfs(self, cancel_event, progress_offset=0):
        pdf_path = self.app_config.firma_pdf_dir.get()
        pdf_files = [f for f in os.listdir(pdf_path) if f.lower().endswith('.pdf')]
        if not pdf_files: self.logger("Nessun PDF da comprimere.", 'WARNING'); return
        self.logger(f"Trovati {len(pdf_files)} PDF da comprimere.")
        gs_exe = self.app_config.firma_ghostscript_path.get()
        for i, pdf_file in enumerate(pdf_files):
            if cancel_event.is_set(): return
            self.gui.after(0, self.update_progress, progress_offset + i + 1)
            input_pdf = os.path.join(pdf_path, pdf_file)
            temp_output_pdf = os.path.join(pdf_path, f"temp_{pdf_file}")
            self.logger(f"Compressione: {pdf_file}", 'INFO')
            args = [gs_exe, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", "-dPDFSETTINGS=/ebook", "-dNOPAUSE", "-dBATCH", "-dQUIET", f"-sOutputFile={temp_output_pdf}", input_pdf]
            try:
                subprocess.run(args, check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
                if os.path.exists(temp_output_pdf) and os.path.getsize(temp_output_pdf) > 100:
                    os.remove(input_pdf); os.rename(temp_output_pdf, input_pdf)
                    self.logger("Compressione OK.", 'SUCCESS')
                else:
                    self.logger("ERRORE: File compresso non valido.", 'ERROR')
                    if os.path.exists(temp_output_pdf): os.remove(temp_output_pdf)
            except subprocess.CalledProcessError as e:
                self.logger(f"ERRORE Ghostscript: {e.stderr}", 'ERROR')
                if os.path.exists(temp_output_pdf): os.remove(temp_output_pdf)
            except Exception as e:
                self.logger(f"ERRORE imprevisto compressione: {e}", 'ERROR')
                if os.path.exists(temp_output_pdf): os.remove(temp_output_pdf)

    def _col_to_num(self, col_str):
        num = 0
        for char in col_str: num = num * 26 + (ord(char.upper()) - ord('A')) + 1
        return num
