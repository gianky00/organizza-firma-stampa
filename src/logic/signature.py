import os
import re
import subprocess
import traceback
from src.utils.excel_handler import ExcelHandler

class SignatureProcessor:
    """
    Handles the logic for signing Excel files and converting them to compressed PDFs.
    """
    def __init__(self, gui, config, setup_progress_cb, update_progress_cb, hide_progress_cb):
        self.gui = gui
        self.config = config
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

    def run_full_signature_process(self):
        """
        Main entry point for the signature process. Called from the GUI thread.
        """
        self.logger("Avvio del processo di firma...", 'HEADER')
        try:
            if not self._validate_paths():
                self.logger("Processo interrotto a causa di percorsi non validi.", 'ERROR')
                return

            self.logger("--- FASE 1: Elaborazione Excel e Conversione PDF ---", 'HEADER')
            success = self._process_excel_files()

            if not success:
                self.logger("Fase 1 terminata con errori. Processo interrotto.", 'ERROR')
                return

            self.logger("--- FASE 2: Compressione dei file PDF ---", 'HEADER')
            self._compress_pdfs()
            self.logger("--- PROCESSO DI FIRMA COMPLETATO ---", 'SUCCESS')

        except Exception as e:
            self.logger(f"ERRORE CRITICO E IMPREVISTO: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            # The pythoncom initialization/uninitialization is now handled by the thread in the GUI
            # and the ExcelHandler context manager. We just need to re-enable the buttons.
            self.gui.after(0, self.hide_progress)
            self.gui.after(0, self.gui.toggle_firma_buttons, 'normal', 'normal') # Enable email button


    def _validate_paths(self):
        """Validates the necessary paths for the signature process."""
        paths_to_check = {
            "Immagine Firma": self.config.firma_image_path.get(),
            "Eseguibile Ghostscript": self.config.firma_ghostscript_path.get()
        }
        for name, path in paths_to_check.items():
            if not path or not os.path.isfile(path):
                self.logger(f"ERRORE: '{name}' non trovato: {path}", 'ERROR')
                return False
        return True

    def _process_excel_files(self):
        """
        Opens Excel, iterates through files, applies signatures, and exports to PDF.
        """
        excel_path = self.config.firma_excel_dir.get()
        if not os.path.isdir(excel_path):
            self.logger(f"Cartella input non trovata: '{excel_path}'", "ERROR")
            return False

        excel_files = [f for f in os.listdir(excel_path) if f.lower().endswith(('.xlsx', '.xls', '.xlsm')) and not f.startswith('~')]
        if not excel_files:
            self.logger(f"Nessun file Excel in: {self.config.FIRMA_EXCEL_INPUT_DIR}", 'WARNING')
            return True

        num_files = len(excel_files)
        self.logger(f"Trovati {num_files} file Excel da elaborare.")
        self.gui.after(0, self.setup_progress, num_files)

        errors = []
        with ExcelHandler(self.logger) as excel:
            if not excel:
                return False  # ExcelHandler already logged the error

            mode = self.config.firma_processing_mode.get()
            for i, file_name in enumerate(excel_files):
                self.gui.after(0, self.update_progress, i + 1)
                file_path = os.path.join(excel_path, file_name)
                self.logger("-" * 50)
                self.logger(f"Elaborazione: {file_name}", 'INFO')
                workbook = None
                try:
                    workbook = excel.Workbooks.Open(file_path, 0, True)
                    self.logger(f"  -> File '{file_name}' aperto con successo.", 'INFO')
                    if mode == "schede":
                        self._apply_signature_schede(workbook, file_name)
                    elif mode == "preventivi":
                        self._apply_signature_preventivi(workbook, file_name)
                except Exception as e:
                    error_msg = f"Impossibile aprire o elaborare il file. Dettagli: {e}"
                    self.logger(f"  -> ERRORE: {error_msg}", 'ERROR')
                    errors.append((file_name, error_msg))
                finally:
                    if workbook:
                        workbook.Close(SaveChanges=False)

        if errors:
            self.logger("\n--- RIEPILOGO ERRORI ---", "HEADER")
            for file_name, error_msg in errors:
                self.logger(f"- {file_name}: {error_msg}", "ERROR")

        return True

    def _apply_signature_schede(self, workbook, file_name):
        """Handles signing logic for 'schede' type documents."""
        try:
            ws = workbook.Worksheets(1)
            # Find model value from different cells
            valE2 = ws.Cells(2, 5).Text.strip()
            valT2 = ws.Cells(2, 20).Text.strip()
            valT5 = ws.Cells(5, 20).Text.strip()
            model_value = valE2 or valT2 or valT5
            cleaned_model = ''.join(filter(str.isalnum, model_value))
            self.logger(f"Modello: '{cleaned_model}'", 'INFO')

            if cleaned_model in self.firma_processing_data:
                data = self.firma_processing_data[cleaned_model]
                ws.PageSetup.PrintArea = data["PrintArea"]

                img_width, img_height = (105, 35) if cleaned_model == "SCHEDAMANUTENZIONE" else (150, 50)
                cell_address = data["FirmaCella"]
                col_str = ''.join(re.findall("[A-Z]+", cell_address))
                row_str = ''.join(re.findall(r"\d+", cell_address))
                target_cell = ws.Cells(int(row_str), self._col_to_num(col_str))

                points_per_cm = 28.35
                offset_1cm = 1.0 * points_per_cm
                offset_03cm = 0.3 * points_per_cm
                top_pos = max(0, target_cell.Top - (offset_03cm if cleaned_model == "SCHEDAMANUTENZIONE" else offset_1cm))
                left_pos = max(0, target_cell.Left - offset_1cm)

                ws.Shapes.AddPicture(self.config.firma_image_path.get(), True, True, left_pos, top_pos, img_width, img_height)
                self.logger("Immagine firma aggiunta.", 'SUCCESS')

                pdf_file_path = os.path.join(self.config.firma_pdf_dir.get(), f"{os.path.splitext(file_name)[0]}.pdf")
                workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_file_path)
                self.logger(f"File PDF esportato.", 'SUCCESS')
            else:
                self.logger(f"Modello non gestito: '{cleaned_model}'. File ignorato.", 'WARNING')
        except Exception as e:
            self.logger(f"ERRORE in _apply_signature_schede: {e}", 'ERROR')

    def _apply_signature_preventivi(self, workbook, file_name):
        """Handles signing logic for 'preventivi' type documents."""
        try:
            ws = next((s for s in workbook.Worksheets if s.Name == "Consuntivo"), None)
            if ws is None:
                self.logger("Foglio 'Consuntivo' non trovato. File ignorato.", 'WARNING')
                return

            ws.Activate()
            ws.PageSetup.PrintArea = "A3:L63"
            target_cell = ws.Cells(59, 3)
            top_position = target_cell.Top + 10
            ws.Shapes.AddPicture(self.config.firma_image_path.get(), True, True, target_cell.Left, top_position, 150, 50)
            self.logger("Immagine firma aggiunta.", 'SUCCESS')

            pdf_file_path = os.path.join(self.config.firma_pdf_dir.get(), f"{os.path.splitext(file_name)[0]}.pdf")
            ws.ExportAsFixedFormat(0, pdf_file_path)
            self.logger(f"File PDF esportato.", 'SUCCESS')
        except Exception as e:
            self.logger(f"ERRORE in _apply_signature_preventivi: {e}", 'ERROR')

    def _compress_pdfs(self):
        """Finds all PDFs in the output directory and compresses them using Ghostscript."""
        pdf_path = self.config.firma_pdf_dir.get()
        pdf_files = [f for f in os.listdir(pdf_path) if f.lower().endswith('.pdf')]
        if not pdf_files:
            self.logger("Nessun PDF da comprimere.", 'WARNING')
            return

        self.logger(f"Trovati {len(pdf_files)} PDF da comprimere.")
        gs_exe = self.config.firma_ghostscript_path.get()

        for pdf_file in pdf_files:
            input_pdf = os.path.join(pdf_path, pdf_file)
            temp_output_pdf = os.path.join(pdf_path, f"temp_{pdf_file}")
            self.logger(f"Compressione: {pdf_file}", 'INFO')
            args = [
                gs_exe,
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                "-dPDFSETTINGS=/ebook",
                "-dNOPAUSE",
                "-dBATCH",
                "-dQUIET",
                f"-sOutputFile={temp_output_pdf}",
                input_pdf
            ]
            try:
                # Using CREATE_NO_WINDOW to prevent flashing console windows
                subprocess.run(args, check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
                # Check if the compressed file is valid before replacing the original
                if os.path.exists(temp_output_pdf) and os.path.getsize(temp_output_pdf) > 100:
                    os.remove(input_pdf)
                    os.rename(temp_output_pdf, input_pdf)
                    self.logger("Compressione OK.", 'SUCCESS')
                else:
                    self.logger("ERRORE: File compresso non valido o vuoto. Operazione annullata.", 'ERROR')
                    if os.path.exists(temp_output_pdf):
                        os.remove(temp_output_pdf)
            except subprocess.CalledProcessError as e:
                self.logger(f"ERRORE durante la compressione con Ghostscript: {e.stderr}", 'ERROR')
                if os.path.exists(temp_output_pdf):
                    os.remove(temp_output_pdf)
            except Exception as e:
                self.logger(f"ERRORE imprevisto durante la compressione: {e}", 'ERROR')
                if os.path.exists(temp_output_pdf):
                    os.remove(temp_output_pdf)

    def _col_to_num(self, col_str):
        """Converts an Excel column letter (e.g., 'A', 'B', 'AA') to its 1-based number."""
        num = 0
        for char in col_str:
            num = num * 26 + (ord(char.upper()) - ord('A')) + 1
        return num
