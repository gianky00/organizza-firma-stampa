import os
import re
import subprocess
import traceback
from pathlib import Path

from src.utils import constants as const
from src.utils.excel_gateway import ExcelGateway
from src.utils.file_utils import clear_folder_content


class SignatureProcessor:
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
        self.logger = gui.log_firma
        self.setup_progress = setup_progress_cb
        self.update_progress = update_progress_cb
        self.hide_progress = hide_progress_cb
        self.excel_gateway = (excel_gateway_class or ExcelGateway)(self.logger)
        self.firma_processing_data = {
            "schedacontrolloSTRUMENTIANALOGICI": {"PrintArea": "A2:N55", "FirmaCella": "G54"},
            "schedacontrolloSTRUMENTIDIGITALI": {"PrintArea": "A2:N50", "FirmaCella": "G49"},
            "SchedacontrolloREPORTMANUTENZIONECORRETTIVA": {"PrintArea": "A2:N55", "FirmaCella": "G54"},
            "SCHEDAMANUTENZIONE": {"PrintArea": "A1:FV106", "FirmaCella": "FO104"},
        }

    def run_full_signature_process(self, cancel_event):
        self.logger("Avvio del processo di firma...", "HEADER")
        try:
            if not self._initialize_process():
                return

            excel_path = self.app_config.firma_excel_dir.get()
            excel_files = self._get_input_files(excel_path)
            if not excel_files:
                return

            self.gui.after(0, self.setup_progress, len(excel_files) * 2, "Processo di firma:")

            self.logger("--- FASE 1: Elaborazione Excel e Conversione PDF ---", "HEADER")
            processed_ok = self._process_excel_files(excel_files, cancel_event)

            if cancel_event.is_set() or not processed_ok:
                return

            self.logger("--- FASE 2: Compressione dei file PDF ---", "HEADER")
            self._compress_pdfs(cancel_event, len(excel_files))

            if not cancel_event.is_set():
                self.logger("--- PROCESSO DI FIRMA COMPLETATO ---", "SUCCESS")

        except Exception as e:
            self.logger(f"ERRORE CRITICO E IMPREVISTO: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            if cancel_event.is_set():
                self.logger("Processo di firma annullato.", "WARNING")
            self.gui.after(0, self.hide_progress)
            self.gui.after(0, self.gui.on_process_finished)

    def _initialize_process(self) -> bool:
        clear_folder_content(
            self.app_config.firma_pdf_dir.get(), self.logger, folder_display_name=const.FIRMA_PDF_OUTPUT_DIR
        )
        if not self._validate_paths():
            self.logger("Processo interrotto a causa di percorsi non validi.", "ERROR")
            return False
        return True

    def _get_input_files(self, excel_path: str) -> list[str]:
        if not Path(excel_path).is_dir():
            self.logger(f"ERRORE: Cartella non trovata: {excel_path}", "ERROR")
            return []
        files = [
            f
            for f in os.listdir(excel_path)
            if f.lower().endswith((".xlsx", ".xls", ".xlsm")) and not f.startswith("~")
        ]
        if not files:
            self.logger(f"Nessun file Excel da elaborare in: {excel_path}", "WARNING")
        return files

    def _validate_paths(self) -> bool:
        paths_to_check = {
            "Immagine Firma": self.app_config.firma_image_path.get(),
            "Eseguibile Ghostscript": self.app_config.firma_ghostscript_path.get(),
        }
        for name, path in paths_to_check.items():
            if not path or not Path(path).is_file():
                self.logger(f"ERRORE: '{name}' non trovato: {path}", "ERROR")
                return False
        return True

    def _process_excel_files(self, excel_files, cancel_event) -> bool:
        excel_path = self.app_config.firma_excel_dir.get()
        pdf_path = self.app_config.firma_pdf_dir.get()
        image_path = self.app_config.firma_image_path.get()
        mode = self.app_config.firma_processing_mode.get()
        
        errors = []
        for i, file_name in enumerate(excel_files):
            if cancel_event.is_set():
                return False
            self.gui.after(0, self.update_progress, i + 1)
            self.logger("-" * 50)
            self.logger(f"Elaborazione: {file_name}", "INFO")
            
            fp = os.path.join(excel_path, file_name)
            output_pdf = os.path.join(pdf_path, f"{Path(file_name).stem}.pdf")
            
            success, err = self._sign_and_export(fp, output_pdf, image_path, mode)
            if not success:
                errors.append((file_name, err))
                
        if errors:
            self._log_errors(errors)
        return not errors

    def _sign_and_export(self, excel_path, pdf_path, image_path, mode) -> tuple[bool, str | None]:
        # Qui usiamo il gateway per l'operazione atomica
        # Nota: SignatureProcessor ha logiche di posizionamento custom che il Gateway attuale non ha del tutto.
        # Estendiamo il Gateway o manteniamo qui la logica ma usando l'handler iniettato.
        # Per ora deleghiamo al gateway le operazioni base.
        
        from src.utils.excel_handler import ExcelHandler
        excel_h_class = getattr(self.excel_gateway, "excel_handler_class", ExcelHandler)
        
        with excel_h_class(self.logger) as excel:
            if not excel:
                return False, "Impossibile avvia r Excel"
            try:
                workbook = excel.Workbooks.Open(excel_path, 0, True)
                if workbook is None:
                    return False, "Apertura fallita: Workbook restituito come None."
                try:
                    if mode == "schede":
                        self._apply_signature_schede(workbook, pdf_path, image_path)
                    else:
                        self._apply_signature_preventivi(workbook, pdf_path, image_path)
                    return True, None
                finally:
                    workbook.Close(SaveChanges=False)
            except Exception as e:
                return False, str(e)

    def _apply_signature_schede(self, workbook, pdf_path, image_path):
        ws = workbook.Worksheets(1)
        val_e2 = ws.Cells(2, 5).Text.strip()
        val_t2 = ws.Cells(2, 20).Text.strip()
        val_t5 = ws.Cells(5, 20).Text.strip()
        model_value = val_e2 or val_t2 or val_t5
        cleaned_model = "".join(filter(str.isalnum, model_value))
        
        if cleaned_model in self.firma_processing_data:
            data = self.firma_processing_data[cleaned_model]
            ws.PageSetup.PrintArea = data["PrintArea"]
            
            # Dimensioni fisse immagine firma: specifiche richieste per il modello SCHEDAMANUTENZIONE
            # rispetto ad altri modelli generici
            img_width, img_height = (105, 35) if cleaned_model == "SCHEDAMANUTENZIONE" else (150, 50)
            
            cell_address = data["FirmaCella"]
            col_str = "".join(re.findall("[A-Z]+", cell_address))
            row_str = "".join(re.findall(r"\d+", cell_address))
            target_cell = ws.Cells(int(row_str), self._col_to_num(col_str))
            
            # 28.35 punti per centimetro in Excel
            points_per_cm = 28.35
            # Offset manuale per far combaciare l'immagine esattamente con l'area pre-stampata del modello
            offset_cm = 0.3 if cleaned_model == "SCHEDAMANUTENZIONE" else 1.0
            top_pos = max(0, target_cell.Top - (offset_cm * points_per_cm))
            left_pos = max(0, target_cell.Left - (1.0 * points_per_cm))
            
            ws.Shapes.AddPicture(image_path, True, True, left_pos, top_pos, img_width, img_height)
            workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
            self.logger("Firma applicata e PDF esportato.", "SUCCESS")
        else:
            self.logger(f"Modello non gestito: '{cleaned_model}'. Solo export PDF.", "WARNING")
            workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_path)

    def _apply_signature_preventivi(self, workbook, pdf_path, image_path):
        ws = next((s for s in workbook.Worksheets if s.Name == "Consuntivo"), None)
        if ws is None:
            self.logger("Foglio 'Consuntivo' non trovato. Export foglio attivo.", "WARNING")
            workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
            return
            
        ws.Activate()
        ws.PageSetup.PrintArea = "A3:L63"
        target_cell = ws.Cells(59, 3)
        ws.Shapes.AddPicture(image_path, True, True, target_cell.Left, target_cell.Top + 10, 150, 50)
        ws.ExportAsFixedFormat(0, pdf_path)
        self.logger("Firma applicata e PDF esportato.", "SUCCESS")

    def _compress_pdfs(self, cancel_event, progress_offset=0):
        pdf_path = self.app_config.firma_pdf_dir.get()
        pdf_files = [f for f in os.listdir(pdf_path) if f.lower().endswith(".pdf")]
        if not pdf_files:
            return
            
        gs_exe = self.app_config.firma_ghostscript_path.get()
        for i, pdf_file in enumerate(pdf_files):
            if cancel_event.is_set():
                break
            self.gui.after(0, self.update_progress, progress_offset + i + 1)
            self._compress_single_pdf(pdf_path, pdf_file, gs_exe)

    def _compress_single_pdf(self, pdf_path, file_name, gs_exe):
        input_pdf = Path(pdf_path) / file_name
        temp_pdf = Path(pdf_path) / f"temp_{file_name}"
        self.logger(f"Compressione: {file_name}", "INFO")
        
        args = [
            gs_exe, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", "-dPDFSETTINGS=/ebook",
            "-dNOPAUSE", "-dBATCH", "-dQUIET", f"-sOutputFile={temp_pdf}", str(input_pdf)
        ]
        try:
            subprocess.run(args, check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            if temp_pdf.exists() and temp_pdf.stat().st_size > 100:
                input_pdf.unlink()
                temp_pdf.rename(input_pdf)
                self.logger("Compressione OK.", "SUCCESS")
            else:
                if temp_pdf.exists(): temp_pdf.unlink()
        except Exception as e:
            self.logger(f"Errore compressione {file_name}: {e}\nComando: {' '.join(args)}", "ERROR")
            if isinstance(e, subprocess.CalledProcessError):
                self.logger(f"Dettagli errore Ghostscript: {e.stderr}", "ERROR")
            if temp_pdf.exists(): temp_pdf.unlink()

    def _log_errors(self, errors):
        self.logger("\n--- RIEPILOGO ERRORI ---", "HEADER")
        for file_name, error_msg in errors:
            self.logger(f"- {file_name}: {error_msg}", "ERROR")

    def _col_to_num(self, col_str):
        num = 0
        for char in col_str:
            num = num * 26 + (ord(char.upper()) - ord("A")) + 1
        return num
