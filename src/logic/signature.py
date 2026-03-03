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
            "schedacontrollostrumentianalogici": {"PrintArea": "A2:N55", "FirmaCella": "G54"},
            "schedacontrollostrumentidigitali": {"PrintArea": "A2:N50", "FirmaCella": "G49"},
            "schedacontrolloreportmanutenzionecorrettiva": {"PrintArea": "A2:N55", "FirmaCella": "G54"},
            "schedamanutenzione": {"PrintArea": "A1:FV106", "FirmaCella": "FO104"},
            "valvolediregolazione": {"PrintArea": "A1:DG110", "FirmaCella": "BU105"},
        }

    def run_full_signature_process(self, cancel_event):
        self.logger("Avvio del processo di firma...", "HEADER")
        self.prepared_files_data = []  # Lista di dict {pdf_path, tcl}
        try:
            # 0. Mostra barra di caricamento immediata
            self.gui.show_indeterminate("Inizializzazione ambiente...")

            # 1. Identificazione Area di Lavoro Locale e Sorgente
            local_work_path = os.path.join(const.APPLICATION_PATH, const.FIRMA_EXCEL_INPUT_DIR)
            source_path = self.app_config.firma_excel_dir.get()

            # 2. Logica di Importazione
            if os.path.normpath(source_path) != os.path.normpath(local_work_path):
                self.gui.show_indeterminate("Importazione file da rete...")
                self.logger(f"Importazione file da sorgente: {source_path}", "INFO")
                # Pulizia locale preventiva
                from src.utils.file_utils import clear_folder_content

                clear_folder_content(local_work_path, self.logger, folder_display_name="Area di Lavoro Locale")

                files_to_import = self._get_input_files(source_path)
                if not files_to_import:
                    return

                import shutil

                for f in files_to_import:
                    shutil.copy2(os.path.join(source_path, f), local_work_path)

                self.logger(f"Importati {len(files_to_import)} file Excel.", "SUCCESS")
                active_excel_path = local_work_path
            else:
                active_excel_path = source_path

            # 3. Inizializzazione (pulizia PDF)
            if not self._initialize_process():
                return

            excel_files = self._get_input_files(active_excel_path)
            if not excel_files:
                return

            # 4. Passaggio a barra di progresso determinata con ETA
            self.setup_progress(len(excel_files) * 2, "Elaborazione in corso:")

            self.logger("--- FASE 1: Elaborazione Excel e Conversione PDF ---", "HEADER")
            # Passiamo il percorso attivo alla funzione di elaborazione
            processed_ok = self._process_excel_files_at_path(active_excel_path, excel_files, cancel_event)

            if cancel_event.is_set() or not processed_ok:
                return

            self.logger("--- FASE 2: Compressione dei file PDF ---", "HEADER")
            self._compress_pdfs(cancel_event, len(excel_files))

            if not cancel_event.is_set():
                self.logger("--- PULIZIA: Spostamento originali in backup... ---", "INFO")
                backup_parent = os.path.join(const.APPLICATION_PATH, const.BACKUP_DIR, "Originali_Firmati")
                from src.utils.file_utils import clear_folder_content, create_backup

                if create_backup(active_excel_path, backup_parent_dir=backup_parent):
                    clear_folder_content(active_excel_path, self.logger, folder_display_name="Excel da Firmare")

                self.logger("--- PROCESSO DI FIRMA COMPLETATO ---", "SUCCESS")

        except Exception as e:
            self.logger(f"ERRORE CRITICO E IMPREVISTO: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            if cancel_event.is_set():
                self.logger("Processo di firma annullato.", "WARNING")
            self.hide_progress()

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

    def _process_excel_files_at_path(self, excel_path, excel_files, cancel_event) -> bool:
        pdf_path = self.app_config.firma_pdf_dir.get()
        image_path = self.app_config.firma_image_path.get()
        mode = self.app_config.firma_processing_mode.get()

        from src.utils.excel_handler import ExcelHandler

        excel_h_class = getattr(self.excel_gateway, "excel_handler_class", ExcelHandler)

        errors = []
        # TURBO: Apriamo Excel UNA SOLA VOLTA per l'intero lotto di file
        with excel_h_class(self.logger) as excel:
            if not excel:
                return False

            for i, file_name in enumerate(excel_files):
                if cancel_event.is_set():
                    return False
                self.update_progress(i + 1)
                self.logger("-" * 50)
                self.logger(f"Elaborazione: {file_name}", "INFO")

                fp = os.path.join(excel_path, file_name)
                output_pdf = os.path.join(pdf_path, f"{Path(file_name).stem}.pdf")

                try:
                    # Passiamo l'istanza excel già aperta per massima velocità
                    success, err, tcl_name = self._sign_and_export_with_instance(
                        excel, fp, output_pdf, image_path, mode
                    )
                    if success:
                        self.prepared_files_data.append({"path": output_pdf, "tcl": tcl_name or "N/D"})
                    else:
                        errors.append((file_name, err))
                except Exception as e:
                    errors.append((file_name, str(e)))

        if errors:
            self._log_errors(errors)
        return not errors

    def _sign_and_export_with_instance(
        self, excel, excel_path, pdf_path, image_path, mode
    ) -> tuple[bool, str | None, str | None]:
        try:
            # Apertura veloce: sola lettura, senza aggiornare link
            workbook = excel.Workbooks.Open(excel_path, 0, True)
            if workbook is None:
                return False, "Apertura fallita.", None
            try:
                # ESTRAZIONE TCL
                tcl_name = self.excel_gateway.extract_tcl_from_worksheet(workbook.Worksheets(1))

                if mode == "schede":
                    self._apply_signature_schede(workbook, pdf_path, image_path)
                else:
                    self._apply_signature_preventivi(workbook, pdf_path, image_path)
                return True, None, tcl_name
            finally:
                workbook.Close(SaveChanges=False)
        except Exception as e:
            return False, str(e), None

    def _apply_signature_schede(self, workbook, pdf_path, image_path):
        ws = workbook.Worksheets(1)

        # PRIORITÀ DI RICONOSCIMENTO: T3/T6 vincono su T2
        # Leggiamo le celle chiave
        cells_to_check = ["T3", "T6", "E2", "T2", "T5", "F2", "Q3", "S3", "N1"]

        cleaned_model = None
        matched_cell = None
        for cell_ref in cells_to_check:
            try:
                raw_val = ws.Range(cell_ref).Value
                norm_val = self.excel_gateway._normalize_model_string(raw_val)
                if norm_val in self.firma_processing_data:
                    cleaned_model = norm_val
                    matched_cell = cell_ref
                    break
            except Exception:
                continue

        if cleaned_model:
            data = self.firma_processing_data[cleaned_model]

            # LOGICA SPECIALE PRINT AREA PER VALVOLE DI REGOLAZIONE
            if cleaned_model == "valvolediregolazione":
                if matched_cell == "T3":
                    ws.PageSetup.PrintArea = "A1:FX106"
                elif matched_cell == "T6":
                    ws.PageSetup.PrintArea = "A4:FX109"
                else:
                    ws.PageSetup.PrintArea = data["PrintArea"]
            else:
                ws.PageSetup.PrintArea = data["PrintArea"]

            # Dimensioni fisse immagine firma: specifiche richieste per certi modelli
            if cleaned_model in ("schedamanutenzione", "valvolediregolazione"):
                img_width, img_height = (105, 35)
            else:
                img_width, img_height = (150, 50)

            cell_address = data["FirmaCella"]
            col_str = "".join(re.findall("[A-Z]+", cell_address))
            row_str = "".join(re.findall(r"\d+", cell_address))
            target_cell = ws.Cells(int(row_str), self._col_to_num(col_str))

            # 28.35 punti per centimetro in Excel
            points_per_cm = 28.35
            # Offset manuale per far combaciare l'immagine esattamente con l'area pre-stampata del modello
            offset_cm = 0.3 if cleaned_model == "schedamanutenzione" else 1.0
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

        # TURBO: Compressione parallela usando ThreadPoolExecutor
        import multiprocessing
        from concurrent.futures import ThreadPoolExecutor

        # Utilizziamo un numero di thread pari al numero di core (max 4 per non saturare IO)
        max_workers = min(multiprocessing.cpu_count(), 4)

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = []
            for pdf_file in pdf_files:
                if cancel_event.is_set():
                    break
                futures.append(executor.submit(self._compress_single_pdf, pdf_path, pdf_file, gs_exe))

            # Monitoraggio progresso
            for i, future in enumerate(futures):
                if cancel_event.is_set():
                    break
                future.result()  # Attende completamento
                self.update_progress(progress_offset + i + 1)

    def _compress_single_pdf(self, pdf_path, file_name, gs_exe):
        input_pdf = Path(pdf_path) / file_name
        temp_pdf = Path(pdf_path) / f"temp_{file_name}"
        self.logger(f"Compressione: {file_name}", "INFO")

        args = [
            gs_exe,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            "-dPDFSETTINGS=/ebook",
            "-dNOPAUSE",
            "-dBATCH",
            "-dQUIET",
            f"-sOutputFile={temp_pdf}",
            str(input_pdf),
        ]
        try:
            subprocess.run(args, check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            if temp_pdf.exists() and temp_pdf.stat().st_size > 100:
                input_pdf.unlink()
                temp_pdf.rename(input_pdf)
                self.logger("Compressione OK.", "SUCCESS")
            else:
                if temp_pdf.exists():
                    temp_pdf.unlink()
        except Exception as e:
            self.logger(f"Errore compressione {file_name}: {e}\nComando: {' '.join(args)}", "ERROR")
            if isinstance(e, subprocess.CalledProcessError):
                self.logger(f"Dettagli errore Ghostscript: {e.stderr}", "ERROR")
            if temp_pdf.exists():
                temp_pdf.unlink()

    def _log_errors(self, errors):
        self.logger("\n--- RIEPILOGO ERRORI ---", "HEADER")
        for file_name, error_msg in errors:
            self.logger(f"- {file_name}: {error_msg}", "ERROR")

    def _col_to_num(self, col_str):
        num = 0
        for char in col_str:
            num = num * 26 + (ord(char.upper()) - ord("A")) + 1
        return num
