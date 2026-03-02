from contextlib import suppress

import pythoncom
import win32com.client


class ExcelHandler:
    """
    Context manager for safely handling Excel COM instance.
    Optimized for high performance by reusing instance and disabling UI updates.
    """

    def __init__(self, logger):
        self.excel = None
        self.logger = logger
        self.co_initialized = False

    def __enter__(self):
        try:
            pythoncom.CoInitialize()
            self.co_initialized = True
            # DispatchEx ensures a fresh separate process, better for performance
            self.excel = win32com.client.DispatchEx("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False

            # TURBO MODE: Prova a disabilitare aggiornamenti pesanti, ma non crashare se Excel rifiuta
            try:
                self.excel.ScreenUpdating = False
                self.excel.EnableEvents = False
                # xlCalculationManual può fallire se Excel è in certi stati
                with suppress(Exception):
                    self.excel.Calculation = -4135  # xlCalculationManual
            except Exception as e:
                self.logger(f"Avviso: Impossibile ottimizzare completamente Excel (Turbo Mode): {e}", "DEBUG")

            return self.excel
        except ImportError:
            self.logger(
                "ERRORE FATALE: Le librerie necessarie (pywin32) per controllare Excel non sono installate.", "ERROR"
            )
            return None
        except Exception as e:
            self.logger(f"ERRORE FATALE: Impossibile avviare l'applicazione Excel. Dettagli: {e}", "ERROR")
            return None

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.excel:
            try:
                # Ripristina impostazioni prima di uscire
                with suppress(Exception):
                    self.excel.Calculation = -4105  # xlCalculationAutomatic
                    self.excel.ScreenUpdating = True
                    self.excel.EnableEvents = True
                self.excel.Quit()
            except Exception as e:
                self.logger(f"Errore durante la chiusura di Excel: {e}", "WARNING")
            finally:
                self.excel = None
        if self.co_initialized:
            with suppress(Exception):
                pythoncom.CoUninitialize()
