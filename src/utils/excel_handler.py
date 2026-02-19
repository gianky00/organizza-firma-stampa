from contextlib import suppress
from tkinter import messagebox

import pythoncom
import win32com.client


class ExcelHandler:
    """
    Context manager for safely handling Excel COM instance.
    Ensures the instance is closed and resources are released.
    """

    def __init__(self, logger):
        self.excel = None
        self.logger = logger

    def __enter__(self):
        try:
            pythoncom.CoInitialize()
            self.excel = win32com.client.DispatchEx("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            return self.excel
        except ImportError:
            self.logger(
                "ERRORE FATALE: Le librerie necessarie (pywin32) per controllare Excel non sono installate.", "ERROR"
            )
            messagebox.showerror(
                "Errore di Sistema",
                "Le librerie necessarie (pywin32) per controllare Excel non sono installate.\n"
                "Eseguire 'pip install pywin32' dal terminale.",
            )
            return None
        except Exception as e:
            self.logger(f"ERRORE FATALE: Impossibile avviare l'applicazione Excel. Dettagli: {e}", "ERROR")
            messagebox.showerror(
                "Errore Excel",
                f"Impossibile avviare Excel. Assicurarsi che sia installato.\n\nDettagli: {e}",
            )
            return None

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.excel:
            try:
                self.excel.Quit()
            except Exception as e:
                self.logger(f"Errore durante la chiusura di Excel: {e}", "WARNING")
            finally:
                self.excel = None
        with suppress(ImportError):
            pythoncom.CoUninitialize()
