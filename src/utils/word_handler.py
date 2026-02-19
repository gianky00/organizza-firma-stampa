from contextlib import suppress

import pythoncom
import win32com.client


class WordHandler:
    """
    Context manager for safely handling Word COM instance.
    """

    def __init__(self, logger):
        self.word = None
        self.logger = logger

    def __enter__(self):
        try:
            pythoncom.CoInitialize()
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = 0  # wdAlertsNone
            return self.word
        except Exception as e:
            self.logger(f"ERRORE FATALE: Impossibile avviare Word. Dettagli: {e}", "ERROR")
            with suppress(BaseException):
                pythoncom.CoUninitialize()
            return None

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.word:
            try:
                self.word.Quit(SaveChanges=0)  # wdDoNotSaveChanges
            except Exception as e:
                self.logger(f"Errore durante la chiusura di Word: {e}", "WARNING")
            finally:
                self.word = None
        with suppress(BaseException):
            pythoncom.CoUninitialize()
        return False
