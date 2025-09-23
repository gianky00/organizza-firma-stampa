import pythoncom
import win32com.client
import traceback

class WordHandler:
    """
    A context manager to safely handle a single instance of the Word application.
    """
    def __init__(self, logger, visible=False):
        self.logger = logger
        self.visible = visible
        self.word_app = None

    def __enter__(self):
        """
        Initializes COM and starts the Word application.
        """
        try:
            # No need to CoInitialize here as the parent thread should do it.
            # But it's safe to call it multiple times.
            pythoncom.CoInitialize()
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = self.visible
            self.logger("Applicazione Word avviata in background.", "INFO")
            return self.word_app
        except Exception as e:
            self.logger(f"ERRORE FATALE: Impossibile avviare l'applicazione Word. Verificare che sia installata correttamente. Dettagli: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
            pythoncom.CoUninitialize()
            return None

    def __exit__(self, exc_type, exc_val, exc_tb):
        """
        Quits the Word application and uninitializes COM.
        """
        if self.word_app:
            try:
                self.word_app.Quit(SaveChanges=0)
                self.logger("Applicazione Word chiusa correttamente.", "INFO")
            except Exception as e:
                self.logger(f"ATTENZIONE: Si è verificato un errore durante la chiusura di Word. Potrebbe rimanere un processo attivo. Dettagli: {e}", "WARNING")

        pythoncom.CoUninitialize()
        return False
