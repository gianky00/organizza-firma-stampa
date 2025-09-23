import pythoncom
import win32com.client
import traceback

class ExcelHandler:
    """
    A context manager to safely handle a single instance of the Excel application.
    Ensures that Excel is properly initialized and terminated.
    """
    def __init__(self, logger, visible=False, display_alerts=False):
        self.logger = logger
        self.visible = visible
        self.display_alerts = display_alerts
        self.excel_app = None

    def __enter__(self):
        """
        Initializes COM and starts the Excel application.
        Returns the Excel application object.
        """
        try:
            pythoncom.CoInitialize()
            self.excel_app = win32com.client.Dispatch("Excel.Application")
            self.excel_app.Visible = self.visible
            self.excel_app.DisplayAlerts = self.display_alerts
            self.logger("Applicazione Excel avviata in background.", "INFO")
            return self.excel_app
        except Exception as e:
            self.logger(f"ERRORE FATALE: Impossibile avviare l'applicazione Excel. Verificare che sia installata correttamente. Dettagli: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
            # Uninitialize if we failed to start
            pythoncom.CoUninitialize()
            return None

    def __exit__(self, exc_type, exc_val, exc_tb):
        """
        Quits the Excel application and uninitializes COM.
        """
        if self.excel_app:
            try:
                self.excel_app.Quit()
                self.logger("Applicazione Excel chiusa correttamente.", "INFO")
            except Exception as e:
                self.logger(f"ATTENZIONE: Si è verificato un errore durante la chiusura di Excel. Potrebbe rimanere un processo attivo. Dettagli: {e}", "WARNING")

        # Always uninitialize COM
        pythoncom.CoUninitialize()

        # Return False to propagate exceptions if they occurred inside the 'with' block
        return False
