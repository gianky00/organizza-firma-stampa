import traceback
import tkinter as tk
from tkinter import messagebox

# Try to import the COM modules, but handle the error if they are not available.
try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None

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
        if not pythoncom or not win32com:
            self.logger("ERRORE FATALE: Le librerie necessarie (pywin32) per controllare Excel non sono installate.", "ERROR")
            messagebox.showerror(
                "Errore di Dipendenze",
                "Le librerie 'pywin32' necessarie per comunicare con Excel non sono installate. "
                "Si prega di installarle eseguendo 'pip install pywin32' da un terminale."
            )
            return None

        try:
            pythoncom.CoInitialize()
            self.excel_app = win32com.client.Dispatch("Excel.Application")
            self.excel_app.Visible = self.visible
            self.excel_app.DisplayAlerts = self.display_alerts
            self.logger("Applicazione Excel avviata in background.", "INFO")
            return self.excel_app
        except Exception as e:
            error_message = f"Impossibile avviare l'applicazione Excel. Verificare che sia installata correttamente. Dettagli: {e}"
            self.logger(f"ERRORE FATALE: {error_message}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
            messagebox.showerror("Errore Avvio Excel", error_message)
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
