import os
import sys
import tkinter as tk
import traceback
from tkinter import messagebox

# Add the 'src' directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "src")))

from pathlib import Path


def log_critical_error(exc_type, exc_value, exc_traceback):
    """
    Scrive l'errore in un file fisico per debug se la GUI crasha.
    """
    error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    try:
        with open("CRASH_REPORT.txt", "w", encoding="utf-8") as f:
            f.write("--- CRASH REPORT - " + os.name + " ---\n")
            f.write(error_msg)
    except Exception:
        pass

    # Mostra anche un messaggio popup se possibile
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "ERRORE CRITICO",
            f"L'applicazione ha riscontrato un errore fatale e verrà chiusa.\nConsulta CRASH_REPORT.txt per i dettagli.\n\nErrore: {exc_value}",
        )
    except Exception:
        pass
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


sys.excepthook = log_critical_error


def main():
    """
    Initializes and runs the main application with a splash screen.
    """
    # 1. Show Splash Screen immediately
    # This gives immediate feedback to the user while heavy imports load.
    splash = tk.Tk()
    splash.overrideredirect(True)  # Borderless window

    # Calculate center position
    width, height = 300, 120
    screen_width = splash.winfo_screenwidth()
    screen_height = splash.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    splash.geometry(f"{width}x{height}+{x}+{y}")

    # Style the splash screen
    bg_color = "#f0f0f0"
    splash.configure(bg=bg_color)

    # Add a frame for a border effect
    frame = tk.Frame(splash, bg=bg_color, highlightbackground="#cccccc", highlightthickness=1)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="Gestione Documenti", font=("Segoe UI", 14, "bold"), bg=bg_color).pack(pady=(25, 5))
    tk.Label(frame, text="Caricamento in corso...", font=("Segoe UI", 10), bg=bg_color, fg="#666666").pack(pady=(0, 20))

    # Force the window to appear immediately
    splash.update()

    # 2. Perform initialization tasks (Create folders)
    # Lazy import of constants to avoid early dependency loading
    try:
        from utils import constants as const

        folders_to_create = [
            os.path.join(const.APPLICATION_PATH, const.FIRMA_EXCEL_INPUT_DIR),
            os.path.join(const.APPLICATION_PATH, const.FIRMA_PDF_OUTPUT_DIR),
            os.path.join(const.APPLICATION_PATH, const.ORGANIZZA_SOURCE_DIR),
            os.path.join(const.APPLICATION_PATH, const.ORGANIZZA_DEST_DIR),
            os.path.join(const.APPLICATION_PATH, const.RINOMINA_DEFAULT_DIR),
        ]
        for folder in folders_to_create:
            Path(folder).mkdir(parents=True, exist_ok=True)
    except Exception as e:
        splash.destroy()
        # Create a hidden root to show the error message properly
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Errore Critico", f"Impossibile creare le cartelle di lavoro:\n{e}")
        return

    # 3. Import MainApplication (Lazy Loading)
    # This is where the heavy lifting (imports of pandas, win32com, etc.) happens.
    try:
        from gui.main_window import MainApplication
    except Exception as e:
        splash.destroy()
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Errore di Avvio", f"Errore durante il caricamento dell'applicazione:\n{e}")
        return

    # 4. Destroy Splash and Launch App
    splash.destroy()
    app = MainApplication()
    app.mainloop()


if __name__ == "__main__":
    main()
