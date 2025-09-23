import os
import sys

# Add the 'src' directory to the Python path
# This allows us to import modules from the src directory
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), 'src')))

from gui.main_window import MainApplication

def main():
    """
    Initializes and runs the main application.
    """
    # Before starting, ensure the necessary work directories exist.
    # This was previously in the old MainApplication's __init__.
    from utils import constants as const
    folders_to_create = [
        os.path.join(const.APPLICATION_PATH, const.FIRMA_EXCEL_INPUT_DIR),
        os.path.join(const.APPLICATION_PATH, const.FIRMA_PDF_OUTPUT_DIR),
        os.path.join(const.APPLICATION_PATH, const.ORGANIZZA_SOURCE_DIR),
        os.path.join(const.APPLICATION_PATH, const.ORGANIZZA_DEST_DIR),
        os.path.join(const.APPLICATION_PATH, const.RINOMINA_DEFAULT_DIR)
    ]
    try:
        for folder in folders_to_create:
            os.makedirs(folder, exist_ok=True)
    except Exception as e:
        # If we can't even create folders, it's a critical error.
        # A simple print is okay for a command-line fallback.
        print(f"ERRORE CRITICO: Impossibile creare le cartelle di lavoro: {e}")
        # In a GUI app, a popup would be better, but tkinter isn't running yet.
        return # Exit if we can't create folders

    app = MainApplication()
    app.mainloop()

if __name__ == "__main__":
    main()
