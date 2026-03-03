import os
import sys
import traceback
from pathlib import Path

# Add the 'src' directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "src")))

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QApplication, QMessageBox, QSplashScreen


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
        app = QApplication.instance()
        if not app:
            app = QApplication(sys.argv)
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setWindowTitle("ERRORE CRITICO")
        msg_box.setText(
            "L'applicazione ha riscontrato un errore fatale e verrà chiusa.\nConsulta CRASH_REPORT.txt per i dettagli."
        )
        msg_box.setDetailedText(f"Errore: {exc_value}")
        msg_box.exec()
    except Exception:
        pass
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


sys.excepthook = log_critical_error


def create_splash_pixmap():
    width, height = 400, 200
    pixmap = QPixmap(width, height)
    pixmap.fill(QColor("#f0f0f0"))

    painter = QPainter(pixmap)

    # Draw border
    pen = QPen(QColor("#cccccc"))
    pen.setWidth(2)
    painter.setPen(pen)
    painter.drawRect(1, 1, width - 2, height - 2)

    # Draw title
    font_title = QFont("Segoe UI", 16, QFont.Bold)
    painter.setFont(font_title)
    painter.setPen(QColor("#333333"))
    painter.drawText(0, 0, width, height // 2 + 20, Qt.AlignCenter, "Gestione Documenti")

    # Draw subtitle
    font_subtitle = QFont("Segoe UI", 11)
    painter.setFont(font_subtitle)
    painter.setPen(QColor("#666666"))
    painter.drawText(
        0, height // 2 + 20, width, height // 2 - 20, Qt.AlignHCenter | Qt.AlignTop, "Caricamento in corso..."
    )

    painter.end()
    return pixmap


def main():
    """
    Initializes and runs the main application with a splash screen.
    """
    app = QApplication(sys.argv)

    # 1. Show Splash Screen immediately
    splash_pixmap = create_splash_pixmap()
    splash = QSplashScreen(splash_pixmap, Qt.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()

    # 2. Perform initialization tasks (Create folders)
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
        splash.finish(None)
        QMessageBox.critical(None, "Errore Critico", f"Impossibile creare le cartelle di lavoro:\n{e}")
        return

    # 3. Import MainApplication (Lazy Loading)
    try:
        from gui.main_window import MainApplication
    except Exception as e:
        splash.finish(None)
        QMessageBox.critical(
            None,
            "Errore di Avvio",
            f"Errore durante il caricamento dell'applicazione:\n{e}\n\n{traceback.format_exc()}",
        )
        return

    # 4. Launch App and Destroy Splash
    main_app = MainApplication()
    main_app.show()
    splash.finish(main_app)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
