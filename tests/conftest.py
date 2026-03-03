from unittest.mock import MagicMock

import pytest
from PySide6.QtWidgets import QWidget


@pytest.fixture
def mock_logger():
    """Fixture per catturare i log durante i test."""
    return MagicMock()


@pytest.fixture
def mock_gui():
    """Fixture per simulare la GUI di PySide6 con gli attributi necessari."""
    gui = MagicMock(spec=QWidget)
    gui.after = MagicMock()
    gui.on_process_finished = MagicMock()
    gui.show_indeterminate = MagicMock()
    # Loggers specifici
    gui.log_rinomina = MagicMock()
    gui.log_organizza = MagicMock()
    gui.log_firma = MagicMock()
    gui.log_canoni = MagicMock()
    return gui


@pytest.fixture
def app_config():
    """Fixture per simulare la configurazione dell'app."""
    config = MagicMock()
    # Setup return values come stringhe semplici per evitare fallimenti os.path
    config.rinomina_path.get.return_value = "/fake/path"
    config.rinomina_password.get.return_value = "secret"
    config.organizza_dest_dir.get.return_value = "/fake/dest"
    config.organizza_source_dir.get.return_value = "/fake/source"
    config.firma_pdf_dir.get.return_value = "/fake/pdf"
    config.firma_image_path.get.return_value = "/fake/img.png"
    config.firma_ghostscript_path.get.return_value = "/fake/gs.exe"
    config.firma_processing_mode.get.return_value = "schede"
    config.canoni_giornaliera_path.get.return_value = "/fake/giorn"
    config.canoni_word_path.get.return_value = "/fake/word.docx"
    return config
