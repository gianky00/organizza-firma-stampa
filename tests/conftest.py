import pytest
from unittest.mock import MagicMock, patch, PropertyMock
import tkinter as tk

@pytest.fixture
def mock_logger():
    return MagicMock()

@pytest.fixture
def mock_gui():
    gui = MagicMock(spec=tk.Tk)
    gui.after = MagicMock()
    gui.on_process_finished = MagicMock()
    gui.log_rinomina = MagicMock()
    gui.log_organizza = MagicMock()
    gui.log_firma = MagicMock()
    gui.log_canoni = MagicMock()
    return gui

@pytest.fixture
def app_config():
    config = MagicMock()
    config.rinomina_path.get.return_value = "/fake/path"
    config.rinomina_password.get.return_value = "secret"
    config.organizza_dest_dir.get.return_value = "/fake/dest"
    config.organizza_source_dir.get.return_value = "/fake/source"
    config.firma_pdf_dir.get.return_value = "/fake/pdf"
    config.firma_image_path.get.return_value = "/fake/assets/TIMBRO.png"
    config.firma_ghostscript_path.get.return_value = "/fake/gs.exe"
    config.firma_processing_mode.get.return_value = "schede"
    config.canoni_giornaliera_path.get.return_value = "/fake/giorn"
    config.canoni_word_path.get.return_value = "/fake/word.docx"
    return config

@pytest.fixture
def mock_excel():
    excel = MagicMock()
    wb = MagicMock()
    ws = MagicMock()
    excel.Workbooks.Open.return_value = wb
    wb.Worksheets.return_value = ws
    wb.Worksheets.__getitem__.return_value = ws
    # Supporta iterazione
    wb.Worksheets = [ws]
    # Configura property Name per Worksheet
    type(ws).Name = PropertyMock(return_value="Sheet1")
    return excel, wb, ws
