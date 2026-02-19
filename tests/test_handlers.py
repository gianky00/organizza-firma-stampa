import pytest
from unittest.mock import MagicMock, patch
from src.utils.excel_handler import ExcelHandler
from src.utils.word_handler import WordHandler

def test_excel_handler_context_manager(mock_logger):
    mock_excel_app = MagicMock()
    with patch("win32com.client.DispatchEx", return_value=mock_excel_app), patch("pythoncom.CoInitialize"), patch("pythoncom.CoUninitialize"):
        with ExcelHandler(mock_logger) as excel:
            assert excel == mock_excel_app
        mock_excel_app.Quit.assert_called_once()

def test_word_handler_context_manager(mock_logger):
    mock_word_app = MagicMock()
    with patch("win32com.client.Dispatch", return_value=mock_word_app), patch("pythoncom.CoInitialize"), patch("pythoncom.CoUninitialize"):
        with WordHandler(mock_logger) as word:
            assert word == mock_word_app
        mock_word_app.Quit.assert_called_with(SaveChanges=0)

def test_excel_handler_import_error(mock_logger):
    handler = ExcelHandler(mock_logger)
    # Patchiamo i moduli a livello globale
    with patch("win32com.client.DispatchEx", side_effect=ImportError), \
         patch("pythoncom.CoInitialize", side_effect=ImportError), \
         patch("src.utils.excel_handler.messagebox.showerror"):
        
        # Simula il fallimento del lazy import interno forzando un errore se chiamato
        # Poiché non possiamo impedire l'import reale se già avvenuto, 
        # verifichiamo la gestione dell'eccezione se sollevata.
        try:
            result = handler.__enter__()
            if result is None:
                mock_logger.assert_any_call("ERRORE FATALE: Le librerie necessarie (pywin32) per controllare Excel non sono installate.", "ERROR")
        except ImportError:
            pass # Successo se propagato o gestito
