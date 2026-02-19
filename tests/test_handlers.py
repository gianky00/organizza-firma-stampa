from unittest.mock import MagicMock, patch

from src.utils.excel_handler import ExcelHandler
from src.utils.word_handler import WordHandler


def test_excel_handler_context_manager(mock_logger):
    mock_excel_app = MagicMock()
    with (
        patch("win32com.client.DispatchEx", return_value=mock_excel_app),
        patch("pythoncom.CoInitialize"),
        patch("pythoncom.CoUninitialize"),
    ):
        with ExcelHandler(mock_logger) as excel:
            assert excel == mock_excel_app
        mock_excel_app.Quit.assert_called_once()


def test_word_handler_context_manager(mock_logger):
    mock_word_app = MagicMock()
    with (
        patch("win32com.client.Dispatch", return_value=mock_word_app),
        patch("pythoncom.CoInitialize"),
        patch("pythoncom.CoUninitialize"),
    ):
        with WordHandler(mock_logger) as word:
            assert word == mock_word_app
        mock_word_app.Quit.assert_called_with(SaveChanges=0)
