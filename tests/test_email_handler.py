from unittest.mock import MagicMock, patch

from src.logic.email_handler import EmailHandler


def test_prepare_email_draft_logic(mock_logger):
    processor = EmailHandler(mock_logger)
    mock_outlook = MagicMock()
    mock_mail = MagicMock()
    mock_mail.HTMLBody = "-- Firma --"
    mock_outlook.CreateItem.return_value = mock_mail
    draft_data = {
        "to": "t",
        "cc": "c",
        "subject": "s",
        "intro_text": "i",
        "file_list": [{"name": "f", "tcl": "test"}],
        "attachments": ["a"],
    }
    with (
        patch("win32com.client.Dispatch", return_value=mock_outlook),
        patch("pythoncom.CoInitialize"),
        patch("pythoncom.CoUninitialize"),
    ):
        processor.create_outlook_draft(draft_data)
        assert mock_mail.To == "t"
        assert mock_mail.Attachments.Add.called
