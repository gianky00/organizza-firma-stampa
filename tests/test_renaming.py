import os
from datetime import datetime
from unittest.mock import MagicMock

from src.logic.renaming import RenameProcessor


def test_rename_excel_files_success(fs, mock_gui, app_config):
    # Setup
    root_dir = "/schede"
    fs.create_dir(root_dir)
    file_path = os.path.join(root_dir, "testfile.xlsx")
    fs.create_file(file_path)
    app_config.rinomina_path.get.return_value = root_dir
    app_config.rinomina_password.get.return_value = ""

    # Mock Excel locale
    mock_excel_app = MagicMock()
    mock_wb = MagicMock()
    mock_ws = MagicMock()
    mock_excel_app.Workbooks.Open.return_value = mock_wb
    mock_wb.Worksheets.return_value = mock_ws

    mock_handler_class = MagicMock()
    mock_handler_class.return_value.__enter__.return_value = mock_excel_app

    # Mock Gateway locale
    mock_gateway_instance = MagicMock()
    mock_gateway_instance.extract_date_from_worksheet.return_value = datetime(2024, 5, 20)
    mock_gateway_instance.excel_handler_class = mock_handler_class

    mock_gateway_class = MagicMock(return_value=mock_gateway_instance)

    processor = RenameProcessor(
        mock_gui, app_config, MagicMock(), MagicMock(), MagicMock(), excel_gateway_class=mock_gateway_class
    )

    cancel_event = MagicMock()
    cancel_event.is_set.return_value = False
    processor.run_rename_process(cancel_event)

    found = False
    for f in os.listdir(root_dir):
        if "20-05-2024" in f:
            found = True
            break
    assert found
