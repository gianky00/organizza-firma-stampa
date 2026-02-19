import os
import pytest
from unittest.mock import MagicMock, patch
from src.logic.organization import OrganizationProcessor

@pytest.fixture
def mock_fees_processor():
    return MagicMock()

def test_organization_backup_creation(fs, mock_gui, app_config, mock_fees_processor):
    dest_dir = app_config.organizza_dest_dir.get.return_value
    fs.create_dir(dest_dir); fs.create_file(os.path.join(dest_dir, "f.txt"))
    processor = OrganizationProcessor(mock_gui, app_config, mock_fees_processor, MagicMock(), MagicMock(), MagicMock())
    with patch("src.logic.organization.shutil.copytree"), patch.object(processor, '_organize_files'):
        processor.run_organization_process(MagicMock())
        found = any("Creazione backup" in str(c) for c in mock_gui.log_organizza.call_args_list)
        assert found

def test_organization_flow_empty(fs, mock_gui, app_config, mock_fees_processor):
    source_dir = app_config.organizza_source_dir.get.return_value; fs.create_dir(source_dir)
    processor = OrganizationProcessor(mock_gui, app_config, mock_fees_processor, MagicMock(), MagicMock(), MagicMock())
    processor._organize_files(MagicMock())
    mock_gui.log_organizza.assert_any_call("Nessun file Excel trovato.", "WARNING")

def test_organization_process_files(fs, mock_gui, app_config, mock_fees_processor, mock_excel):
    source_dir = app_config.organizza_source_dir.get.return_value; fs.create_dir(source_dir)
    fs.create_file(os.path.join(source_dir, "S.xlsx"))
    excel_app, wb, ws = mock_excel
    ws.Range.return_value.Value = "12345"
    processor = OrganizationProcessor(mock_gui, app_config, mock_fees_processor, MagicMock(), MagicMock(), MagicMock())
    with patch("src.logic.organization.ExcelHandler") as h, patch("src.logic.organization.shutil.copy2") as c:
        h.return_value.__enter__.return_value = excel_app
        processor._organize_files(MagicMock())
        assert c.called

def test_organization_print_logic(fs, mock_gui, app_config, mock_fees_processor, mock_excel):
    d = "/dir"; fs.create_dir(d); fs.create_file(os.path.join(d, "F.xlsx"))
    excel_app, wb, ws = mock_excel
    ws.Cells.return_value.Value = "SCHEDAMANUTENZIONE"
    processor = OrganizationProcessor(mock_gui, app_config, mock_fees_processor, MagicMock(), MagicMock(), MagicMock())
    with patch("src.logic.organization.ExcelHandler") as h:
        h.return_value.__enter__.return_value = excel_app
        processor._print_files_in_folders(MagicMock(), [d])
        # Almeno verifichiamo che non crashi e apra il file
        assert excel_app.Workbooks.Open.called
