import os
import pytest
from unittest.mock import MagicMock, patch
from src.logic.organization import OrganizationProcessor

def test_organization_backup_creation(fs, mock_gui, app_config):
    dest_dir = "/fake/dest"
    app_config.organizza_dest_dir.get.return_value = dest_dir
    fs.create_dir(dest_dir); fs.create_file(os.path.join(dest_dir, "f.txt"))
    processor = OrganizationProcessor(mock_gui, app_config, MagicMock(), MagicMock(), MagicMock(), MagicMock())
    with patch("src.logic.organization.shutil.copytree"), patch.object(processor, '_organize_files'):
        processor.run_organization_process(MagicMock())
        found = any("Creazione backup" in str(c) for c in mock_gui.log_organizza.call_args_list)
        assert found

def test_organization_flow_empty(fs, mock_gui, app_config):
    source_dir = "/fake/source"
    app_config.organizza_source_dir.get.return_value = source_dir
    fs.create_dir(source_dir)
    processor = OrganizationProcessor(mock_gui, app_config, MagicMock(), MagicMock(), MagicMock(), MagicMock())
    cancel_event = MagicMock(); cancel_event.is_set.return_value = False
    processor._organize_files(cancel_event)
    mock_gui.log_organizza.assert_any_call("Nessun file Excel trovato.", "WARNING")

def test_organization_process_files_full(fs, mock_gui, app_config):
    s_dir = "/src"; d_dir = "/dest"
    app_config.organizza_source_dir.get.return_value = s_dir
    app_config.organizza_dest_dir.get.return_value = d_dir
    fs.create_dir(s_dir); fs.create_file(os.path.join(s_dir, "S.xlsx"))
    
    # Mock Gateway locale
    mock_gateway_instance = MagicMock()
    mock_gateway_instance.get_odc_value.return_value = "ODC_TEST"
    
    mock_gateway_class = MagicMock(return_value=mock_gateway_instance)
    
    processor = OrganizationProcessor(
        mock_gui, app_config, MagicMock(), MagicMock(), MagicMock(), MagicMock(), 
        excel_gateway_class=mock_gateway_class
    )
    
    cancel_event = MagicMock(); cancel_event.is_set.return_value = False
    with patch("src.logic.organization.shutil.copy2") as mock_copy:
        processor._organize_files(cancel_event)
        assert mock_copy.called
        assert "ODC_TEST" in mock_copy.call_args[0][1]
