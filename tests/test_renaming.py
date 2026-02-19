import os
import pytest
from unittest.mock import MagicMock, patch
from src.logic.renaming import RenameProcessor
from datetime import datetime

def test_rename_excel_files_success(fs, mock_gui, app_config):
    # 1. Setup Filesystem virtuale
    root_dir = "/schede"
    fs.create_dir(root_dir)
    # Usiamo un nome senza spazi perché RenameProcessor.replace(" ", "") li rimuove
    file_path = os.path.join(root_dir, "testfile.xlsx")
    fs.create_file(file_path)
    
    app_config.rinomina_path.get.return_value = root_dir
    
    # 2. Setup Mock Excel
    mock_excel_app = MagicMock()
    mock_wb = MagicMock()
    mock_ws = MagicMock()
    
    mock_excel_app.Workbooks.Open.return_value = mock_wb
    mock_wb.Worksheets.return_value = mock_ws
    
    # Simula valore della cella
    def range_side_effect(cell_ref):
        cell = MagicMock()
        if cell_ref == "AK2":
            cell.Value = datetime(2024, 5, 20)
        elif cell_ref == "N1":
            cell.Value = "Scheda Tecnica Verifica Disco Calibro"
        else:
            cell.Value = None
        return cell
        
    mock_ws.Range.side_effect = range_side_effect
    
    # 3. Esecuzione
    processor = RenameProcessor(mock_gui, app_config, MagicMock(), MagicMock(), MagicMock())
    
    with patch("src.logic.renaming.ExcelHandler") as mock_handler:
        mock_handler.return_value.__enter__.return_value = mock_excel_app
        
        cancel_event = MagicMock()
        cancel_event.is_set.return_value = False
        processor.run_rename_process(cancel_event)
    
    # 4. Verifiche
    expected_path = os.path.join(root_dir, "testfile (20-05-2024).xlsx")
    assert os.path.exists(expected_path)
