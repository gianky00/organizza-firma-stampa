import os
import pytest
from src.utils.file_utils import clear_folder_content

def test_clear_folder_content_success(fs, mock_logger):
    # Setup filesystem virtuale
    test_dir = "/test_folder"
    fs.create_dir(test_dir)
    fs.create_file(os.path.join(test_dir, "file1.txt"))
    
    clear_folder_content(test_dir, mock_logger)
    
    assert os.path.exists(test_dir)
    assert len(os.listdir(test_dir)) == 0
    mock_logger.assert_any_call("--- Pulizia della cartella 'test_folder' in corso... ---", "HEADER")
    mock_logger.assert_any_call("--- Pulizia di 'test_folder' completata. ---", "SUCCESS")

def test_clear_folder_content_non_existent(fs, mock_logger):
    # Se il path non è una directory, il codice logga comunque HEADER e SUCCESS ma non entra nel loop
    clear_folder_content("/non_existent", mock_logger)
    mock_logger.assert_any_call("--- Pulizia della cartella 'non_existent' in corso... ---", "HEADER")
    mock_logger.assert_any_call("--- Pulizia di 'non_existent' completata. ---", "SUCCESS")

def test_clear_folder_content_item_error(fs, mock_logger, mocker):
    test_dir = "/protected"
    fs.create_dir(test_dir)
    fs.create_file(os.path.join(test_dir, "file.txt"))
    
    # Simula errore di cancellazione di un file specifico
    mocker.patch("os.remove", side_effect=PermissionError("Access denied"))
    
    clear_folder_content(test_dir, mock_logger)
    mock_logger.assert_any_call("Impossibile eliminare 'file.txt': Access denied", "ERROR")
