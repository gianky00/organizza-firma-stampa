import os
from unittest.mock import patch

from src.utils.file_utils import clear_folder_content, create_backup


def test_clear_folder_content_success(fs, mock_logger):
    test_dir = "/test_clear"
    fs.create_dir(test_dir)
    fs.create_file(os.path.join(test_dir, "file1.txt"))
    fs.create_dir(os.path.join(test_dir, "subdir"))

    clear_folder_content(test_dir, mock_logger)

    assert os.path.exists(test_dir)
    assert len(os.listdir(test_dir)) == 0
    mock_logger.assert_any_call(f"--- Pulizia di '{os.path.basename(test_dir)}' completata. ---", "SUCCESS")


def test_clear_folder_content_non_existent(mock_logger):
    clear_folder_content("/non/existent", mock_logger)
    mock_logger.assert_any_call("--- Pulizia di 'existent' completata. ---", "SUCCESS")


def test_clear_folder_content_item_error(fs, mock_logger):
    test_dir = "/protected"
    fs.create_dir(test_dir)
    fs.create_file(os.path.join(test_dir, "file.txt"))

    with patch("os.remove", side_effect=PermissionError("Access denied")):
        clear_folder_content(test_dir, mock_logger)
        mock_logger.assert_any_call("Impossibile eliminare 'file.txt': Access denied", "ERROR")


def test_create_backup_success(fs):
    source = "/src_data"
    fs.create_dir(source)
    fs.create_file(source + "/f.txt")
    result = create_backup(source)
    assert result is True
    backups = [d for d in os.listdir("/") if "_backup_" in d]
    assert len(backups) > 0


def test_create_backup_fail(fs):
    assert create_backup("/non_existent") is False
