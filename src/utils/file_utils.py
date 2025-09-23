import os
import shutil

def clear_folder_content(folder_path, logger, folder_display_name=None):
    """
    Utility to clear all contents (files and subdirectories) of a given folder.

    Args:
        folder_path (str): The absolute path to the folder to clear.
        logger (function): A logging function to log messages.
        folder_display_name (str, optional): A user-friendly name for the folder for logging.
                                            Defaults to the folder's basename.
    """
    if folder_display_name is None:
        folder_display_name = os.path.basename(folder_path)

    logger(f"--- Pulizia della cartella '{folder_display_name}' in corso... ---", 'HEADER')
    if os.path.isdir(folder_path):
        for item_name in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item_name)
            try:
                if os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                else:
                    os.remove(item_path)
            except Exception as e:
                logger(f"Impossibile eliminare '{item_name}': {e}", 'ERROR')
    logger(f"--- Pulizia di '{folder_display_name}' completata. ---", 'SUCCESS')
