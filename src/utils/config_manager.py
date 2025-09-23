import json
import os
from datetime import datetime
from . import constants as const

class ConfigManager:
    """
    Manages loading and saving the application's configuration file.
    """
    def __init__(self):
        self.config_path = os.path.join(const.APPLICATION_PATH, const.CONFIG_FILE_NAME)
        self.settings = {}
        self._load_defaults()

    def _load_defaults(self):
        """
        Sets the default values for all configuration settings.
        """
        self.defaults = {
            "firma_ghostscript_path": const.DEFAULT_GHOSTSCRIPT_PATH,
            "rinomina_path": os.path.join(const.APPLICATION_PATH, const.RINOMINA_DEFAULT_DIR),
            "rinomina_password": "coemi", # Default password
            "organizza_source_dir": os.path.join(const.APPLICATION_PATH, const.ORGANIZZA_SOURCE_DIR),
            "canoni_selected_year": str(datetime.now().year),
            "canoni_selected_month": const.NOMI_MESI_ITALIANI[datetime.now().month - 1],
            "canoni_messina_num": "",
            "canoni_naselli_num": "",
            "canoni_caldarella_num": "",
            "canoni_word_path": const.CANONI_WORD_DEFAULT_PATH,
            "selected_printer": "",
            "email_to": "",
            "email_subject": "Documenti Firmati",
            "email_tcl": "",
            "email_is_formal": False
        }

    def load(self):
        """
        Loads settings from the JSON file. If the file doesn't exist or is invalid,
        it falls back to the default settings.
        """
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r') as f:
                    self.settings = json.load(f)
            else:
                self.settings = self.defaults
        except (json.JSONDecodeError, IOError) as e:
            print(f"Error loading {const.CONFIG_FILE_NAME}: {e}. Using default values.")
            self.settings = self.defaults

    def get(self, key):
        """
        Gets a value from the loaded settings, falling back to the default if not found.
        """
        return self.settings.get(key, self.defaults.get(key))

    def save(self, current_config):
        """
        Saves the provided dictionary of current settings to the JSON file.

        Args:
            current_config (dict): A dictionary with the latest values from the GUI.
        """
        try:
            with open(self.config_path, 'w') as f:
                json.dump(current_config, f, indent=4)
        except IOError as e:
            # In a real app, this might log to a status bar or a log file
            print(f"Error saving configuration: {e}")
