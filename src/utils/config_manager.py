import json
import os
from contextlib import suppress
from pathlib import Path

from src.utils import constants as const


class ConfigManager:
    """
    Manages application configuration, including loading from and saving to a JSON file.
    """

    def __init__(self, config_file_path=None):
        self.config_file_path = config_file_path or os.path.join(const.APPLICATION_PATH, const.CONFIG_FILE_NAME)
        self.defaults = self._get_defaults()
        self.settings = {}
        self.load()

    def _get_defaults(self):
        """
        Returns the default configuration values.
        """
        # Get the previous month's date for defaults
        from datetime import datetime, timedelta

        today = datetime.now()
        first_day_current_month = today.replace(day=1)
        prev_month_date = first_day_current_month - timedelta(days=1)

        return {
            "firma_excel_dir": os.path.join(const.APPLICATION_PATH, const.FIRMA_EXCEL_INPUT_DIR),
            "firma_pdf_dir": os.path.join(const.APPLICATION_PATH, const.FIRMA_PDF_OUTPUT_DIR),
            "firma_image_path": os.path.join(const.APPLICATION_PATH, "src", "assets", const.FIRMA_IMAGE_NAME),
            "firma_ghostscript_path": const.DEFAULT_GHOSTSCRIPT_PATH,
            "firma_processing_mode": "schede",
            "rinomina_path": os.path.join(const.APPLICATION_PATH, const.RINOMINA_DEFAULT_DIR),
            "rinomina_password": "",  # Lasciata vuota per sicurezza, l'utente dovrà inserirla
            "organizza_source_dir": os.path.join(const.APPLICATION_PATH, const.ORGANIZZA_SOURCE_DIR),
            "canoni_selected_year": str(prev_month_date.year),
            "canoni_selected_month": const.NOMI_MESI_ITALIANI[prev_month_date.month - 1],
            "canoni_messina_num": "",
            "canoni_naselli_num": "",
            "canoni_caldarella_num": "",
            "canoni_caldarella2_num": "",
            "canoni_messina_print": True,
            "canoni_naselli_print": True,
            "canoni_caldarella_print": True,
            "canoni_caldarella2_print": False,
            "canoni_word_path": const.CANONI_WORD_DEFAULT_PATH,
            "canoni_macro_name": const.DEFAULT_MACRO_NAME,
            "selected_printer": "",
            "email_to": "",
            "email_cc": "",
            "email_subject": "Documenti Firmati",
            "email_tcl": "",
            "email_is_formal": False,
            "email_size_limit": "6",
        }

    def load(self):
        """
        Loads settings from the JSON file. If the file doesn't exist or is invalid,
        it uses default values.
        """
        try:
            if Path(self.config_file_path).exists():
                with open(self.config_file_path, encoding="utf-8") as f:
                    loaded_settings = json.load(f)
                    # Merge loaded settings with defaults to ensure all keys exist
                    self.settings = {**self.defaults, **loaded_settings}
            else:
                self.settings = self.defaults
        except (OSError, json.JSONDecodeError):
            self.settings = self.defaults

    def save(self, settings_to_save=None):
        """
        Saves the current settings to the JSON file.
        """
        if settings_to_save:
            self.settings.update(settings_to_save)

        try:
            with open(self.config_file_path, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=4)
        except OSError:
            # In a real app, this might log to a status bar or a log file
            with suppress(OSError):
                pass

    def get(self, key):
        """
        Returns the value for a given configuration key.
        """
        return self.settings.get(key, self.defaults.get(key))
