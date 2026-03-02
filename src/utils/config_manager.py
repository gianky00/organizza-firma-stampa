import base64
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

        prev_month_date = datetime.now().replace(day=1) - timedelta(days=1)

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
            "canoni_tcl_list": [
                {"name": "Messina", "tcl": "MESSINA", "num": "036", "print": False},
                {"name": "Agusta", "tcl": "AGUSTA", "num": "007", "print": False},
                {"name": "Caldarella", "tcl": "CALDARELLA", "num": "034", "print": False},
                {"name": "Caldarella2 (manuale)", "tcl": "CALDARELLA", "num": "011", "print": False},
            ],
            "canoni_word_path": const.CANONI_WORD_DEFAULT_PATH,
            "canoni_macro_name": const.DEFAULT_MACRO_NAME,
            "selected_printer": "",
            "email_to": "",
            "email_cc": "",
            "email_subject": "Documenti Firmati",
            "email_tcl": "",
            "email_is_formal": False,
            "email_size_limit": "6",
            "canoni_giornaliera_base_dir": const.CANONI_GIORNALIERA_BASE_DIR,
            "canoni_consuntivi_base_dir": const.CANONI_CONSUNTIVI_BASE_DIR,
            "organizza_base_dir": const.ORGANIZZA_BASE_DIR,
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

                    # Migrazione automatica da 'canoni_referenti' a 'canoni_tcl_list'
                    if "canoni_referenti" in loaded_settings and "canoni_tcl_list" not in loaded_settings:
                        loaded_settings["canoni_tcl_list"] = loaded_settings.pop("canoni_referenti")

                    # Merge loaded settings with defaults to ensure all keys exist
                    self.settings = {**self.defaults, **loaded_settings}

                    if self.settings.get("rinomina_password"):
                        pw = self.settings["rinomina_password"]
                        if pw.startswith("b64:"):
                            try:
                                self.settings["rinomina_password"] = base64.b64decode(pw[4:].encode()).decode()
                            except Exception:
                                self.settings["rinomina_password"] = ""
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

        settings_copy = self.settings.copy()
        if settings_copy.get("rinomina_password"):
            settings_copy["rinomina_password"] = (
                "b64:" + base64.b64encode(settings_copy["rinomina_password"].encode()).decode()
            )

        try:
            with open(self.config_file_path, "w", encoding="utf-8") as f:
                json.dump(settings_copy, f, indent=4)
        except OSError:
            # In a real app, this might log to a status bar or a log file
            with suppress(OSError):
                pass

    def get(self, key):
        """
        Returns the value for a given configuration key.
        """
        return self.settings.get(key, self.defaults.get(key))
