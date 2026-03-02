from __future__ import annotations

import os
from contextlib import suppress
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    pass

from src.utils.word_handler import WordHandler


class WordGateway:
    """
    Astrae le operazioni su Microsoft Word (COM).
    Segue il Gateway Pattern per isolare la logica di business.
    """

    def __init__(self, logger, word_handler_class=None):
        self.logger = logger
        self.word_handler_class = word_handler_class or WordHandler

    def print_document(self, file_path: str, printer_name: str) -> bool:
        """Apre un documento Word, imposta la stampante e lo stampa."""
        if not Path(file_path).is_file():
            self.logger(f"ERRORE: File Word non trovato: {file_path}", "ERROR")
            return False

        with self.word_handler_class(self.logger) as word:
            if not word:
                return False
            doc = None
            try:
                word.ActivePrinter = printer_name
                doc = word.Documents.Open(file_path)
                doc.PrintOut()
                self.logger(f"Comando di stampa inviato per: {os.path.basename(file_path)}", "SUCCESS")
                return True
            except Exception as e:
                self.logger(f"Errore stampa Word {os.path.basename(file_path)}: {e}", "ERROR")
                return False
            finally:
                if doc:
                    with suppress(Exception):
                        doc.Close(SaveChanges=0)

    def open_document(self, word_app, file_path: str):
        """Apre un documento utilizzando un'istanza di Word esistente."""
        try:
            return word_app.Documents.Open(file_path)
        except Exception as e:
            self.logger(f"Errore apertura documento Word {os.path.basename(file_path)}: {e}", "ERROR")
            return None
