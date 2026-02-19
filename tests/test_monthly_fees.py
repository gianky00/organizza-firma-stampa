import os
import pytest
from unittest.mock import MagicMock, patch, call
from src.logic.monthly_fees import MonthlyFeesProcessor

def test_monthly_fees_word_filling(fs, mock_gui, app_config, mock_logger):
    # Setup
    giornaliera_dir = "/giornaliere"
    cons_dir = "/consuntivi"
    word_template = "/template.docx"
    fs.create_dir(giornaliera_dir)
    fs.create_dir(cons_dir)
    fs.create_file(word_template)
    
    # Creiamo file fittizi
    giornaliera_path = os.path.join(giornaliera_dir, "Giornaliera.xlsx")
    fs.create_file(giornaliera_path)
    cons_path = os.path.join(cons_dir, "CANONE_GENNAIO_TCL_PREZZAVENTO.xlsx")
    fs.create_file(cons_path)
    
    app_config.canoni_giornaliera_path.get.return_value = giornaliera_dir
    app_config.canoni_word_path.get.return_value = word_template
    
    # Processore
    processor = MonthlyFeesProcessor(mock_gui, app_config)
    
    cancel_event = MagicMock()
    cancel_event.is_set.return_value = False
    
    # Argomenti per run_printing_process
    paths_to_print = {
        "giornaliera": giornaliera_path,
        "word": word_template,
        "consuntivi": [
            {"name": "Prezzavento", "path": cons_path, "print": True}
        ]
    }
    printer_name = "FakePrinter"
    macro_name = "FakeMacro"
    
    # Mock Excel e Word
    mock_excel = MagicMock()
    mock_word = MagicMock()
    
    mock_wb_giorn = MagicMock()
    mock_wb_cons = MagicMock()
    mock_doc = MagicMock()
    
    def workbooks_open_side_effect(path, *args, **kwargs):
        if path == giornaliera_path:
            return mock_wb_giorn
        elif path == cons_path:
            return mock_wb_cons
        return MagicMock()
        
    mock_excel.Workbooks.Open.side_effect = workbooks_open_side_effect
    mock_word.Documents.Open.return_value = mock_doc
    
    with (
        patch("src.logic.monthly_fees.ExcelHandler") as mock_xls_handler,
        patch("src.logic.monthly_fees.WordHandler") as mock_doc_handler
    ):
        mock_xls_handler.return_value.__enter__.return_value = mock_excel
        mock_doc_handler.return_value.__enter__.return_value = mock_word
        
        processor.run_printing_process(cancel_event, paths_to_print, printer_name, macro_name)
        
        # Verifica che abbia tentato di aprire Word
        mock_word.Documents.Open.assert_called_with(word_template)
        # Verifica che abbia tentato di eseguire la macro
        mock_excel.Run.assert_called()
