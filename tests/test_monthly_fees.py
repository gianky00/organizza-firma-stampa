from unittest.mock import MagicMock

from src.logic.monthly_fees import MonthlyFeesProcessor


def test_monthly_fees_run_process_simple(fs, mock_gui, app_config):
    # Setup
    g_dir = "/g"
    c_dir = "/c"
    w_file = "/t.docx"
    fs.create_dir(g_dir)
    fs.create_dir(c_dir)
    fs.create_file(w_file)
    g_path = g_dir + "/G.xlsx"
    fs.create_file(g_path)
    c_path = c_dir + "/C.xlsx"
    fs.create_file(c_path)

    # Mock Excel e Word locali
    mock_excel_app = MagicMock()
    mock_word_app = MagicMock()
    mock_wb = MagicMock()
    mock_doc = MagicMock()
    mock_excel_app.Workbooks.Open.return_value = mock_wb
    mock_word_app.Documents.Open.return_value = mock_doc

    # Data
    paths = {"giornaliera": g_path, "word": w_file, "consuntivi": [{"name": "P", "path": c_path, "print": True}]}

    mock_xls_h = MagicMock()
    mock_xls_h.return_value.__enter__.return_value = mock_excel_app
    mock_doc_h = MagicMock()
    mock_doc_h.return_value.__enter__.return_value = mock_word_app

    # Gateway Mocks
    mock_xls_gateway_inst = MagicMock()
    mock_xls_gateway_inst.excel_handler_class = mock_xls_h
    mock_xls_gateway_class = MagicMock(return_value=mock_xls_gateway_inst)

    mock_doc_gateway_inst = MagicMock()
    mock_doc_gateway_inst.word_handler_class = mock_doc_h
    mock_doc_gateway_class = MagicMock(return_value=mock_doc_gateway_inst)

    processor = MonthlyFeesProcessor(
        mock_gui, app_config, excel_gateway_class=mock_xls_gateway_class, word_gateway_class=mock_doc_gateway_class
    )

    # Mocca il sistema delle stampanti in modo che riconosca "Printer"
    processor.get_printers = MagicMock(return_value=(["Printer"], "Printer"))

    cancel_event = MagicMock()
    cancel_event.is_set.return_value = False
    processor.run_printing_process(cancel_event, paths, "Printer", "Macro")

    assert mock_word_app.Documents.Open.called
    assert mock_excel_app.Run.called
