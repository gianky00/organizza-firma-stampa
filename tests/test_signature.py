import os
from unittest.mock import MagicMock, patch

from src.logic.signature import SignatureProcessor


def test_signature_apply_logic_simple(fs, mock_gui, app_config):
    # Setup
    ed = "/ed"
    fs.create_dir(ed)
    fs.create_file(os.path.join(ed, "F.xlsx"))
    pd = "/pd"
    fs.create_dir(pd)
    app_config.firma_pdf_dir.get.return_value = pd
    ip = "/i.png"
    fs.create_file(ip)
    app_config.firma_image_path.get.return_value = ip
    gp = "/g.exe"
    fs.create_file(gp)
    app_config.firma_ghostscript_path.get.return_value = gp
    app_config.firma_excel_dir.get.return_value = ed
    app_config.firma_processing_mode.get.return_value = "schede"

    # Mock Excel locale
    mock_excel_app = MagicMock()
    mock_wb = MagicMock()
    mock_ws = MagicMock()
    mock_excel_app.Workbooks.Open.return_value = mock_wb
    mock_wb.Worksheets.return_value = mock_ws

    # Matching modello
    def cells_side_effect(r, c):
        cell = MagicMock()
        if r == 2 and c == 5:
            cell.Text = "SCHEDAMANUTENZIONE"
        else:
            cell.Text = ""
        cell.Top = 100
        cell.Left = 100
        return cell

    mock_ws.Cells.side_effect = cells_side_effect

    mock_handler_class = MagicMock()
    mock_handler_class.return_value.__enter__.return_value = mock_excel_app

    # Passiamo il mock come excel_handler_class tramite il gateway mockato
    mock_gateway_instance = MagicMock()
    mock_gateway_class = MagicMock(return_value=mock_gateway_instance)

    # Ma nel test semplice, vogliamo che _sign_and_export usi l'handler iniettato
    # SignatureProcessor._sign_and_export cerca self.excel_gateway.excel_handler_class
    mock_gateway_instance.excel_handler_class = mock_handler_class

    processor = SignatureProcessor(
        mock_gui, app_config, MagicMock(), MagicMock(), MagicMock(), excel_gateway_class=mock_gateway_class
    )

    cancel_event = MagicMock()
    cancel_event.is_set.return_value = False
    with patch("src.logic.signature.subprocess.run"):
        processor.run_full_signature_process(cancel_event)

    assert mock_wb.ActiveSheet.ExportAsFixedFormat.called
    assert mock_ws.Shapes.AddPicture.called


def test_signature_preventivi_logic_simple(fs, mock_gui, app_config):
    # Setup
    ed = "/ed"
    fs.create_dir(ed)
    fs.create_file(os.path.join(ed, "F.xlsx"))
    pd = "/pd"
    fs.create_dir(pd)
    app_config.firma_pdf_dir.get.return_value = pd
    ip = "/i.png"
    fs.create_file(ip)
    app_config.firma_image_path.get.return_value = ip
    gp = "/g.exe"
    fs.create_file(gp)
    app_config.firma_ghostscript_path.get.return_value = gp
    app_config.firma_excel_dir.get.return_value = ed
    app_config.firma_processing_mode.get.return_value = "preventivi"

    # Mock Excel locale
    mock_excel_app = MagicMock()
    mock_wb = MagicMock()
    mock_ws = MagicMock()
    mock_ws.Name = "Consuntivo"
    mock_excel_app.Workbooks.Open.return_value = mock_wb
    mock_wb.Worksheets = [mock_ws]

    mock_handler_class = MagicMock()
    mock_handler_class.return_value.__enter__.return_value = mock_excel_app

    mock_gateway_instance = MagicMock()
    mock_gateway_instance.excel_handler_class = mock_handler_class
    mock_gateway_class = MagicMock(return_value=mock_gateway_instance)

    processor = SignatureProcessor(
        mock_gui, app_config, MagicMock(), MagicMock(), MagicMock(), excel_gateway_class=mock_gateway_class
    )

    cancel_event = MagicMock()
    cancel_event.is_set.return_value = False
    with patch("src.logic.signature.subprocess.run"):
        processor.run_full_signature_process(cancel_event)

    assert mock_ws.Activate.called
    assert mock_ws.ExportAsFixedFormat.called
