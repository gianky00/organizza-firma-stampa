import os
import pytest
from unittest.mock import MagicMock, patch, PropertyMock
from src.logic.signature import SignatureProcessor

def test_signature_apply_logic_refactored(fs, mock_gui, app_config, mock_excel):
    # Setup
    app_config.firma_processing_mode.get.return_value = "schede"
    ed = "/ed"; fs.create_dir(ed); fs.create_file(os.path.join(ed, "F.xlsx"))
    pd = "/pd"; fs.create_dir(pd); app_config.firma_pdf_dir.get.return_value = pd
    ip = "/i.png"; fs.create_file(ip); app_config.firma_image_path.get.return_value = ip
    gp = "/g.exe"; fs.create_file(gp); app_config.firma_ghostscript_path.get.return_value = gp
    
    excel_app, wb, ws = mock_excel
    # Mocking Worksheet.Cells access
    def cells_side_effect(r, c):
        cell = MagicMock()
        if r == 2 and c == 5: cell.Text = "SCHEDAMANUTENZIONE"
        else: cell.Text = ""
        cell.Top = 100; cell.Left = 100
        return cell
    ws.Cells.side_effect = cells_side_effect
    
    # Prepariamo la classe mock dell'handler
    mock_handler_class = MagicMock()
    mock_handler_instance = mock_handler_class.return_value
    mock_handler_instance.__enter__.return_value = excel_app
    
    # Iniezione della dipendenza!
    processor = SignatureProcessor(
        mock_gui, app_config, MagicMock(), MagicMock(), MagicMock(),
        excel_handler_class=mock_handler_class
    )
    
    with patch("src.logic.signature.subprocess.run"):
        processor.run_full_signature_process(MagicMock())
        
    # Ora le asserzioni sono garantite perché usiamo l'iniezione
    assert excel_app.Workbooks.Open.called
    assert ws.Shapes.AddPicture.called
    assert wb.ActiveSheet.ExportAsFixedFormat.called
