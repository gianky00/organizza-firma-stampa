import traceback
from contextlib import suppress


class EmailHandler:
    """
    Handles the creation of email drafts in Microsoft Outlook.
    """

    def __init__(self, logger):
        self.logger = logger

    def _generate_html_body(self, intro_text, file_list):
        """Genera il contenuto HTML raggruppando i file per TCL in tabelle separate e affiancate."""
        # Pulizia testo introduzione
        intro_html = intro_text.replace("{file_list}", "").replace("Elenco file:", "").strip().replace("\n", "<br>")

        # Raggruppamento file per TCL
        grouped_files: dict[str, list[str]] = {}
        for file_data in file_list:
            tcl = file_data.get("tcl", "N/D")
            if tcl not in grouped_files:
                grouped_files[tcl] = []
            grouped_files[tcl].append(file_data.get("name", "N/D"))

        # 1. Creazione del Riepilogo (Summary)
        summary_html = """
        <div style="margin-bottom: 25px; padding: 15px; background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 5px;">
            <h4 style="color: #333333; margin-top: 0; margin-bottom: 10px; font-size: 11pt; text-transform: uppercase;">Riepilogo TCL</h4>
            <ul style="margin: 0; padding-left: 20px; color: #555555; font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 10pt;">
        """
        for tcl, names in sorted(grouped_files.items()):
            count = len(names)
            noun = "scheda" if count == 1 else "schede"
            summary_html += f"<li style='margin-bottom: 5px;'><strong>{tcl}</strong>: {count} {noun}</li>"
        summary_html += """
            </ul>
        </div>
        """

        # 2. Generazione delle tabelle affiancate (fluide)
        sections_html = '<div>\n'
        
        for tcl, names in sorted(grouped_files.items()):
            table_rows = ""
            for i, filename in enumerate(names):
                bg_color = "#f9f9f9" if i % 2 == 0 else "#ffffff"
                table_rows += f"""
                    <tr style="background-color: {bg_color};">
                        <td style="padding: 6px 12px; border-bottom: 1px solid #eeeeee; color: #333333; font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 9pt; white-space: nowrap;">
                            {filename}
                        </td>
                    </tr>
                """
            
            sections_html += f"""
            <table align="left" style="margin-right: 25px; margin-bottom: 20px; border-collapse: collapse; border: 1px solid #eeeeee;">
                <thead>
                    <tr>
                        <td style="padding: 0 0 8px 0; border: none;">
                            <h4 style="color: #333333; margin: 0; font-size: 12pt; text-transform: uppercase; border-left: 3px solid #0078D4; padding-left: 8px;">{tcl}</h4>
                        </td>
                    </tr>
                    <tr style="background-color: #f0f0f0;">
                        <th style="padding: 6px 12px; border-bottom: 2px solid #dddddd; text-align: left; font-size: 10pt; color: #555555; white-space: nowrap;">NOME FILE</th>
                    </tr>
                </thead>
                <tbody>
                    {table_rows}
                </tbody>
            </table>
            """
            
        sections_html += '<br style="clear:both;">\n</div>\n'

        # Template Finale
        html = f"""
        <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin-bottom: 20px;">
            <div style="padding: 10px 0; background-color: #ffffff;">
                <!-- Introduzione -->
                <div style="color: #444444; font-size: 11pt; line-height: 1.6; margin-bottom: 25px;">
                    {intro_html}
                </div>
                
                {summary_html}

                <h3 style="color: #333333; font-size: 12pt; border-bottom: 2px solid #eeeeee; padding-bottom: 5px; margin-bottom: 20px;">DETTAGLIO SCHEDE ALLEGATE</h3>
                
                {sections_html}
            </div>
        </div>
        <br style="clear:both;">
        <br>
        """
        return html

    def create_outlook_draft(self, draft_info):
        """
        Creates a draft in Outlook with the provided information and attachments.
        """
        try:
            import pythoncom
            import win32com.client
        except ImportError:
            self.logger(
                "ERRORE FATALE: Le librerie necessarie (pywin32) per controllare Outlook non sono installate.", "ERROR"
            )
            return

        pythoncom.CoInitialize()
        outlook = None
        mail = None
        try:
            to = draft_info["to"]
            cc = draft_info.get("cc", "")
            subject = draft_info["subject"]
            intro_text = draft_info["intro_text"]
            file_list = draft_info["file_list"]
            attachments = draft_info["attachments"]

            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = to
            mail.CC = cc
            mail.Subject = subject

            # Genera il corpo HTML professionale
            html_content = self._generate_html_body(intro_text, file_list)

            # Get user signature and prepending content
            mail.Display()
            signature = mail.HTMLBody

            # Uniamo il nostro contenuto HTML con la firma esistente di Outlook
            mail.HTMLBody = html_content + signature

            for file_path in attachments:
                if file_path:
                    mail.Attachments.Add(file_path)

            self.logger("Bozza email creata e mostrata con successo.", "SUCCESS")

        except Exception as e:
            self.logger(
                f"ERRORE FATALE: Impossibile creare la bozza dell'email. Verificare che Outlook sia installato e configurato. Dettagli: {e}",
                "ERROR",
            )
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            if mail:
                del mail
            if outlook:
                del outlook
            with suppress(Exception):
                import pythoncom

                pythoncom.CoUninitialize()
