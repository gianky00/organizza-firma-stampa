import traceback
from contextlib import suppress


class EmailHandler:
    """
    Handles the creation of email drafts in Microsoft Outlook.
    """

    def __init__(self, logger):
        self.logger = logger

    def _generate_html_body(self, intro_text, file_list):
        """Genera il contenuto HTML raggruppando i file per TCL in tabelle separate."""
        # Pulizia testo introduzione
        intro_html = intro_text.replace("{file_list}", "").replace("Elenco file:", "").strip().replace("\n", "<br>")

        # Raggruppamento file per TCL
        grouped_files: dict[str, list[str]] = {}
        for file_data in file_list:
            tcl = file_data.get("tcl", "N/D")
            if tcl not in grouped_files:
                grouped_files[tcl] = []
            grouped_files[tcl].append(file_data.get("name", "N/D"))

        # Generazione delle sezioni (una tabella per ogni TCL)
        sections_html = ""
        for tcl, names in grouped_files.items():
            table_rows = ""
            for i, filename in enumerate(names):
                bg_color = "#f9f9f9" if i % 2 == 0 else "#ffffff"
                table_rows += f"""
                    <tr style="background-color: {bg_color};">
                        <td style="padding: 6px 12px; border-bottom: 1px solid #eeeeee; color: #333333; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 10pt; white-space: nowrap; border-right: 1px solid #eeeeee;">
                            {filename}
                        </td>
                        <td style="padding: 6px 12px; border-bottom: 1px solid #eeeeee; color: #666666; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 9.5pt; text-align: center;">
                            {tcl}
                        </td>
                    </tr>
                """

            sections_html += f"""
                <div style="margin-bottom: 25px;">
                    <h4 style="color: #333333; margin-bottom: 8px; font-size: 10pt; text-transform: uppercase; letter-spacing: 1px; border-left: 3px solid #eeeeee; padding-left: 8px; display: block;">REFERENTE: {tcl}</h4>
                    <table style="width: auto; border-collapse: collapse; border: 1px solid #eeeeee;">
                        <thead>
                            <tr style="background-color: #f0f0f0;">
                                <th style="padding: 6px 12px; border-bottom: 2px solid #dddddd; text-align: left; font-size: 9pt; color: #555555; border-right: 1px solid #eeeeee;">NOME FILE</th>
                                <th style="padding: 6px 12px; border-bottom: 2px solid #dddddd; text-align: center; font-size: 9pt; color: #555555;">TCL</th>
                            </tr>
                        </thead>
                        <tbody>
                            {table_rows}
                        </tbody>
                    </table>
                </div>
            """

        # Template Finale
        html = f"""
        <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin-bottom: 20px;">
            <div style="padding: 10px 0; background-color: #ffffff;">
                <!-- Introduzione -->
                <div style="color: #444444; font-size: 11pt; line-height: 1.6; margin-bottom: 25px;">
                    {intro_html}
                </div>

                <h3 style="color: #333333; font-size: 11pt; border-bottom: 2px solid #eeeeee; padding-bottom: 5px; margin-bottom: 20px;">SCHEDE ALLEGATE</h3>
                
                {sections_html}
            </div>
        </div>
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
