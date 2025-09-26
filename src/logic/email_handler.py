import win32com.client
import traceback
import pythoncom

class EmailHandler:
    """
    Handles the creation of email drafts in Microsoft Outlook.
    """
    def __init__(self, logger):
        self.logger = logger

    def create_outlook_draft(self, draft_info):
        """
        Creates and displays a single Outlook email draft from a draft info object.

        Args:
            draft_info (dict): A dictionary containing 'to', 'subject', 'intro_text',
                               'file_list', and 'attachments'.
        """
        pythoncom.CoInitialize()
        try:
            to = draft_info['to']
            cc = draft_info.get('cc', '')
            subject = draft_info['subject']
            intro_text = draft_info['intro_text']
            file_list = draft_info['file_list']
            attachments = draft_info['attachments']

            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)

            mail.To = to
            mail.CC = cc
            mail.Subject = subject

            mail.Display()
            signature = mail.HTMLBody

            files_html = "<br>".join(file_list)
            body_with_files = intro_text.replace("{file_list}", files_html)
            body_with_br = body_with_files.replace('\n', '<br>')

            mail.HTMLBody = f"<p style='font-family:calibri; font-size:11pt'>{body_with_br}</p>" + signature

            self.logger(f"Aggiunta di {len(attachments)} allegati alla bozza '{subject}'...", "INFO")
            for attachment_path in attachments:
                try:
                    mail.Attachments.Add(attachment_path)
                except Exception as e:
                    self.logger(f"Impossibile aggiungere l'allegato: {attachment_path}. Errore: {e}", "ERROR")

            self.logger("Bozza email creata e mostrata con successo.", "SUCCESS")

        except Exception as e:
            self.logger(f"ERRORE FATALE: Impossibile creare la bozza dell'email. Verificare che Outlook sia installato e configurato. Dettagli: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
        finally:
            pythoncom.CoUninitialize()
