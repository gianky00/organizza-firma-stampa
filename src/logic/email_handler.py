import win32com.client
import traceback

class EmailHandler:
    """
    Handles the creation of email drafts in Microsoft Outlook.
    """
    def __init__(self, logger):
        self.logger = logger

    def create_outlook_draft(self, to, subject, intro_text, file_list, attachments):
        """
        Creates and displays an Outlook email draft, preserving the default signature.

        Args:
            to (str): The recipient's email address(es).
            subject (str): The email's subject line.
            intro_text (str): The initial part of the email body (greeting, etc.).
            file_list (list): A list of filenames (without extension) to include in the body.
            attachments (list): A list of full file paths to attach.
        """
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)

            mail.To = to
            mail.Subject = subject

            # This is the key part: Display the item first to resolve the signature.
            mail.Display()

            # Get the signature, which is now part of the HTMLBody
            signature = mail.HTMLBody

            # Create the file list as a simple HTML list
            files_html = "<br>".join(file_list)

            # Construct the full HTML body
            # We replace the placeholder and prepend it to the signature
            body_html = intro_text.replace("{file_list}", files_html)

            # Convert to HTML paragraphs
            body_with_br = body_html.replace('\n', '<br>')

            mail.HTMLBody = f"<p style='font-family:calibri; font-size:11pt'>{body_with_br}</p>" + signature

            self.logger(f"Aggiunta di {len(attachments)} allegati...", "INFO")
            for attachment_path in attachments:
                try:
                    mail.Attachments.Add(attachment_path)
                except Exception as e:
                    self.logger(f"Impossibile aggiungere l'allegato: {attachment_path}. Errore: {e}", "ERROR")

            self.logger("Bozza email creata e mostrata con successo.", "SUCCESS")

        except Exception as e:
            self.logger(f"ERRORE FATALE: Impossibile creare la bozza dell'email. Verificare che Outlook sia installato e configurato. Dettagli: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
