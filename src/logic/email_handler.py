import win32com.client
import traceback

class EmailHandler:
    """
    Handles the creation of email drafts in Microsoft Outlook.
    """
    def __init__(self, logger):
        self.logger = logger

    def create_outlook_draft(self, to, subject, body, attachments):
        """
        Creates and displays an Outlook email draft with the specified details.

        Args:
            to (str): The recipient's email address(es).
            subject (str): The email's subject line.
            body (str): The plain text body of the email.
            attachments (list): A list of full file paths to attach.
        """
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)

            mail.To = to
            mail.Subject = subject
            mail.Body = body

            if not attachments:
                self.logger("Nessun allegato da aggiungere.", "WARNING")
            else:
                self.logger(f"Aggiunta di {len(attachments)} allegati...", "INFO")
                for attachment_path in attachments:
                    try:
                        mail.Attachments.Add(attachment_path)
                    except Exception as e:
                        self.logger(f"Impossibile aggiungere l'allegato: {attachment_path}. Errore: {e}", "ERROR")

            # Display the draft in an Outlook window
            mail.Display(True) # True makes it modal
            self.logger("Bozza email creata e mostrata con successo.", "SUCCESS")

        except Exception as e:
            self.logger(f"ERRORE FATALE: Impossibile creare la bozza dell'email. Verificare che Outlook sia installato e configurato. Dettagli: {e}", "ERROR")
            self.logger(traceback.format_exc(), "ERROR")
