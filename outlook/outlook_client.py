# outlook/outlook_client.py

import win32com.client


class OutlookClient:
    """
    Client Outlook basé sur win32com (MAPI)
    """

    INBOX_FOLDER_ID = 6  # Outlook constant: Inbox

    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(self.INBOX_FOLDER_ID)

    def get_messages_sorted(self):
        """
        Retourne la collection des messages triés par date de réception décroissante
        """
        messages = self.inbox.Items
        messages.Sort("[ReceivedTime]", True)
        return messages
