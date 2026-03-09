import win32com.client
from config.settings import OUTLOOK_FOLDER_PATH


class OutlookClient:
    """
    Client Outlook basé sur win32com (MAPI)
    """

    INBOX_FOLDER_ID = 6  # Outlook constant: Inbox

    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        folder = self.outlook.GetDefaultFolder(self.INBOX_FOLDER_ID)

        for folder_name in OUTLOOK_FOLDER_PATH:
            folder = folder.Folders[folder_name]

        self.folder = folder

    def get_messages_sorted(self):
        """
        Retourne la collection des messages triés par date de réception décroissante
        """
        messages = self.folder.Items
        messages.Sort("[ReceivedTime]", True)
        return messages