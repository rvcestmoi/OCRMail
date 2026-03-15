from __future__ import annotations

from datetime import datetime

import win32com.client

from config.settings import OUTLOOK_FOLDER_PATH, OUTLOOK_MAILBOX


class OutlookClient:
    """
    Client Outlook basé sur win32com (MAPI)
    """

    INBOX_FOLDER_ID = 6  # Outlook constant: Inbox

    def __init__(self, date_min: datetime | None = None):
        self.date_min = self._to_python_datetime(date_min)
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.mailbox_name = OUTLOOK_MAILBOX

        folder = self._get_mailbox_inbox(self.mailbox_name)

        for folder_name in OUTLOOK_FOLDER_PATH:
            folder = folder.Folders[folder_name]

        self.folder = folder

    @staticmethod
    def _to_python_datetime(value):
        if value is None:
            return None

        if not isinstance(value, datetime):
            value = datetime(
                value.year,
                value.month,
                value.day,
                value.hour,
                value.minute,
                value.second,
            )

        if value.tzinfo is not None:
            return value.replace(tzinfo=None)

        return value

    @staticmethod
    def _norm(value) -> str:
        return str(value or "").strip().lower()

    def _get_mailbox_inbox(self, mailbox_name: str):
        target = self._norm(mailbox_name)

        if not target:
            raise ValueError("Le setting 'outlook_mailbox' est vide.")

        accounts = self.outlook.Accounts
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)

            smtp_address = self._norm(getattr(account, "SmtpAddress", ""))
            display_name = self._norm(getattr(account, "DisplayName", ""))

            if smtp_address == target or display_name == target:
                delivery_store = account.DeliveryStore
                return delivery_store.GetDefaultFolder(self.INBOX_FOLDER_ID)

        stores = self.outlook.Stores
        for i in range(1, stores.Count + 1):
            store = stores.Item(i)
            display_name = self._norm(getattr(store, "DisplayName", ""))

            if display_name == target or target in display_name:
                return store.GetDefaultFolder(self.INBOX_FOLDER_ID)

        raise ValueError(f"Boite Outlook introuvable : {mailbox_name}")

    def get_messages_sorted(self):
        """
        Retourne la liste des messages triés du plus ancien au plus récent.
        Si date_min est renseignée, seuls les messages reçus à partir de cette date sont conservés.
        """
        items = self.folder.Items
        items.Sort("[ReceivedTime]", False)

        messages = []
        item = items.GetFirst()

        while item:
            received_dt = self._to_python_datetime(getattr(item, "ReceivedTime", None))

            if self.date_min is not None and received_dt is not None and received_dt < self.date_min:
                item = items.GetNext()
                continue

            messages.append(item)
            item = items.GetNext()

        return messages
