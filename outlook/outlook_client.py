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

        # 1) Recherche par comptes Outlook (le plus fiable si la boite est un vrai compte)
        accounts = self.outlook.Accounts
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)

            smtp_address = self._norm(getattr(account, "SmtpAddress", ""))
            display_name = self._norm(getattr(account, "DisplayName", ""))

            if smtp_address == target or display_name == target:
                delivery_store = account.DeliveryStore
                return delivery_store.GetDefaultFolder(self.INBOX_FOLDER_ID)

        # 2) Recherche par stores (utile pour boite partagée / boite supplémentaire)
        stores = self.outlook.Stores
        for i in range(1, stores.Count + 1):
            store = stores.Item(i)
            display_name = self._norm(getattr(store, "DisplayName", ""))

            if display_name == target or target in display_name:
                return store.GetDefaultFolder(self.INBOX_FOLDER_ID)

        available_accounts = []
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)
            smtp_address = getattr(account, "SmtpAddress", "")
            display_name = getattr(account, "DisplayName", "")
            available_accounts.append(f"{display_name} / {smtp_address}")

        available_stores = []
        for i in range(1, stores.Count + 1):
            store = stores.Item(i)
            available_stores.append(str(getattr(store, "DisplayName", "")))

        raise ValueError(
            "Boite Outlook introuvable pour 'outlook_mailbox'="
            f"{mailbox_name}. Comptes disponibles: {available_accounts}. "
            f"Stores disponibles: {available_stores}"
        )

    def get_messages_sorted(self):
        """
        Retourne la liste des messages triés par date de réception décroissante.
        Si date_min est renseignée, seuls les messages reçus à partir de cette date sont conservés.
        """
        items = self.folder.Items
        items.Sort("[ReceivedTime]", True)

        messages = []
        item = items.GetFirst()

        while item:
            received_dt = self._to_python_datetime(getattr(item, "ReceivedTime", None))

            if self.date_min is not None and received_dt is not None and received_dt < self.date_min:
                break

            messages.append(item)
            item = items.GetNext()

        return messages