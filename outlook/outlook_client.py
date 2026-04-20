from __future__ import annotations

from datetime import datetime
import ctypes

import win32com.client

from config.settings import OUTLOOK_FOLDER_PATH, OUTLOOK_MAILBOX, OUTLOOK_FOLDER_BASE


class OutlookClient:
    """
    Client Outlook basé sur win32com (MAPI)
    """

    INBOX_FOLDER_ID = 6  # Outlook constant: Inbox


    def __init__(self, date_min: datetime | None = None):
        self.date_min = self._to_python_datetime(date_min)
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.mailbox_name = OUTLOOK_MAILBOX

        mailbox_root = self._get_mailbox_root(self.mailbox_name)

        if OUTLOOK_FOLDER_BASE == "root":
            start_folder = mailbox_root
            start_label = "Mailbox"
        else:
            start_folder = self._get_mailbox_inbox(self.mailbox_name)
            start_label = "Inbox"

        self.folder = self._get_target_folder(start_folder, OUTLOOK_FOLDER_PATH, start_label)



    @staticmethod
    def _show_dialog(title: str, message: str, icon: int = 0x10):
        try:
            ctypes.windll.user32.MessageBoxW(0, message, title, icon)
        except Exception:
            pass

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

    @staticmethod
    def _clean(value) -> str:
        return str(value or "").strip()

    def _list_available_mailboxes(self) -> list[str]:
        lines: list[str] = []
        seen: set[str] = set()

        accounts = self.outlook.Accounts
        for i in range(1, accounts.Count + 1):
            try:
                account = accounts.Item(i)
                display_name = self._clean(getattr(account, "DisplayName", ""))
                smtp_address = self._clean(getattr(account, "SmtpAddress", ""))
                label = f"- {display_name} <{smtp_address}>" if smtp_address else f"- {display_name}"
                if label and label not in seen:
                    seen.add(label)
                    lines.append(label)
            except Exception:
                continue

        stores = self.outlook.Stores
        for i in range(1, stores.Count + 1):
            try:
                store = stores.Item(i)
                display_name = self._clean(getattr(store, "DisplayName", ""))
                if display_name:
                    label = f"- {display_name}"
                    if label not in seen:
                        seen.add(label)
                        lines.append(label)
            except Exception:
                continue

        return lines or ["- Aucune boite detectee"]

    def _list_folder_tree(self, folder, level: int = 0, max_depth: int = 3) -> list[str]:
        if folder is None or level > max_depth:
            return []

        lines: list[str] = []

        try:
            folders = folder.Folders
            for i in range(1, folders.Count + 1):
                child = folders.Item(i)
                child_name = self._clean(getattr(child, "Name", ""))
                if not child_name:
                    continue

                indent = "  " * level
                lines.append(f"{indent}- {child_name}")

                if level < max_depth:
                    lines.extend(self._list_folder_tree(child, level + 1, max_depth=max_depth))
        except Exception:
            pass

        return lines




    def _get_target_folder(self, start_folder, folder_path: list[str], start_label: str):
        folder = start_folder
        resolved_parts: list[str] = []

        for folder_name in folder_path:
            try:
                folder = folder.Folders[folder_name]
                resolved_parts.append(folder_name)
            except Exception:
                asked_path = " / ".join(folder_path) if folder_path else f"({start_label})"
                resolved_path = start_label
                if resolved_parts:
                    resolved_path += " / " + " / ".join(resolved_parts)

                available_folders = "\n".join(self._list_folder_tree(folder, max_depth=5))
                if not available_folders:
                    available_folders = "- Aucun sous-dossier detecte"

                message = (
                    f"Boite Outlook trouvee : {self.mailbox_name}\n"
                    f"Base de recherche : {start_label}\n"
                    f"Chemin demande : {start_label} / {asked_path}\n"
                    f"Chemin resolu : {resolved_path}\n\n"
                    f"Sous-dossiers disponibles depuis '{resolved_path}' :\n"
                    f"{available_folders}"
                )
                self._show_dialog("Dossier Outlook introuvable", message, 0x10)
                raise ValueError(f"Dossier Outlook introuvable : {start_label} / {asked_path}")

        return folder




    def _get_mailbox_inbox(self, mailbox_name: str):
        target = self._norm(mailbox_name)

        if not target:
            message = "Le setting 'outlook_mailbox' est vide."
            self._show_dialog("Boite Outlook introuvable", message, 0x10)
            raise ValueError(message)

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

        available_mailboxes = "\n".join(self._list_available_mailboxes())
        message = (
            f"Boite Outlook introuvable : {mailbox_name}\n\n"
            f"Boites disponibles :\n{available_mailboxes}"
        )
        self._show_dialog("Boite Outlook introuvable", message, 0x10)
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


    def _get_mailbox_root(self, mailbox_name: str):
        target = self._norm(mailbox_name)

        if not target:
            message = "Le setting 'outlook.mailbox' est vide."
            self._show_dialog("Boite Outlook introuvable", message, 0x10)
            raise ValueError(message)

        accounts = self.outlook.Accounts
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)

            smtp_address = self._norm(getattr(account, "SmtpAddress", ""))
            display_name = self._norm(getattr(account, "DisplayName", ""))

            if smtp_address == target or display_name == target:
                return account.DeliveryStore.GetRootFolder()

        stores = self.outlook.Stores
        for i in range(1, stores.Count + 1):
            store = stores.Item(i)
            display_name = self._norm(getattr(store, "DisplayName", ""))

            if display_name == target or target in display_name:
                return store.GetRootFolder()

        available_mailboxes = "\n".join(self._list_available_mailboxes())
        message = (
            f"Boite Outlook introuvable : {mailbox_name}\n\n"
            f"Boites disponibles :\n{available_mailboxes}"
        )
        self._show_dialog("Boite Outlook introuvable", message, 0x10)
        raise ValueError(f"Boite Outlook introuvable : {mailbox_name}")
    

    def _get_mailbox_root(self, mailbox_name: str):
        target = self._norm(mailbox_name)

        if not target:
            message = "Le setting 'outlook.mailbox' est vide."
            self._show_dialog("Boite Outlook introuvable", message, 0x10)
            raise ValueError(message)

        accounts = self.outlook.Accounts
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)

            smtp_address = self._norm(getattr(account, "SmtpAddress", ""))
            display_name = self._norm(getattr(account, "DisplayName", ""))

            if smtp_address == target or display_name == target:
                return account.DeliveryStore.GetRootFolder()

        stores = self.outlook.Stores
        for i in range(1, stores.Count + 1):
            store = stores.Item(i)
            display_name = self._norm(getattr(store, "DisplayName", ""))

            if display_name == target or target in display_name:
                return store.GetRootFolder()

        available_mailboxes = "\n".join(self._list_available_mailboxes())
        message = (
            f"Boite Outlook introuvable : {mailbox_name}\n\n"
            f"Boites disponibles :\n{available_mailboxes}"
        )
        self._show_dialog("Boite Outlook introuvable", message, 0x10)
        raise ValueError(f"Boite Outlook introuvable : {mailbox_name}")