from __future__ import annotations

import ctypes

from mail_sources.folder_mail_source import FolderMessage
from utils.text_utils import normalize_latin


def _show_dialog(title: str, message: str, icon: int = 0x40):
    ctypes.windll.user32.MessageBoxW(0, message, title, icon)


def _get_first_message(messages):
    for message in messages:
        return message
    return None


def _get_message_subject(message) -> str:
    if isinstance(message, FolderMessage):
        return normalize_latin(getattr(message, 'subject', '') or '')

    return normalize_latin(getattr(message, 'Subject', '') or '')


def run_debug_first_pdf(mail_source):
    try:
        messages = mail_source.get_messages_sorted()

        try:
            count = len(messages)
        except TypeError:
            count = getattr(messages, 'Count', '?')

        print('Elements trouves :', count)

        first_message = _get_first_message(messages)

        if first_message is None:
            print('Aucun mail trouve.')
            _show_dialog('Debug Outlook', 'Aucun mail trouve dans le dossier', 0x10)
            return

        subject = _get_message_subject(first_message).strip()

        if not subject:
            subject = '(sans sujet)'

        print('Premier mail trouve :', subject)
        _show_dialog('Debug Outlook', subject, 0x40)

    except Exception as e:
        print(f'Erreur debug : {e}')
        _show_dialog('Erreur debug Outlook', str(e), 0x10)