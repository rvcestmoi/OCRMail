from __future__ import annotations

import ctypes
from pathlib import Path

from mail_sources.folder_mail_source import FolderMessage
from utils.text_utils import normalize_latin_filename


def _show_dialog(pdf_name: str):
    # titre = nom du PDF, contenu = nom du PDF
    ctypes.windll.user32.MessageBoxW(0, pdf_name, pdf_name, 0x40)


def _get_first_message(messages):
    for message in messages:
        return message
    return None


def _get_pdf_names(message):
    """
    Compatible avec :
    - FolderMailSource
    - Outlook win32com
    """

    # Mode dossier local
    if isinstance(message, FolderMessage):
        return [
            normalize_latin_filename(Path(p).name)
            for p in message.attachments
            if str(p).lower().endswith(".pdf")
        ]

    # Mode Outlook
    pdf_names = []
    attachments = getattr(message, "Attachments", None)

    if attachments is None or attachments.Count <= 0:
        return pdf_names

    for i in range(1, attachments.Count + 1):
        attachment = attachments.Item(i)
        filename = normalize_latin_filename(str(attachment.FileName))

        if filename.lower().endswith(".pdf"):
            pdf_names.append(filename)

    return pdf_names


def run_debug_first_pdf(mail_source):
    messages = mail_source.get_messages_sorted()

    try:
        count = len(messages)
    except TypeError:
        count = getattr(messages, "Count", "?")

    print("Elements trouves :", count)

    first_message = _get_first_message(messages)

    if first_message is None:
        print("Aucun mail trouve.")
        _show_dialog("Aucun mail trouve")
        return

    pdf_names = _get_pdf_names(first_message)

    if not pdf_names:
        print("Le premier mail ne contient aucun PDF.")
        _show_dialog("Aucun PDF sur le premier mail")
        return

    first_pdf_name = pdf_names[0]

    print("Premier PDF trouve :", first_pdf_name)
    _show_dialog(first_pdf_name)