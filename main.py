from __future__ import annotations

from pathlib import Path

from config.settings import (
    ALLOWED_EXTENSIONS,
    BASE_DIR,
    DEBUG_FIRST_PDF,
    DOWNLOAD_FOLDER,
    MAIL_DATE_MIN,
    MAIL_INPUT_FOLDER,
    MAIL_SOURCE_TYPE,
    MAX_FILES_TO_FETCH,
)
from config.settings_loader import save_setting
from database.mail_repository import MailRepository
from debug_first_pdf import run_debug_first_pdf
from mail_sources.folder_mail_source import FolderMailSource
from outlook.attachment_handler import AttachmentHandler
from outlook.outlook_client import OutlookClient
from utils.text_utils import normalize_latin, normalize_latin_filename
import os
import sys
import json


SETTING_KEY_MAIL_DATE_MIN = 'mail_date_min'


def build_mail_source():
    if MAIL_SOURCE_TYPE == 'folder':
        return FolderMailSource(Path(BASE_DIR) / MAIL_INPUT_FOLDER, date_min=MAIL_DATE_MIN)
    if MAIL_SOURCE_TYPE == 'outlook':
        return OutlookClient(date_min=MAIL_DATE_MIN)
    raise ValueError(f'Type de source mail inconnu : {MAIL_SOURCE_TYPE}')


def normalize_mail_datetime(value):
    if value is None:
        return None

    if value.tzinfo is not None:
        return value.replace(tzinfo=None)

    return value


def read_message_fields(message):
    """
    Retourne :
    - entry_id
    - message_id
    - subject
    - sender
    - mail_date
    - store_id
    """
    if hasattr(message, 'entry_id'):
        return (
            message.entry_id,
            message.message_id or message.entry_id,
            normalize_latin(message.subject),
            normalize_latin(message.sender_email_address),
            normalize_mail_datetime(getattr(message, 'received_time', None)),
            getattr(message, 'store_id', None),
        )

    entry_id = message.EntryID
    message_id = getattr(message, 'InternetMessageID', None) or entry_id
    subject = normalize_latin(getattr(message, 'Subject', ''))
    sender = normalize_latin(getattr(message, 'SenderEmailAddress', ''))
    mail_date = normalize_mail_datetime(getattr(message, 'ReceivedTime', None))
    store_id = getattr(message.Parent, 'StoreID', None)

    return entry_id, message_id, subject, sender, mail_date, store_id



def format_date_min_for_log():
    if MAIL_DATE_MIN is None:
        return 'aucun filtre'
    return MAIL_DATE_MIN.strftime('%Y-%m-%d %H:%M:%S')



def format_allowed_extensions_for_log():
    return ', '.join(ALLOWED_EXTENSIONS)



def main():
    print('Demarrage du loader PJ')
    print('Source mail :', MAIL_SOURCE_TYPE)
    print('Dossier input :', MAIL_INPUT_FOLDER)
    print('Date minimale :', format_date_min_for_log())
    print('Extensions autorisees :', format_allowed_extensions_for_log())
    from config.settings_loader import SETTINGS_FILE
    mail_repo = None
    attachment_limit_reached = False
    last_cursor_mail_date = MAIL_DATE_MIN
    cursor_can_be_saved = False

    try:
        mail_source = build_mail_source()

        if DEBUG_FIRST_PDF:
            run_debug_first_pdf(mail_source)
            return

        today_folder = Path(DOWNLOAD_FOLDER)
        attachment_handler = AttachmentHandler(str(today_folder), allowed_extensions=ALLOWED_EXTENSIONS)
        mail_repo = MailRepository()
        attachment_count = 0

        messages = mail_source.get_messages_sorted()
        for message in messages:
            if attachment_count >= MAX_FILES_TO_FETCH:
                attachment_limit_reached = True
                print(f'Termine : {MAX_FILES_TO_FETCH} piece(s) jointe(s) recuperee(s).')
                break

            entry_id, message_id, subject, sender, mail_date, store_id = read_message_fields(message)
            if mail_date is not None:
                last_cursor_mail_date = mail_date
                cursor_can_be_saved = True

            if mail_repo.exists_entry_id(entry_id):
                print(
                    f'Mail ignore (entry_id deja present) : '
                    f'sujet={subject} | entry_id={entry_id} | date_mail={mail_date}'
                )
                continue

            saved_attachments = attachment_handler.save_allowed_attachments(message)

            for attachment_name in saved_attachments:
                safe_attachment_name = normalize_latin_filename(attachment_name)

                mail_repo.upsert_mail_attachment(
                    message_id=message_id,
                    entry_id=entry_id,
                    nom_pdf=safe_attachment_name,
                    sujet=subject,
                    expediteur=sender,
                    date_mail=mail_date,
                    store_id=store_id,
                )

                print(
                    f'PJ traitee : {safe_attachment_name} | '
                    f'sujet={subject} | '
                    f'expediteur={sender} | '
                    f'entry_id={entry_id} | '
                    f'date_mail={mail_date}'
                )

                attachment_count += 1
                if attachment_count >= MAX_FILES_TO_FETCH:
                    attachment_limit_reached = True
                    break

            if attachment_limit_reached:
                print(f'Termine : {MAX_FILES_TO_FETCH} piece(s) jointe(s) recuperee(s).')
                break

        if not attachment_limit_reached and cursor_can_be_saved and last_cursor_mail_date is not None:
            save_setting(SETTING_KEY_MAIL_DATE_MIN, last_cursor_mail_date)
            print(
                'Setting JSON mis a jour : '
                f"{SETTING_KEY_MAIL_DATE_MIN}={last_cursor_mail_date.strftime('%Y-%m-%d %H:%M:%S')}"
            )
        else:
            print('Setting JSON inchange : aucun message date a parcourir.')

    except Exception as e:
        print(f'Erreur globale : {e}')
        raise

    finally:
        if mail_repo is not None:
            mail_repo.close()
            print('Connexion SQL fermee')



def get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_settings_path() -> str:
    return os.path.join(get_app_dir(), "settings", "settings.json")


def load_settings() -> dict:
    settings_path = get_settings_path()

    if not os.path.exists(settings_path):
        raise FileNotFoundError(f"Fichier settings introuvable : {settings_path}")

    with open(settings_path, "r", encoding="utf-8") as f:
        return json.load(f)


if __name__ == '__main__':
    main()
