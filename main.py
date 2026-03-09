from __future__ import annotations

from datetime import datetime
from pathlib import Path

from config.settings import BASE_DIR, DOWNLOAD_FOLDER, MAIL_INPUT_FOLDER, MAIL_SOURCE_TYPE, MAX_PDF
from database.mail_repository import MailRepository
from mail_sources.folder_mail_source import FolderMailSource
from outlook.attachment_handler import AttachmentHandler
from outlook.outlook_client import OutlookClient
from utils.text_utils import normalize_latin, normalize_latin_filename


def build_mail_source():
    if MAIL_SOURCE_TYPE == "folder":
        return FolderMailSource(Path(BASE_DIR) / MAIL_INPUT_FOLDER)
    if MAIL_SOURCE_TYPE == "outlook":
        return OutlookClient()
    raise ValueError(f"Type de source mail inconnu : {MAIL_SOURCE_TYPE}")


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
    if hasattr(message, "entry_id"):
        return (
            message.entry_id,
            message.message_id or message.entry_id,
            normalize_latin(message.subject),
            normalize_latin(message.sender_email_address),
            getattr(message, "received_time", None),
            getattr(message, "store_id", None),
        )

    entry_id = message.EntryID
    message_id = getattr(message, "InternetMessageID", None) or entry_id
    subject = normalize_latin(getattr(message, "Subject", ""))
    sender = normalize_latin(getattr(message, "SenderEmailAddress", ""))
    mail_date = getattr(message, "ReceivedTime", None)
    store_id = getattr(message.Parent, "StoreID", None)

    return entry_id, message_id, subject, sender, mail_date, store_id


def main():
    print("Demarrage du loader PDF")
    print("Source mail :", MAIL_SOURCE_TYPE)
    print("Dossier input :", MAIL_INPUT_FOLDER)

    today_folder = Path(BASE_DIR) / DOWNLOAD_FOLDER / datetime.now().strftime("%Y-%m-%d")

    mail_source = build_mail_source()
    attachment_handler = AttachmentHandler(str(today_folder))
    mail_repo = MailRepository()
    pdf_count = 0

    try:
        messages = mail_source.get_messages_sorted()
        print("Elements trouves :", len(messages))

        for message in messages:
            if pdf_count >= MAX_PDF:
                print(f"Termine : {MAX_PDF} PDF recuperes.")
                break

            entry_id, message_id, subject, sender, mail_date, store_id = read_message_fields(message)
            saved_pdfs = attachment_handler.save_pdf_attachments(message)

            for pdf_name in saved_pdfs:
                safe_pdf_name = normalize_latin_filename(pdf_name)

                mail_repo.upsert_mail_attachment(
                    message_id=message_id,
                    entry_id=entry_id,
                    nom_pdf=safe_pdf_name,
                    sujet=subject,
                    expediteur=sender,
                    date_mail=mail_date,
                    store_id=store_id,
                )

                print(
                    f"PDF traite : {safe_pdf_name} | "
                    f"sujet={subject} | "
                    f"expediteur={sender} | "
                    f"entry_id={entry_id} | "
                    f"date_mail={mail_date}"
                )

                pdf_count += 1

                if pdf_count >= MAX_PDF:
                    break

    except Exception as e:
        print(f"Erreur globale : {e}")
        raise

    finally:
        mail_repo.close()
        print("Connexion SQL fermee")


if __name__ == "__main__":
    main()