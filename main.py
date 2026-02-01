# main.py

import os
from datetime import datetime

from outlook.outlook_client import OutlookClient
from outlook.attachment_handler import AttachmentHandler
from database.mail_repository import MailRepository
from config.settings import DOWNLOAD_FOLDER, MAX_PDF
from utils.text_utils import normalize_latin


def main():
    print("🚀 Démarrage du loader Outlook PDF")

    # Dossier du jour (ex: data/PJ/2025-06-11)
    today_folder = os.path.join(
        DOWNLOAD_FOLDER,
        datetime.now().strftime("%Y-%m-%d")
    )

    outlook_client = OutlookClient()
    attachment_handler = AttachmentHandler(today_folder)
    mail_repo = MailRepository()

    pdf_count = 0

    try:
        messages = outlook_client.get_messages_sorted()

        for message in messages:

            if pdf_count >= MAX_PDF:
                print(f"\n🎯 Terminé : {MAX_PDF} PDF récupérés.")
                break

            # --- Identifiants Outlook ---
            entry_id = message.EntryID
            message_id = getattr(message, "InternetMessageID", None)
            if not message_id:
                message_id = entry_id  # fallback garanti non-null

            # --- Métadonnées normalisées ---
            subject = normalize_latin(getattr(message, "Subject", ""))
            sender = normalize_latin(getattr(message, "SenderEmailAddress", ""))

            # --- Sauvegarde des PDF ---
            saved_pdfs = attachment_handler.save_pdf_attachments(message)

            for pdf_name in saved_pdfs:

                safe_pdf_name = normalize_latin(pdf_name)

                mail_repo.upsert_mail_attachment(
                    message_id=message_id,
                    entry_id=entry_id,
                    nom_pdf=safe_pdf_name,
                    sujet=subject,
                    expediteur=sender
                )

                print("✅ PDF traité :", safe_pdf_name)
                print("   📩 Sujet :", subject)
                print("   👤 Expéditeur :", sender)
                print("   🆔 EntryID :", entry_id)
                print("-" * 50)

                pdf_count += 1

                if pdf_count >= MAX_PDF:
                    break

    except Exception as e:
        print("❌ Erreur globale :", e)

    finally:
        mail_repo.close()
        print("🔒 Connexion SQL fermée")


if __name__ == "__main__":
    main()
