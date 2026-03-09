# outlook/attachment_handler.py

import os
from utils.text_utils import normalize_latin_filename


class AttachmentHandler:
    def __init__(self, download_folder: str):
        self.download_folder = download_folder
        os.makedirs(self.download_folder, exist_ok=True)

    @staticmethod
    def is_pdf(filename: str) -> bool:
        return filename.lower().endswith(".pdf")

    def save_pdf_attachments(self, message):
        saved_files = []

        if message.Attachments.Count <= 0:
            return saved_files

        for i in range(1, message.Attachments.Count + 1):
            attachment = message.Attachments.Item(i)

            original_filename = attachment.FileName
            safe_filename = normalize_latin_filename(original_filename)

            if self.is_pdf(safe_filename):
                filepath = os.path.join(self.download_folder, safe_filename)
                attachment.SaveAsFile(filepath)
                saved_files.append(safe_filename)

        return saved_files