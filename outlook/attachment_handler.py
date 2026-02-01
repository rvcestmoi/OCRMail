# outlook/attachment_handler.py

import os


class AttachmentHandler:
    """
    Gestionnaire des pièces jointes (filtrage, sauvegarde)
    """

    def __init__(self, download_folder: str):
        self.download_folder = download_folder
        os.makedirs(self.download_folder, exist_ok=True)

    @staticmethod
    def is_pdf(filename: str) -> bool:
        return filename.lower().endswith(".pdf")

    def save_pdf_attachments(self, message):
        """
        Sauvegarde toutes les PJ PDF d'un message.
        Retourne une liste des fichiers sauvegardés (noms).
        """
        saved_files = []

        if message.Attachments.Count <= 0:
            return saved_files

        for i in range(1, message.Attachments.Count + 1):
            attachment = message.Attachments.Item(i)
            filename = attachment.FileName

            if self.is_pdf(filename):
                filepath = os.path.join(self.download_folder, filename)
                attachment.SaveAsFile(filepath)
                saved_files.append(filename)

        return saved_files
