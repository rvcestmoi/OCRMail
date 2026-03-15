from __future__ import annotations

import os
import shutil
from pathlib import Path

from mail_sources.folder_mail_source import FolderMessage
from utils.text_utils import normalize_latin_filename
import time


class AttachmentHandler:
    PR_ATTACHMENT_HIDDEN = 'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B'

    def __init__(self, download_folder: str, allowed_extensions: list[str] | None = None):
        self.download_folder = download_folder
        self.allowed_extensions = {ext.lower() for ext in (allowed_extensions or ['.pdf'])}
        os.makedirs(self.download_folder, exist_ok=True)

    def is_allowed_file(self, filename: str) -> bool:
        return Path(filename).suffix.lower() in self.allowed_extensions

    def _build_destination_path(self, safe_filename: str) -> Path:
        destination = Path(self.download_folder) / safe_filename
        if not destination.exists():
            return destination

        stem = destination.stem
        suffix = destination.suffix
        index = 1
        while True:
            candidate = destination.with_name(f'{stem}_{index}{suffix}')
            if not candidate.exists():
                return candidate
            index += 1

    def _is_hidden_outlook_attachment(self, attachment) -> bool:
        try:
            accessor = getattr(attachment, 'PropertyAccessor', None)
            if accessor is None:
                return False
            return bool(accessor.GetProperty(self.PR_ATTACHMENT_HIDDEN))
        except Exception:
            return False

    def _save_folder_message_attachments(self, message: FolderMessage):
        saved_files = []

        for attachment_path in message.attachments:
            original_filename = Path(attachment_path).name
            safe_filename = self._build_unique_filename(original_filename)

            if not self.is_allowed_file(safe_filename):
                continue

            destination = self._build_destination_path(safe_filename)
            shutil.copy2(attachment_path, destination)
            saved_files.append(destination.name)

        return saved_files

    def save_allowed_attachments(self, message):
        if isinstance(message, FolderMessage):
            return self._save_folder_message_attachments(message)

        saved_files = []
        attachments = getattr(message, 'Attachments', None)

        if attachments is None or attachments.Count <= 0:
            return saved_files

        for i in range(1, attachments.Count + 1):
            attachment = attachments.Item(i)

            original_filename = str(
                getattr(attachment, 'FileName', '') or
                getattr(attachment, 'DisplayName', '') or
                ''
            ).strip()

            safe_filename = self._build_unique_filename(original_filename)

            if not safe_filename:
                continue

            if not self.is_allowed_file(safe_filename):
                continue

            filepath = self._build_destination_path(safe_filename)
            attachment.SaveAsFile(str(filepath))
            saved_files.append(filepath.name)

        return saved_files

    def save_pdf_attachments(self, message):
        return self.save_allowed_attachments(message)
    

    def _build_unique_filename(self, original_filename: str) -> str:
        path = Path(original_filename)

        base_name = normalize_latin_filename(path.stem) or 'fichier'
        suffix = path.suffix.lower()

        unique_num = str(time.time_ns())

        max_base_len = 240 - len(suffix) - len(unique_num) - 3  # ___
        if max_base_len < 1:
            max_base_len = 1

        base_name = base_name[:max_base_len].strip(' ._')

        return f'{base_name}___{unique_num}{suffix}'
