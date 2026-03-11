from __future__ import annotations

import hashlib
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, List


@dataclass
class FolderMessage:
    entry_id: str
    message_id: str
    subject: str
    sender_email_address: str
    attachments: List[Path]
    received_time: datetime | None = None
    store_id: str | None = None


class FolderMailSource:
    def __init__(self, input_folder: str | Path, sender: str = 'folder@test.local', date_min: datetime | None = None):
        self.input_folder = Path(input_folder)
        self.sender = sender
        self.date_min = date_min

    def get_messages_sorted(self) -> Iterable[FolderMessage]:
        if not self.input_folder.exists():
            raise FileNotFoundError(f"Dossier source introuvable : {self.input_folder}")

        files = [p for p in self.input_folder.iterdir() if p.is_file()]
        files.sort(key=lambda p: p.stat().st_mtime, reverse=True)

        messages: List[FolderMessage] = []
        for file_path in files:
            file_dt = datetime.fromtimestamp(file_path.stat().st_mtime)

            if self.date_min is not None and file_dt < self.date_min:
                continue

            digest = hashlib.md5(str(file_path).encode('utf-8')).hexdigest()[:24].upper()
            entry_id = f"FOLDER_{digest}"
            subject = self.input_folder.name
            messages.append(
                FolderMessage(
                    entry_id=entry_id,
                    message_id=entry_id,
                    subject=subject,
                    sender_email_address=self.sender,
                    attachments=[file_path],
                    received_time=file_dt,
                )
            )
        return messages
