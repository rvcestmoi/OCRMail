from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
SETTINGS_FILE = BASE_DIR / 'config' / 'settings.json'
DB_CONFIG_FILE = BASE_DIR / 'config' / 'db_config.json'


DATE_FORMATS = (
    '%Y-%m-%d %H:%M:%S',
    '%Y-%m-%d',
    '%d/%m/%Y %H:%M:%S',
    '%d/%m/%Y',
)

DEFAULT_ALLOWED_EXTENSIONS = ['.pdf', '.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.gif', '.webp']


def _load_json(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Fichier de configuration introuvable : {path}")

    with path.open('r', encoding='utf-8') as file:
        return json.load(file)


def parse_optional_datetime(value):
    if value is None:
        return None

    if isinstance(value, datetime):
        return value

    text = str(value).strip()
    if not text:
        return None

    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    raise ValueError(
        "Format invalide pour 'mail_date_min'. "
        "Formats acceptes : YYYY-MM-DD, YYYY-MM-DD HH:MM:SS, "
        "DD/MM/YYYY, DD/MM/YYYY HH:MM:SS"
    )


def parse_bool(value, default: bool = False) -> bool:
    if value is None:
        return default

    if isinstance(value, bool):
        return value

    text = str(value).strip().lower()
    if text in {'1', 'true', 'vrai', 'yes', 'oui', 'on'}:
        return True
    if text in {'0', 'false', 'faux', 'no', 'non', 'off'}:
        return False
    return default


def parse_allowed_extensions(value) -> list[str]:
    if value is None:
        return list(DEFAULT_ALLOWED_EXTENSIONS)

    if isinstance(value, (list, tuple, set)):
        raw_extensions = list(value)
    else:
        raw_extensions = str(value).replace(';', ',').split(',')

    normalized_extensions: list[str] = []
    for raw_extension in raw_extensions:
        extension = str(raw_extension).strip().lower()
        if not extension:
            continue
        if not extension.startswith('.'):
            extension = f'.{extension}'
        if extension not in normalized_extensions:
            normalized_extensions.append(extension)

    return normalized_extensions or list(DEFAULT_ALLOWED_EXTENSIONS)


def load_settings() -> dict:
    settings = _load_json(SETTINGS_FILE)
    settings['base_dir'] = str(BASE_DIR)
    settings['mail_date_min'] = parse_optional_datetime(settings.get('mail_date_min'))
    settings['allowed_extensions'] = parse_allowed_extensions(settings.get('allowed_extensions'))
    settings['debug_first_pdf'] = parse_bool(settings.get('debug_first_pdf'), default=False)
    return settings


def load_db_config() -> dict:
    return _load_json(DB_CONFIG_FILE)
