from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

BASE_DIR = Path(__file__).resolve().parents[1]
SETTINGS_FILE = BASE_DIR / 'config' / 'settings.json'


DATE_FORMATS = (
    '%Y-%m-%d %H:%M:%S',
    '%Y-%m-%d',
    '%d/%m/%Y %H:%M:%S',
    '%d/%m/%Y',
)

DEFAULT_ALLOWED_EXTENSIONS = ['.pdf', '.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.gif', '.webp']

SETTING_PATHS = {
    'driver': ('database', 'driver'),
    'server': ('database', 'server'),
    'database': ('database', 'database'),
    'username': ('database', 'username'),
    'password': ('database', 'password'),
    'trusted_connection': ('database', 'trusted_connection'),
    'table_log_mail': ('database', 'table_log_mail'),
    'table_settings': ('database', 'table_settings'),
    'mail_source_type': ('mail', 'source_type'),
    'mail_input_folder': ('mail', 'input_folder'),
    'outlook_mailbox': ('outlook', 'mailbox'),
    'outlook_folder_path': ('outlook', 'folder_path'),
    'download_folder': ('folders', 'download_folder'),
    'max_files_to_fetch': ('processing', 'max_files_to_fetch'),
    'allowed_extensions': ('processing', 'allowed_extensions'),
    'mail_date_min': ('processing', 'mail_date_min'),
    'debug_first_pdf': ('processing', 'debug_first_pdf'),
}


def _load_json(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Fichier de configuration introuvable : {path}")

    with path.open('r', encoding='utf-8') as file:
        return json.load(file)


def _save_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    with NamedTemporaryFile('w', encoding='utf-8', delete=False, dir=str(path.parent), suffix='.tmp') as tmp:
        json.dump(data, tmp, indent=2, ensure_ascii=False)
        tmp.write('\n')
        temp_path = Path(tmp.name)

    temp_path.replace(path)


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


def _copy_section(settings: dict, raw_settings: dict, section_name: str, key_map: dict[str, str]) -> None:
    section = raw_settings.get(section_name)
    if not isinstance(section, dict):
        return

    for target_key, source_key in key_map.items():
        if source_key in section:
            settings[target_key] = section[source_key]


def load_raw_settings() -> dict:
    return _load_json(SETTINGS_FILE)


def load_settings() -> dict:
    raw_settings = load_raw_settings()
    settings = {k: v for k, v in raw_settings.items() if not isinstance(v, dict)}
    settings['base_dir'] = str(BASE_DIR)

    _copy_section(
        settings,
        raw_settings,
        'database',
        {
            'driver': 'driver',
            'server': 'server',
            'database': 'database',
            'username': 'username',
            'password': 'password',
            'trusted_connection': 'trusted_connection',
            'table_log_mail': 'table_log_mail',
            'table_settings': 'table_settings',
        },
    )
    _copy_section(
        settings,
        raw_settings,
        'mail',
        {
            'mail_source_type': 'source_type',
            'mail_input_folder': 'input_folder',
        },
    )
    _copy_section(
        settings,
        raw_settings,
        'outlook',
        {
            'outlook_mailbox': 'mailbox',
            'outlook_folder_path': 'folder_path',
        },
    )
    _copy_section(
        settings,
        raw_settings,
        'folders',
        {
            'download_folder': 'download_folder',
        },
    )
    _copy_section(
        settings,
        raw_settings,
        'processing',
        {
            'max_files_to_fetch': 'max_files_to_fetch',
            'allowed_extensions': 'allowed_extensions',
            'mail_date_min': 'mail_date_min',
            'debug_first_pdf': 'debug_first_pdf',
        },
    )

    settings['mail_date_min'] = parse_optional_datetime(settings.get('mail_date_min'))
    settings['allowed_extensions'] = parse_allowed_extensions(settings.get('allowed_extensions'))
    settings['debug_first_pdf'] = parse_bool(settings.get('debug_first_pdf'), default=False)
    settings['trusted_connection'] = parse_bool(settings.get('trusted_connection'), default=False)
    settings['max_files_to_fetch'] = int(settings.get('max_files_to_fetch', settings.get('max_pdf', 50)))
    return settings


def load_db_config() -> dict:
    settings = load_settings()
    return {
        'driver': settings.get('driver', 'ODBC Driver 17 for SQL Server'),
        'server': settings.get('server', ''),
        'database': settings.get('database', ''),
        'username': settings.get('username', ''),
        'password': settings.get('password', ''),
        'trusted_connection': settings.get('trusted_connection', False),
        'table_log_mail': settings.get('table_log_mail', 'XXA_LOGMAIL_228794'),
        'table_settings': settings.get('table_settings', 'XXA_SETTINGS_228794'),
    }


def _serialize_setting_value(value):
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d %H:%M:%S')
    return value


def save_setting(key: str, value) -> None:
    raw_settings = load_raw_settings()
    serialized_value = _serialize_setting_value(value)

    setting_path = SETTING_PATHS.get(key)
    if setting_path is None:
        raw_settings[key] = serialized_value
        _save_json(SETTINGS_FILE, raw_settings)
        return

    section_name, setting_name = setting_path
    section = raw_settings.get(section_name)
    if not isinstance(section, dict):
        section = {}
        raw_settings[section_name] = section

    section[setting_name] = serialized_value
    _save_json(SETTINGS_FILE, raw_settings)
