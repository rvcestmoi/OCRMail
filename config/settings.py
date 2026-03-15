from __future__ import annotations

from config.settings_loader import load_settings, parse_allowed_extensions, parse_bool, parse_optional_datetime
from database.settings_repository import SettingsRepository


def _load_runtime_settings() -> dict:
    settings = load_settings()

    settings_repo = SettingsRepository()
    try:
        settings_repo.ensure_table_exists()
        db_settings = settings_repo.get_all_settings()
    finally:
        settings_repo.close()

    settings.update(db_settings)
    settings['mail_date_min'] = parse_optional_datetime(settings.get('mail_date_min'))
    settings['allowed_extensions'] = parse_allowed_extensions(settings.get('allowed_extensions'))
    settings['debug_first_pdf'] = parse_bool(settings.get('debug_first_pdf'), default=False)
    settings['max_pdf'] = int(settings.get('max_pdf', 50))
    return settings


SETTINGS = _load_runtime_settings()
BASE_DIR = SETTINGS['base_dir']
MAIL_SOURCE_TYPE = SETTINGS.get('mail_source_type', 'folder')
MAIL_INPUT_FOLDER = SETTINGS.get('mail_input_folder', 'FASTFACT')
OUTLOOK_FOLDER_PATH = SETTINGS.get('outlook_folder_path', [])
OUTLOOK_MAILBOX = SETTINGS.get('outlook_mailbox', 'invoice@ed-trans.com')
DOWNLOAD_FOLDER = SETTINGS.get('download_folder', 'data/PJ')
MAX_FILES_TO_FETCH = int(SETTINGS.get('max_files_to_fetch', 50))
ALLOWED_EXTENSIONS = SETTINGS.get('allowed_extensions', ['.pdf'])
MAIL_DATE_MIN = SETTINGS.get('mail_date_min')
DEBUG_FIRST_PDF = SETTINGS.get('debug_first_pdf', False)