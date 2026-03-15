from __future__ import annotations

from config.settings_loader import load_settings


SETTINGS = load_settings()
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
