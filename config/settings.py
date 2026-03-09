from config.settings_loader import load_settings

SETTINGS = load_settings()
BASE_DIR = SETTINGS['base_dir']
MAIL_SOURCE_TYPE = SETTINGS.get('mail_source_type', 'folder')
MAIL_INPUT_FOLDER = SETTINGS.get('mail_input_folder', "FASTFACT")
OUTLOOK_FOLDER_PATH = SETTINGS.get('outlook_folder_path', [])
DOWNLOAD_FOLDER = SETTINGS.get('download_folder', 'data/PJ')
MAX_PDF = int(SETTINGS.get('max_pdf', 50))
ALLOWED_EXTENSIONS = SETTINGS.get('allowed_extensions', ['.pdf'])