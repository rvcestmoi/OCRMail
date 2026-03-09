from config.settings_loader import load_db_config

DB_CONFIG = load_db_config()
TABLE_LOG_MAIL = DB_CONFIG.get('table_log_mail', 'XXA_LOGMAIL_228794')
