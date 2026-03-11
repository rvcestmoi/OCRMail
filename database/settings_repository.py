from __future__ import annotations

from datetime import datetime

from config.db_config import TABLE_SETTINGS
from database.connection import get_connection


class SettingsRepository:
    def __init__(self):
        self.connection = get_connection()

    def ensure_table_exists(self):
        cursor = self.connection.cursor()
        sql = f"""
        IF OBJECT_ID('{TABLE_SETTINGS}', 'U') IS NULL
        BEGIN
            CREATE TABLE {TABLE_SETTINGS} (
                setting_key VARCHAR(100) NOT NULL PRIMARY KEY,
                setting_value NVARCHAR(1000) NULL,
                updated_at DATETIME2(0) NOT NULL CONSTRAINT DF_{TABLE_SETTINGS}_updated_at DEFAULT SYSDATETIME()
            );
        END
        """
        try:
            cursor.execute(sql)
            self.connection.commit()
        except Exception:
            self.connection.rollback()
            raise
        finally:
            cursor.close()

    def get_all_settings(self) -> dict:
        cursor = self.connection.cursor()
        sql = f"SELECT setting_key, setting_value FROM {TABLE_SETTINGS};"
        try:
            cursor.execute(sql)
            rows = cursor.fetchall()
            return {row.setting_key: row.setting_value for row in rows}
        finally:
            cursor.close()

    def set_setting(self, key: str, value: str | None):
        cursor = self.connection.cursor()
        sql = f"""
        MERGE {TABLE_SETTINGS} AS target
        USING (SELECT ? AS setting_key, ? AS setting_value) AS source
        ON target.setting_key = source.setting_key
        WHEN MATCHED THEN
            UPDATE SET
                setting_value = source.setting_value,
                updated_at = SYSDATETIME()
        WHEN NOT MATCHED THEN
            INSERT (setting_key, setting_value, updated_at)
            VALUES (source.setting_key, source.setting_value, SYSDATETIME());
        """
        try:
            cursor.execute(sql, (key, value))
            self.connection.commit()
        except Exception:
            self.connection.rollback()
            raise
        finally:
            cursor.close()

    def set_datetime_setting(self, key: str, value: datetime | None):
        text_value = None if value is None else value.strftime('%Y-%m-%d %H:%M:%S')
        self.set_setting(key, text_value)

    def close(self):
        if self.connection:
            self.connection.close()
