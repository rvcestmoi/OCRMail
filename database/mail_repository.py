from database.connection import get_connection
from config.db_config import TABLE_LOG_MAIL


class MailRepository:

    def __init__(self):
        self.connection = get_connection()

    def upsert_mail_attachment(
        self,
        message_id: str,
        entry_id: str,
        nom_pdf: str,
        sujet: str,
        expediteur: str,
        date_mail,
        store_id: str | None,
    ):
        merge_sql = f"""
        MERGE {TABLE_LOG_MAIL} AS target
        USING (
            SELECT
                ? AS message_id,
                ? AS entry_id,
                ? AS nom_pdf,
                ? AS sujet,
                ? AS expediteur,
                ? AS date_mail,
                ? AS store_id
        ) AS source
        ON (
            target.entry_id = source.entry_id
            AND target.nom_pdf = source.nom_pdf
        )
        WHEN MATCHED THEN
            UPDATE SET
                message_id = source.message_id,
                sujet = source.sujet,
                expediteur = source.expediteur,
                date_mail = source.date_mail,
                store_id = source.store_id,
                date_creation = SYSDATETIME()
        WHEN NOT MATCHED THEN
            INSERT (
                date_creation,
                date_mail,
                message_id,
                entry_id,
                nom_pdf,
                sujet,
                expediteur,
                store_id
            )
            VALUES (
                SYSDATETIME(),
                source.date_mail,
                source.message_id,
                source.entry_id,
                source.nom_pdf,
                source.sujet,
                source.expediteur,
                source.store_id
            );
        """

        cursor = self.connection.cursor()

        try:
            cursor.execute(
                merge_sql,
                (
                    message_id,
                    entry_id,
                    nom_pdf,
                    sujet,
                    expediteur,
                    date_mail,
                    store_id,
                )
            )
            self.connection.commit()

        except Exception as e:
            self.connection.rollback()
            raise Exception(f"Erreur MERGE XXA_LOGMAIL_228794 : {e}")

        finally:
            cursor.close()

    def close(self):
        if self.connection:
            self.connection.close()