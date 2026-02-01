# database/connection.py

import pyodbc
from config.db_config import DB_CONFIG


def get_connection():
    """
    Crée et retourne une connexion SQL Server via pyodbc
    """
    try:
        if DB_CONFIG.get("trusted_connection"):
            conn_str = (
                f"DRIVER={{{DB_CONFIG['driver']}}};"
                f"SERVER={DB_CONFIG['server']};"
                f"DATABASE={DB_CONFIG['database']};"
                "Trusted_Connection=yes;"
            )
        else:
            conn_str = (
                f"DRIVER={{{DB_CONFIG['driver']}}};"
                f"SERVER={DB_CONFIG['server']};"
                f"DATABASE={DB_CONFIG['database']};"
                f"UID={DB_CONFIG['username']};"
                f"PWD={DB_CONFIG['password']};"
            )

        connection = pyodbc.connect(conn_str, autocommit=False)
        return connection

    except pyodbc.Error as e:
        raise Exception(f"❌ Erreur de connexion SQL Server : {e}")
