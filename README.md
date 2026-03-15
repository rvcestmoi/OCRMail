# OCRMail

Le projet lit maintenant **toute sa configuration depuis `config/settings.json`**.

Il n'y a plus de dependance fonctionnelle a une table SQL de settings pour :
- la connexion SQL Server ;
- la boite Outlook et le chemin du dossier ;
- le dossier de telechargement ;
- la date minimale de recuperation ;
- les extensions autorisees ;
- le mode debug.

## Structure du fichier `config/settings.json`

```json
{
  "database": {
    "driver": "ODBC Driver 17 for SQL Server",
    "server": "THINKPAD-hro1",
    "database": "DB228794",
    "username": "gobabygo",
    "password": "comeback",
    "trusted_connection": false,
    "table_log_mail": "XXA_LOGMAIL_228794"
  },
  "mail": {
    "source_type": "outlook",
    "input_folder": "FASTFACT"
  },
  "outlook": {
    "mailbox": "invoice@ed-trans.com",
    "folder_path": ["FASTFACT"]
  },
  "folders": {
    "download_folder": "C:\\temp\\OCRMail"
  },
  "processing": {
    "max_files_to_fetch": 50,
    "allowed_extensions": [".pdf", ".png", ".jpg", ".jpeg"],
    "mail_date_min": "2026-03-01 00:00:00",
    "debug_first_pdf": false
  }
}
```

## Date minimale

`processing.mail_date_min` accepte :
- `YYYY-MM-DD`
- `YYYY-MM-DD HH:MM:SS`
- `DD/MM/YYYY`
- `DD/MM/YYYY HH:MM:SS`

Si la valeur est vide ou `null`, aucun filtre n'est applique.

En fin de traitement complet, le programme met a jour automatiquement `processing.mail_date_min` dans `config/settings.json`.

Si la limite `processing.max_files_to_fetch` est atteinte, la date n'est pas mise a jour pour eviter de sauter des messages.

## Compatibilite

Le chargeur garde une compatibilite avec les anciennes cles plates si besoin, mais la structure recommandee est celle ci-dessus.
