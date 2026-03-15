# OCRMail

Structure refactorisee :
- les chemins et options sont lus depuis `config/settings.json`
- la connexion SQL Server est lue depuis `config/db_config.json`
- pour les tests, la source des mails peut etre un dossier
- pour repasser sur Outlook, mettre `"mail_source_type": "outlook"` dans `config/settings.json`

## Nouveau filtre par date

Dans `config/settings.json`, tu peux maintenant definir une date minimale de traitement via `mail_date_min`.

Exemple :

```json
{
  "mail_source_type": "outlook",
  "outlook_folder_path": ["FASTFACT"],
  "download_folder": "C:\\Users\\hrouillard\\Documents\\clients\\ED trans\\OCR\\modeles2",
  "max_pdf": 50,
  "allowed_extensions": [".pdf"],
  "mail_date_min": "2026-03-01 00:00:00",
  "debug_first_pdf": false
}
```

Formats acceptes pour `mail_date_min` :
- `YYYY-MM-DD`
- `YYYY-MM-DD HH:MM:SS`
- `DD/MM/YYYY`
- `DD/MM/YYYY HH:MM:SS`

Comportement :
- en mode `outlook`, seuls les mails recus a partir de cette date sont lus
- en mode `folder`, seuls les fichiers modifies a partir de cette date sont pris
- si `mail_date_min` est vide ou absent, aucun filtre n'est applique

## Debug du premier PDF

Le mode debug n'est plus force dans le code.

Pour afficher uniquement le premier PDF trouve puis quitter, utilise :

```json
"debug_first_pdf": true
```

Sinon laisse :

```json
"debug_first_pdf": false
```


## Settings SQL Server

La date minimale n'est plus lue en priorité dans `config/settings.json`.
Elle est maintenant chargée depuis la table SQL Server `XXA_SETTINGS_228794`.

### Clé utilisée

- `mail_date_min` : date minimale de lecture des mails au format `YYYY-MM-DD HH:MM:SS`

### Script SQL

Voir `database/create_settings_table.sql`.

### Comportement

- au démarrage, le programme crée la table si elle n'existe pas ;
- les settings SQL remplacent les valeurs du fichier JSON si la même clé existe ;
- après un traitement complet, `mail_date_min` est mise à jour avec la date la plus récente lue ;
- si `MAX_PDF` est atteint, la valeur n'est pas mise à jour pour éviter de sauter des messages.


## Pieces jointes autorisees

Le programme peut maintenant recuperer les PDF et les images autorisees.

Extensions autorisees par defaut :
- `.pdf`
- `.png`
- `.jpg`
- `.jpeg`
- `.tif`
- `.tiff`
- `.bmp`
- `.gif`
- `.webp`

Tu peux surcharger cette liste depuis :
- `config/settings.json` avec un tableau JSON ;
- la table SQL Server des settings avec la cle `allowed_extensions` et une valeur de type `.pdf,.png,.jpg`.

Exemple JSON :

```json
{
  "allowed_extensions": [".pdf", ".png", ".jpg", ".jpeg"]
}
```

Exemple SQL :

```sql
MERGE XXA_SETTINGS_228794 AS target
USING (SELECT 'allowed_extensions' AS setting_key, '.pdf,.png,.jpg,.jpeg' AS setting_value) AS source
ON target.setting_key = source.setting_key
WHEN MATCHED THEN
    UPDATE SET setting_value = source.setting_value, updated_at = SYSDATETIME()
WHEN NOT MATCHED THEN
    INSERT (setting_key, setting_value, updated_at)
    VALUES (source.setting_key, source.setting_value, SYSDATETIME());
```

Remarque : les images inline masquees d'Outlook sont ignorees pour eviter de recuperer les logos de signature.
