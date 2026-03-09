from __future__ import annotations

import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
SETTINGS_FILE = BASE_DIR / 'config' / 'settings.json'
DB_CONFIG_FILE = BASE_DIR / 'config' / 'db_config.json'


def _load_json(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Fichier de configuration introuvable : {path}")

    with path.open('r', encoding='utf-8') as file:
        return json.load(file)


def load_settings() -> dict:
    settings = _load_json(SETTINGS_FILE)
    settings['base_dir'] = str(BASE_DIR)
    return settings


def load_db_config() -> dict:
    return _load_json(DB_CONFIG_FILE)
