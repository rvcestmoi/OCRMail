# utils/text_utils.py

import unicodedata


def normalize_latin(text: str) -> str:
    """
    Convertit un texte Unicode en latin lisible :
    - supprime accents
    - supprime caractères non latins
    """
    if not text:
        return ""

    # Normalisation Unicode (é → e, Ž → Z, etc.)
    text = unicodedata.normalize("NFKD", text)

    # Garde uniquement les caractères ASCII
    text = text.encode("ascii", "ignore").decode("ascii")

    # Nettoyage final
    return text.strip()
