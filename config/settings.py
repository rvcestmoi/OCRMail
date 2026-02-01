# config/settings.py

# ==============================
# DOSSIERS
# ==============================

# Dossier racine du projet
BASE_DIR = r"C:\git\OCRMail"

# Dossier de téléchargement des pièces jointes
DOWNLOAD_FOLDER = f"{BASE_DIR}\\data\\PJ"

# ==============================
# TRAITEMENT OUTLOOK
# ==============================

# Nombre maximum de PDF à traiter par exécution
MAX_PDF = 5

# Extensions autorisées
ALLOWED_EXTENSIONS = [".pdf"]
