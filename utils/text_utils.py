import re
import unicodedata


def normalize_latin(text: str) -> str:
    if not text:
        return ''
    text = unicodedata.normalize('NFKD', str(text))
    text = text.encode('ascii', 'ignore').decode('ascii')
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def normalize_latin_filename(text: str) -> str:
    text = normalize_latin(text)
    text = re.sub(r'[<>:"/\\|?*]', '_', text)
    text = re.sub(r'\s+', ' ', text)
    return text[:240].strip(' .')
