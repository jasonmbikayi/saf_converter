import unicodedata
import re

REPLACEMENTS = {
    "’": "-",   # replace apostrophe-like character with hyphen
    "'": "-",   # replace normal apostrophe with hyphen
    "(": "_",   # parentheses → underscore
    ")": "_",
}

def clean_filename(name: str) -> str:
    """
    Normalize, Clean and standardize a filename by:
    - Removing diacritics 
    - lowercase
    - replace spaces and invalid chars with underscores
    - keep version numbers intact (v2.0 etc.)
    - handle apostrophes as dash instead of deletion
    """
    original = name

    # Normalize Unicode (decompose accents)
    name = unicodedata.normalize("NFD", name)
    # Remove diacritics
    name = "".join(ch for ch in name if unicodedata.category(ch) != "Mn")

    # Apply custom replacements
    for old, new in REPLACEMENTS.items():
        name = name.replace(old, new)

    # Replace spaces with underscores
    name = name.replace(" ", "_")

    # Replace any remaining non-alphanumeric except dot, underscore, or hyphen
    name = re.sub(r"[^a-zA-Z0-9._-]", "_", name)

    # Lowercase for consistency
    name = name.lower()

    # Collapse multiple underscores or hyphens
    name = re.sub(r"_+", "_", name)
    name = re.sub(r"-+", "-", name)

    # Remove underscores or hyphens directly before the extension dot
    name = re.sub(r"[_-]+(?=\.)", "", name)

    # Strip leading/trailing underscores or hyphens
    name = name.strip("_-")

    return original, name # pyright: ignore[reportReturnType]

# Example usage
filenames = [
    "Rome’s_File.txt",
    "Résumé (Final) v2.0.pdf",
    "Café-à-l’Ouest.doc",
    "Test---file__2025!!.txt"
]

for f in filenames:
    original, clean = clean_filename(f)
    print(f"{original} -> {clean}")
# Exambank/ExamsProcess/filename_utils.py