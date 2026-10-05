"""Normalisation helpers. These produce COMPARISON keys only; original values are never altered."""
import re
import unicodedata

_HIDDEN = dict.fromkeys(list(range(0x200B, 0x2010)) + [0xFEFF, 0x2060], None)


def clean_text(v):
    """Strip hidden characters / collapse whitespace. Returns '' for None."""
    if v is None:
        return ""
    s = str(v).translate(_HIDDEN)
    s = unicodedata.normalize("NFKC", s)
    return re.sub(r"\s+", " ", s).strip()


def dummy_key(v):
    """Comparison key for a dummy number: digits only, tolerant of float/str/leading-zero formats.
    Returns None if the value has no usable digits."""
    if v is None:
        return None
    if isinstance(v, float):
        if v != v:  # NaN
            return None
        if v.is_integer():
            v = int(v)
    s = clean_text(v)
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"[\s,\-_]", "", s)
    if not s.isdigit():
        return None
    return s.lstrip("0") or "0"


def code_key(v):
    """Comparison key for a subject code: upper-case alphanumerics only."""
    return re.sub(r"[^A-Z0-9]", "", clean_text(v).upper())


def name_key(v):
    """Comparison key for a subject name: upper-case alphanumerics only ('&' -> 'AND')."""
    s = clean_text(v).upper().replace("&", " AND ")
    return re.sub(r"[^A-Z0-9]", "", s)


SUBJECT_CODE_RE = re.compile(r"\b(\d{3}[A-Z]{3}\d{2})\b", re.I)


def extract_codes(text):
    """All official-format subject codes (e.g. 422BMT03) found in free text."""
    return [m.upper() for m in SUBJECT_CODE_RE.findall(clean_text(text))]
