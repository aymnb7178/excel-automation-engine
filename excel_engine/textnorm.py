"""Text, number and date normalisation helpers (Arabic and English aware)."""

from __future__ import annotations

import math
import re
import unicodedata
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd

_DIGITS = str.maketrans(
    "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹٫٬،",
    "01234567890123456789.,,",
)
_DIACRITICS = re.compile("[ؐ-ًؚ-ٰٟۖ-ۭ]")
_LETTERS = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ى": "ي"})
_MISSING = {"", "nan", "none", "null", "n/a", "na", "#n/a", "-", "--", "غير متوفر", "لا يوجد"}
_CURRENCY = re.compile(r"(sar|usd|eur|aed|sr|ر\.?\s?س\.?|ريال|﷼|دولار|[$€£¥])", re.I)
_NUMBER = re.compile(r"\d+(\.\d+)?|\.\d+")


def is_missing(v) -> bool:
    """True for None/NaN/NaT/pd.NA and common textual missing markers ('N/A', '-', '')."""
    if v is None or v is pd.NA or v is pd.NaT:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, str):
        return v.strip().lower() in _MISSING
    return False


def to_western_digits(s: str) -> str:
    """Arabic-Indic / Persian digits and Arabic separators -> ASCII equivalents."""
    return s.translate(_DIGITS)


def normalize_text(s: str, *, case: str | None = None, arabic_letters: bool = False) -> str:
    """Light normalisation: NFKC, strip tatweel and diacritics, Western digits, collapse spaces.

    ``arabic_letters=True`` also unifies alef variants and alef maksura.
    """
    t = unicodedata.normalize("NFKC", str(s))
    t = _DIACRITICS.sub("", t.replace("ـ", ""))
    t = to_western_digits(t)
    if arabic_letters:
        t = t.translate(_LETTERS)
    t = re.sub(r"\s+", " ", t).strip()
    if case == "title":
        t = t.title()
    elif case == "upper":
        t = t.upper()
    elif case == "lower":
        t = t.lower()
    return t


def match_key(s) -> str:
    """Aggressive key for comparisons: lower-case, unified Arabic letters, no punctuation."""
    t = normalize_text(s, arabic_letters=True).lower().replace("ة", "ه")
    t = re.sub(r"[^\w\s]|_", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def header_key(s) -> str:
    """Key for header matching: 'Order ID', 'order_id' and 'ORDER-ID' all give 'orderid'."""
    return re.sub(r"\s+", "", match_key(s))


def parse_number(v) -> float | None:
    """Parse numbers with currency symbols, Arabic-Indic digits, separators, (negatives)."""
    if is_missing(v) or isinstance(v, (bool, np.bool_)):
        return None
    if isinstance(v, (int, float, np.integer, np.floating)):
        x = float(v)
        return x if math.isfinite(x) else None
    if not isinstance(v, str):
        return None
    s = to_western_digits(unicodedata.normalize("NFKC", v)).strip()
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1]
    s = _CURRENCY.sub("", s)
    s = re.sub(r"[\s ]", "", s)
    if s.endswith("-"):
        neg, s = True, s[:-1]
    if s.startswith("-"):
        neg, s = True, s[1:]
    elif s.startswith("+"):
        s = s[1:]
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):  # 1.234,56
            s = s.replace(".", "").replace(",", ".")
        else:  # 1,234.56
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if all(len(p) == 3 for p in parts[1:]):  # 1,250 / 1,250,000 -> thousands
            s = s.replace(",", "")
        else:  # 12,5 -> decimal comma
            s = s.replace(",", ".")
    if not _NUMBER.fullmatch(s):
        return None
    x = float(s)
    return -x if neg else x


def parse_integer(v) -> int | None:
    x = parse_number(v)
    if x is None or abs(x - round(x)) > 1e-9:
        return None
    return int(round(x))


_ISO_FORMATS = [
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
    "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M",
]
_TEXT_FORMATS = [
    "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y", "%d-%b-%Y", "%d-%b-%y",
    "%d %b, %Y", "%b %d %Y", "%B %d %Y",
]
_DMY = ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y",
        "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M"]
_MDY = [f.replace("%d/%m", "%m/%d").replace("%d-%m", "%m-%d").replace("%d.%m", "%m.%d") for f in _DMY]


def parse_date(v, dayfirst: bool = True) -> pd.Timestamp | None:
    """Parse many date spellings; returns a normalised (midnight) Timestamp or None."""
    if is_missing(v) or isinstance(v, (bool, np.bool_)):
        return None
    if isinstance(v, (datetime, date, pd.Timestamp)):
        try:
            return pd.Timestamp(v).normalize()
        except (ValueError, OverflowError):
            return None
    if isinstance(v, (int, float, np.integer, np.floating)):
        return _serial(float(v))
    if not isinstance(v, str):
        return None
    s = re.sub(r"\s+", " ", to_western_digits(unicodedata.normalize("NFKC", v))).strip()
    if re.fullmatch(r"\d{5}(\.\d+)?", s):  # Excel serial stored as text
        return _serial(float(s))
    if re.fullmatch(r"\d{8}", s):
        formats = ["%Y%m%d"]
    else:
        formats = _ISO_FORMATS + _TEXT_FORMATS + (_DMY + _MDY if dayfirst else _MDY + _DMY)
    for fmt in formats:
        try:
            return pd.Timestamp(datetime.strptime(s, fmt)).normalize()
        except ValueError:
            continue
    return None


def _serial(x: float) -> pd.Timestamp | None:
    if 20000 <= x <= 80000:  # roughly 1954-2119; avoids treating small numbers as dates
        return pd.Timestamp(datetime(1899, 12, 30) + timedelta(days=int(x)))
    return None
