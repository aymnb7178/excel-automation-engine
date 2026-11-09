import pandas as pd
import pytest

from excel_engine.textnorm import (header_key, is_missing, match_key, normalize_text, parse_date,
                                   parse_integer, parse_number)


@pytest.mark.parametrize("raw,expected", [
    ("SAR 1,250.50", 1250.5), ("١٬٢٥٠٫٥٠ ر.س", 1250.5), ("$1,000", 1000.0), ("(45.5)", -45.5),
    ("1.234,56", 1234.56), ("12,5", 12.5), ("1,250", 1250.0), ("٣٠", 30.0), ("5.0", 5.0), (7, 7.0),
])
def test_parse_number_ok(raw, expected):
    assert parse_number(raw) == expected


@pytest.mark.parametrize("raw", ["free", "abc", "", "N/A", None, "1.2.3", float("nan"), True])
def test_parse_number_rejects(raw):
    assert parse_number(raw) is None


def test_parse_integer_rejects_fractions():
    assert parse_integer("٥") == 5 and parse_integer("5.5") is None


@pytest.mark.parametrize("raw,expected", [
    ("2025-03-12", "2025-03-12"), ("12/03/2025", "2025-03-12"), ("١٢/٠٣/٢٠٢٥", "2025-03-12"),
    ("12 Mar 2025", "2025-03-12"), ("March 12, 2025", "2025-03-12"), ("2025/03/12", "2025-03-12"),
    ("20250312", "2025-03-12"), (45728, "2025-03-12"),
])
def test_parse_date_formats(raw, expected):
    assert parse_date(raw) == pd.Timestamp(expected)


def test_parse_date_dayfirst_flag_and_invalid():
    assert parse_date("03/04/2025", dayfirst=True) == pd.Timestamp("2025-04-03")
    assert parse_date("03/04/2025", dayfirst=False) == pd.Timestamp("2025-03-04")
    assert all(parse_date(x) is None for x in ["31/02/2025", "TBD", "", "13/13/2025"])


def test_arabic_normalisation():
    assert normalize_text("أَحْمَد  ", arabic_letters=True) == "احمد"
    assert normalize_text("مــحمد") == "محمد"
    assert match_key("أحمد") == match_key("احمد")
    assert header_key("Order ID") == header_key("order_id") == header_key("ORDER-ID")


def test_is_missing_markers():
    assert all(is_missing(x) for x in [None, float("nan"), "", " N/A ", "-", "غير متوفر", pd.NaT])
    assert not is_missing(0) and not is_missing("0")
