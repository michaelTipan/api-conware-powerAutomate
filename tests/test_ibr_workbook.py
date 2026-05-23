"""Tests servicio IBR_DIARIO."""

import io
from datetime import date

import openpyxl
import pytest

from app.application.services.ibr_workbook import find_ibr_for_date, normalize_ibr_value
from app.application.use_cases.setup_ibr_workbook import IBR_SHEET_NAME


def _ibr_workbook_bytes(rows: list[tuple]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = IBR_SHEET_NAME
    ws.append(["Inicio", "Fin", "Valor"])
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.mark.parametrize(
    "raw,expected",
    [
        ("10.580%", 0.10580),
        ("10,580%", 0.10580),
        ("10.580", 0.10580),
        ("10,580", 0.10580),
        (0.10580, 0.10580),
        ("0,10580", 0.10580),
    ],
)
def test_normalize_ibr_value(raw, expected):
    got = normalize_ibr_value(raw)
    assert got is not None
    assert abs(got - expected) < 1e-6


def test_normalize_ibr_value_invalid_returns_none():
    assert normalize_ibr_value("") is None
    assert normalize_ibr_value("n/a") is None


def test_find_ibr_exact_range():
    raw = _ibr_workbook_bytes(
        [
            (date(2026, 5, 1), date(2026, 5, 31), "10.580%"),
            (date(2026, 6, 1), date(2026, 6, 30), "11.000%"),
        ]
    )
    rate = find_ibr_for_date(raw, date(2026, 5, 22))
    assert rate is not None
    assert abs(rate - 0.10580) < 1e-6


def test_find_ibr_lookup_by_range_boundary():
    raw = _ibr_workbook_bytes([(date(2026, 1, 1), date(2026, 12, 31), 0.12)])
    assert find_ibr_for_date(raw, date(2026, 1, 1)) == pytest.approx(0.12)
    assert find_ibr_for_date(raw, date(2026, 12, 31)) == pytest.approx(0.12)


def test_find_ibr_missing_returns_none():
    raw = _ibr_workbook_bytes([(date(2026, 5, 1), date(2026, 5, 31), "10%")])
    assert find_ibr_for_date(raw, date(2026, 7, 1)) is None
