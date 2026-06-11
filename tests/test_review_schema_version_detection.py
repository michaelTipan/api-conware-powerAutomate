"""Detección estricta de ReviewSchemaVersion 2 en encabezados Distribucion."""

import pytest

from app.application.services.review_schema import (
    REVIEW_SCHEMA_VERSION,
    REVIEW_SCHEMA_VERSION_V1,
    DistribucionCols,
    detect_distrib_schema_version_from_headers,
)


def _headers_v2(*, capital_label: str = DistribucionCols.ABONO_A_CAPITAL) -> list[str]:
    return [
        DistribucionCols.MORA_A_APLICAR,
        capital_label,
        DistribucionCols.OTROS_VALORES,
        DistribucionCols.SALDO_POR_ASIGNAR,
    ]


def test_schema_v2_canonical_headers():
    assert detect_distrib_schema_version_from_headers(_headers_v2()) == REVIEW_SCHEMA_VERSION


def test_schema_v2_legacy_abono_k_compatible():
    headers = _headers_v2(capital_label="Abono a K")
    assert detect_distrib_schema_version_from_headers(headers) == REVIEW_SCHEMA_VERSION


def test_schema_v1_only_mora_without_capital():
    headers = [
        DistribucionCols.MORA_A_APLICAR,
        DistribucionCols.OTROS_VALORES,
        DistribucionCols.SALDO_POR_ASIGNAR,
    ]
    assert detect_distrib_schema_version_from_headers(headers) == REVIEW_SCHEMA_VERSION_V1


def test_schema_invalid_only_capital_without_mora():
    headers = [
        DistribucionCols.ABONO_A_CAPITAL,
        DistribucionCols.OTROS_VALORES,
        DistribucionCols.SALDO_POR_ASIGNAR,
    ]
    assert detect_distrib_schema_version_from_headers(headers) == REVIEW_SCHEMA_VERSION_V1


@pytest.mark.parametrize(
    "missing",
    [
        DistribucionCols.OTROS_VALORES,
        DistribucionCols.SALDO_POR_ASIGNAR,
    ],
)
def test_schema_invalid_mora_and_capital_but_missing_required_column(missing: str):
    headers = [
        DistribucionCols.MORA_A_APLICAR,
        DistribucionCols.ABONO_A_CAPITAL,
        DistribucionCols.OTROS_VALORES,
        DistribucionCols.SALDO_POR_ASIGNAR,
    ]
    headers = [h for h in headers if h != missing]
    assert detect_distrib_schema_version_from_headers(headers) == REVIEW_SCHEMA_VERSION_V1


def test_schema_v2_legacy_intereses_de_mora_counts_as_otros():
    headers = [
        DistribucionCols.MORA_A_APLICAR,
        DistribucionCols.ABONO_A_CAPITAL,
        "Intereses de mora",
        DistribucionCols.SALDO_POR_ASIGNAR,
    ]
    assert detect_distrib_schema_version_from_headers(headers) == REVIEW_SCHEMA_VERSION


def test_schema_v1_does_not_map_mora_header_to_capital():
    headers = [DistribucionCols.MORA_A_APLICAR, DistribucionCols.OTROS_VALORES, DistribucionCols.SALDO_POR_ASIGNAR]
    assert DistribucionCols.ABONO_A_CAPITAL not in headers
    assert detect_distrib_schema_version_from_headers(headers) == REVIEW_SCHEMA_VERSION_V1
