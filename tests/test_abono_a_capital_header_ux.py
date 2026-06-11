"""Schema v2: columnas Mora a aplicar y Abono a capital independientes."""

import asyncio
import io
from datetime import date, datetime

import openpyxl
import pytest
from openpyxl.utils import get_column_letter

from app.application.services.review_schema import (
    ApplicationSubtype,
    ControlCols,
    DistribucionCols,
    REVIEW_SCHEMA_VERSION,
    ReviewSheets,
    TipoAplicacionVisible,
    detect_distrib_schema_version_from_headers,
    normalize_distrib_row_keys,
    resolve_application_policy,
)
from app.application.use_cases.payment_validation_finalize import finalize_payment_validation
from tests.test_finalize_validation import (
    MockGraphClient,
    create_review_workbook,
    make_distrib_row,
    set_env_vars,
)
from tests.test_generate_validation import (
    _dist_col,
    _first_data_row,
    _header_row_index,
    run_generate,
)
from tests.test_pago_y_abono_capital_finalize import make_pago_y_abono_capital_row


def _run(coro):
    return asyncio.run(coro)


def test_generate_v2_has_both_mora_and_capital_headers():
    _, _, wb = run_generate(
        [[datetime(2025, 12, 23), 25443565, "GEOEXCON", ""]],
        date(2025, 12, 23),
    )
    ws = wb[ReviewSheets.DISTRIBUCION_PAGOS]
    hr = _header_row_index(ws, DistribucionCols.ID_PAGO)
    headers = [c.value for c in ws[hr]]
    assert DistribucionCols.MORA_A_APLICAR in headers
    assert DistribucionCols.ABONO_A_CAPITAL in headers
    assert headers.count(DistribucionCols.MORA_A_APLICAR) == 1
    assert headers.count(DistribucionCols.ABONO_A_CAPITAL) == 1
    assert detect_distrib_schema_version_from_headers(headers) == REVIEW_SCHEMA_VERSION


def test_generate_control_writes_review_schema_version_2():
    _, _, wb = run_generate(
        [[datetime(2025, 12, 23), 25443565, "GEOEXCON", ""]],
        date(2025, 12, 23),
    )
    ws = wb[ReviewSheets.CONTROL]
    found = False
    for row in ws.iter_rows(values_only=True):
        if row and row[0] == ControlCols.ROW_REVIEW_SCHEMA_VERSION:
            assert int(row[1]) == REVIEW_SCHEMA_VERSION
            found = True
    assert found


def test_saldo_formula_uses_all_four_assignment_columns():
    _, _, wb = run_generate(
        [[datetime(2025, 12, 23), 25443565, "GEOEXCON", ""]],
        date(2025, 12, 23),
    )
    ws = wb[ReviewSheets.DISTRIBUCION_PAGOS]
    dr = _first_data_row(ws, DistribucionCols.ID_PAGO)
    formula = ws.cell(dr, _dist_col(DistribucionCols.SALDO_POR_ASIGNAR)).value
    for col_name in (
        DistribucionCols.APLICAR_A_EXTRACTO,
        DistribucionCols.MORA_A_APLICAR,
        DistribucionCols.ABONO_A_CAPITAL,
        DistribucionCols.OTROS_VALORES,
    ):
        letter = get_column_letter(_dist_col(col_name))
        assert f"${letter}$" in formula


def test_finalize_accepts_v2_workbook():
    async def run_test():
        set_env_vars()
        client = MockGraphClient()
        r, _ = make_pago_y_abono_capital_row(valor_extracto=800, abono_capital=200)
        client.downloaded_files["revision/val_latest.xlsx"] = create_review_workbook(
            distrib_specs=[(r, None)]
        )
        res = await finalize_payment_validation(client, "val_latest.xlsx", process_date=date(2026, 5, 10))
        assert res["status"] == "success"

    _run(run_test())


def test_finalize_rejects_v1_review_workbook():
    async def run_test():
        set_env_vars()
        client = MockGraphClient()
        r, _ = make_distrib_row(valor_int=70, mora_a_aplicar=10, abono_capital=20)
        headers_v1 = list(DistribucionCols.HEADERS)
        idx_cap = headers_v1.index(DistribucionCols.ABONO_A_CAPITAL)
        headers_v1.pop(headers_v1.index(DistribucionCols.MORA_A_APLICAR))
        wb = openpyxl.Workbook()
        ws_ctrl = wb.active
        ws_ctrl.title = ReviewSheets.CONTROL
        ws_ctrl.append(["Campo", "Valor"])
        ws_ctrl.append([ControlCols.ROW_PROCESAR, "SI"])
        ws_ctrl.append([ControlCols.ROW_REVIEW_SCHEMA_VERSION, 1])
        ws_ctrl.append([ControlCols.ROW_ESTADO, "EN_REVISION"])
        from app.application.services.review_schema import CasosPagoCols

        ws_casos = wb.create_sheet(ReviewSheets.CASOS_PAGO)
        ws_casos.append(CasosPagoCols.HEADERS)
        ws_casos.append(["ID1", None, "CLI", "c", 100, ""])
        ws_dist = wb.create_sheet(ReviewSheets.DISTRIBUCION)
        ws_dist.append(headers_v1)
        vals = dict(zip(DistribucionCols.HEADERS, r))
        row_v1 = [vals.get(h, "") for h in headers_v1]
        ws_dist.append(row_v1)
        buf = io.BytesIO()
        wb.save(buf)
        client.downloaded_files["revision/val_latest.xlsx"] = buf.getvalue()
        with pytest.raises(ValueError, match="review_schema_version_1_requires_regenerate"):
            await finalize_payment_validation(client, "val_latest.xlsx", process_date=date(2026, 5, 10))

    _run(run_test())


def test_legacy_v1_row_normalize_does_not_map_mora_to_capital():
    row = normalize_distrib_row_keys(
        {DistribucionCols.MORA_A_APLICAR: 50, DistribucionCols.APLICAR_A_EXTRACTO: 100},
        schema_version=1,
    )
    assert row[DistribucionCols.MORA_A_APLICAR] == 50
    assert DistribucionCols.ABONO_A_CAPITAL not in row


def test_abono_mora_policy_independent_of_capital_column():
    policy = resolve_application_policy(TipoAplicacionVisible.ABONO_MORA, from_bank=False)
    assert policy.subtipo_aplicacion == ApplicationSubtype.MORA
