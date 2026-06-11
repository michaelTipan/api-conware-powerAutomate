"""Tests retiro del estado INCOMPLETO y bandeja pagos_incompletos.xlsx."""

from __future__ import annotations

import asyncio
import os
from datetime import date
from pathlib import Path

import pytest

from app.application.config import payment_validation_settings as pvs
from app.application.job_status_enrichment import _FINALIZE_MESSAGES
from app.application.services.review_schema import (
    EstadoLinea,
    EstadoPago,
    apply_legacy_estado_migration,
    DistribucionCols,
)
from app.application.use_cases.setup_payment_followup_workbooks import FILENAME_ADELANTADOS


def test_estado_pago_options_exclude_incompleto():
    assert "INCOMPLETO" not in EstadoPago.ALLOWED
    assert "INCOMPLETO" not in EstadoPago.OPTIONS_ORDERED
    assert EstadoPago.OPTIONS_ORDERED == [
        EstadoPago.ADELANTADO,
        EstadoPago.ATRASADO,
        EstadoPago.NORMAL,
        EstadoPago.REVISION_MANUAL,
    ]


def test_legacy_validar_parcial_no_longer_maps_to_incompleto():
    row = {DistribucionCols.ESTADO_PAGO: "VALIDAR_PARCIAL", DistribucionCols.VALIDAR_PAGO: "SI"}
    apply_legacy_estado_migration(row)
    assert row[DistribucionCols.ESTADO_PAGO] == "VALIDAR_PARCIAL"


def test_incompleto_not_in_legacy_estado_linea_options():
    assert not hasattr(EstadoLinea, "VALIDAR_PARCIAL")


def test_finalize_message_incompleto_not_supported():
    user_msg, next_action = _FINALIZE_MESSAGES["INCOMPLETO_NOT_SUPPORTED"]
    assert "INCOMPLETO" in user_msg
    assert "ya no forma parte" in user_msg
    assert "ABONO" in next_action


def test_config_has_no_followup_incompletos_constant():
    assert not hasattr(pvs, "DEFAULT_FOLLOWUP_INCOMPLETOS")


def test_resolve_followup_only_adelantados_path(monkeypatch):
    monkeypatch.setenv("GRAPH_PAYMENT_VALIDATION_CONTROL_PATH", "00 CONTROL")
    path = pvs.resolve_followup_workbook_path(FILENAME_ADELANTADOS)
    assert path.endswith(FILENAME_ADELANTADOS)
    assert "pagos_incompletos" not in path


def test_app_source_has_no_unjustified_incompleto_references():
    app_root = Path(__file__).resolve().parents[1] / "app"
    allowed_files = {
        "app/application/use_cases/payment_validation_finalize.py",
        "app/application/job_status_enrichment.py",
        "app/application/use_cases/setup_payment_followup_workbooks.py",
    }
    hits: list[str] = []
    for py_file in app_root.rglob("*.py"):
        rel = str(py_file.relative_to(app_root.parent)).replace("\\", "/")
        text = py_file.read_text(encoding="utf-8")
        if "pagos_incompletos" in text and rel not in allowed_files:
            hits.append(f"{rel}: referencia a pagos_incompletos")
            continue
        if "INCOMPLETO" in text and rel not in allowed_files:
            hits.append(f"{rel}: referencia a INCOMPLETO")
    assert hits == [], "Referencias no justificadas en app/: " + "; ".join(hits)


def test_finalize_incompleto_does_not_auto_convert_to_abono():
    from tests.test_finalize_validation import (
        MockGraphClient,
        create_review_workbook,
        make_distrib_row,
        set_env_vars,
    )
    from app.application.use_cases.payment_validation_finalize import finalize_payment_validation

    async def run_test():
        set_env_vars()
        client = MockGraphClient()
        r, _ = make_distrib_row(estado="INCOMPLETO", valor_int=50, abono_k=0, mora=0)
        client.downloaded_files["revision/val_latest.xlsx"] = create_review_workbook(distrib_specs=[(r, None)])
        with pytest.raises(ValueError, match="INCOMPLETO_NOT_SUPPORTED"):
            await finalize_payment_validation(client, "val_latest.xlsx", process_date=date(2026, 5, 10))
        assert "history" not in client.uploaded_files or not any(
            "abono" in k.lower() for k in client.uploaded_files
        )

    asyncio.run(run_test())
