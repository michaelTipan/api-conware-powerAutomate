"""Dry-run Fase 4: preflight y cuadre ABONO."""

import asyncio
import io
import json
from datetime import date
from decimal import Decimal
from urllib.parse import unquote

import openpyxl
import pytest
from pypdf import PdfWriter

from app.application.services.abono_dry_run import (
    ABONO_ASIENTO_DUPLICADO,
    ABONO_ASIENTOS_NO_CUADRAN,
    ABONO_RECONCILIATION_TOLERANCE,
    ABONO_SCHEDULE_RULE_NOT_CONFIGURED,
    build_abono_group_from_manifest_output,
    detect_duplicate_asiento_paths,
    infer_manifest_tipo_aplicacion,
    reconcile_abono_group,
    validate_abono_group_structure,
)
from app.application.services.accounting_pdf_parser import ACCOUNT_VALOR_PAGADO_CLIENTE
from app.application.use_cases.amortization_fill_dry_run import run_amortization_fill_dry_run
from tests.test_amortization_fill_dry_run import (
    MockGraphDryRun,
    _accounting_text,
    _amort_table_bytes,
    _base_files,
    _hist_bytes,
    _ibr_bytes,
)


def _asiento_pdf() -> bytes:
    w = PdfWriter()
    w.add_blank_page(width=72, height=72)
    buf = io.BytesIO()
    w.write(buf)
    return buf.getvalue()


def _abono_manifest_output(
    *,
    id_pago: str = "AB1",
    monto_banco: float = 50_000_000.0,
    credits: tuple[str, ...] = ("258",),
    asiento_paths: tuple[str, ...] | None = None,
    montos: tuple[float, ...] | None = None,
) -> dict:
    paths = asiento_paths or tuple(
        f"clientes/EQUINORTE/CREDITO # {c}/ASIENTOS CONTABLES CRED {c}/asiento {c}.pdf"
        for c in credits
    )
    items = []
    for i, cred in enumerate(credits):
        items.append(
            {
                "credito": cred,
                "tipo_aplicacion": "ABONO",
                "ruta_tabla_amortizacion": f"TABLAS/amort_{cred}.xlsx",
                "ruta_unidad_credito": f"clientes/EQUINORTE/CREDITO # {cred}",
                "ruta_asientos_contables": f"clientes/EQUINORTE/CREDITO # {cred}/ASIENTOS CONTABLES CRED {cred}",
                "asiento_pdf_paths": [paths[i]],
                "extracto_pdf_paths": [],
            }
        )
    return {
        "id_pago": id_pago,
        "cliente": "EQUINORTE",
        "credito": ", ".join(credits),
        "tipo_aplicacion": "ABONO",
        "requiere_extracto": False,
        "monto_banco": monto_banco,
        "fecha_banco": "2026-05-22",
        "creditos_seleccionados": list(credits),
        "credit_items": items,
        "output_relative_path": "OUT/abono.pdf",
    }


def _abono_dry_run_files(
    *,
    abono_output: dict,
    pago_manifest_output: dict | None = None,
    fecha: date = date(2026, 5, 22),
) -> dict[str, bytes]:
    outputs = [abono_output]
    if pago_manifest_output:
        outputs.insert(0, pago_manifest_output)
    manifest = {
        "report_date_iso": fecha.isoformat(),
        "historico_excel_path": "HIST/cartera.xlsx",
        "outputs": outputs,
        "skipped": [],
    }
    manifest_key = f"LOGS/merge_manifest_{fecha.isoformat()}.json"
    files: dict[str, bytes] = {
        "CTL/dummy.xlsx": b"x",
        manifest_key: json.dumps(manifest).encode("utf-8"),
        "CTL/IBR_DIARIO.xlsx": _ibr_bytes(),
    }
    for cred in abono_output.get("creditos_seleccionados") or []:
        tabla = f"TABLAS/amort_{cred}.xlsx"
        files[tabla] = _amort_table_bytes(fecha)
        for ci in abono_output.get("credit_items") or []:
            if str(ci.get("credito")) == str(cred):
                for ap in ci.get("asiento_pdf_paths") or []:
                    files[ap] = _asiento_pdf()
    if pago_manifest_output:
        hist = _hist_bytes(
            pago_manifest_output["id_pago"],
            pago_manifest_output.get("credito", "CREDITO # 258"),
            "TABLAS/amort_258.xlsx",
            fecha,
        )
        files["HIST/cartera.xlsx"] = hist
        files["TABLAS/amort_258.xlsx"] = _amort_table_bytes(fecha)
        ap = pago_manifest_output.get("asiento_pdf_path", "")
        if ap:
            files[ap] = _asiento_pdf()
    else:
        files["HIST/cartera.xlsx"] = _hist_bytes("legacy", "258", "", fecha)
    return files


@pytest.fixture(autouse=True)
def env_sharepoint(monkeypatch):
    monkeypatch.setenv("GRAPH_SHAREPOINT_SITE_SEARCH", "TEST")
    monkeypatch.setenv("GRAPH_SHAREPOINT_DRIVE_NAME", "")
    monkeypatch.setenv("GRAPH_PAYMENT_VALIDATION_CONTROL_PATH", "CTL")
    monkeypatch.setenv("GRAPH_PAYMENT_VALIDATION_LOGS_PATH", "LOGS")
    monkeypatch.setenv("GRAPH_IBR_DIARIO_PATH", "CTL/IBR_DIARIO.xlsx")


def test_infer_manifest_legacy_as_pago():
    assert infer_manifest_tipo_aplicacion({}) == "PAGO"
    assert infer_manifest_tipo_aplicacion({"tipo_aplicacion": "ABONO"}) == "ABONO"


def test_validate_abono_structure_rejects_duplicate_credit():
    out = _abono_manifest_output(credits=("258", "258"))
    group = build_abono_group_from_manifest_output(out, bank_code="banco_bogota", process_key="pk")
    errors = validate_abono_group_structure(group)
    assert any(e["error_code"] == "ABONO_CREDITO_DUPLICADO" for e in errors)


def test_detect_duplicate_asiento_across_credits():
    shared = "clientes/X/asiento_shared.pdf"
    out = _abono_manifest_output(
        credits=("258", "265"),
        asiento_paths=(shared, shared),
    )
    g = build_abono_group_from_manifest_output(out, bank_code="banco_bogota", process_key="pk")
    dupes = detect_duplicate_asiento_paths([g])
    assert dupes
    assert any("asiento_shared.pdf" in k for k in dupes)


def test_reconcile_abono_exact_match_decimal(monkeypatch):
    out = _abono_manifest_output(monto_banco=50_000_000.0, credits=("258",))
    group = build_abono_group_from_manifest_output(out, bank_code="banco_bogota", process_key="pk")
    monkeypatch.setattr(
        "app.application.services.abono_dry_run.extract_text_from_pdf",
        lambda _b: _accounting_text(),
    )

    async def _run():
        async def _dl(_path: str) -> bytes:
            return _asiento_pdf()

        return await reconcile_abono_group(group, _dl)

    result = asyncio.run(_run())
    assert result.reconciliation_status == "PASSED"
    assert result.diferencia == Decimal("0")
    assert result.tolerancia == ABONO_RECONCILIATION_TOLERANCE


def test_reconcile_abono_fails_outside_tolerance(monkeypatch):
    out = _abono_manifest_output(monto_banco=10_000_000.0, credits=("258",))
    group = build_abono_group_from_manifest_output(out, bank_code="banco_bogota", process_key="pk")
    monkeypatch.setattr(
        "app.application.services.abono_dry_run.extract_text_from_pdf",
        lambda _b: _accounting_text(),
    )

    async def _run():
        async def _dl(_path: str) -> bytes:
            return _asiento_pdf()

        return await reconcile_abono_group(group, _dl)

    result = asyncio.run(_run())
    assert result.reconciliation_status == "FAILED"
    assert any(e["error_code"] == ABONO_ASIENTOS_NO_CUADRAN for e in result.blocking_errors)


def test_dry_run_abono_reconciled_not_required_ready_to_apply(monkeypatch):
    out = _abono_manifest_output()
    files = _abono_dry_run_files(abono_output=out)
    g = MockGraphDryRun(files)
    monkeypatch.setattr(
        "app.application.services.abono_dry_run.extract_text_from_pdf",
        lambda _b: _accounting_text(),
    )
    result = asyncio.run(
        run_amortization_fill_dry_run(g, report_date_iso="2026-05-22", historical_file_path="HIST/cartera.xlsx")
    )
    assert result["abono_groups_total"] == 1
    assert result["abono_groups_reconciled"] == 1
    assert result["abono_groups_ready"] == 1
    assert result["abono_groups_schedule_rule_missing"] == 0
    assert result["can_apply"] is True
    assert result["requires_business_rule"] is False
    gr = result["abono_group_results"][0]
    assert gr["reconciliation_status"] == "PASSED"
    assert gr["schedule_resolution_status"] == "NOT_REQUIRED"
    assert gr["group_ready_for_apply"] is True
    abono_items = [i for i in result["items"] if i.get("tipo_aplicacion") == "ABONO"]
    assert abono_items
    assert abono_items[0]["application_status"] == "WOULD_APPLY"
    assert abono_items[0]["application_row"] is not None
    assert abono_items[0]["due_date_row"] is None
    assert abono_items[0]["fecha_limite_pago"] is None
    assert abono_items[0]["updates_ibr"] is False
    assert abono_items[0]["ibr_skipped_reason"] == "NOT_REQUIRED_FOR_ABONO"
    assert abono_items[0]["ibr"]["status"] == "NOT_REQUIRED"
    assert not any(i.get("error_code") == "FECHA_LIMITE_NOT_FOUND" for i in abono_items)


def test_dry_run_legacy_pago_unchanged(monkeypatch):
    fecha = date(2026, 5, 22)
    pago = {
        "id_pago": "7785e37e",
        "cliente": "EQUINORTE",
        "credito": "CREDITO # 258",
        "asiento_pdf_path": "clientes/EQUINORTE/CREDITO # 258/ASIENTOS CONTABLES CRED 258/asiento.pdf",
        "extracto_pdf_path": "clientes/EQUINORTE/extracto.pdf",
        "output_relative_path": "OUT/pago.pdf",
    }
    hist = _hist_bytes("7785e37e", "CREDITO # 258", "TABLAS/amort.xlsx", fecha)
    files = _base_files(
        hist=hist,
        amort=_amort_table_bytes(fecha),
        asiento_pdf=_asiento_pdf(),
        ibr=_ibr_bytes(),
        fecha=fecha,
    )
    g = MockGraphDryRun(files)
    monkeypatch.setattr(
        "app.application.use_cases.amortization_fill_dry_run.extract_text_from_pdf",
        lambda _b: _accounting_text(),
    )
    result = asyncio.run(run_amortization_fill_dry_run(g, report_date_iso="2026-05-22"))
    assert result["payment_groups_total"] == 1
    assert result["abono_groups_total"] == 0
    pago_items = [i for i in result["items"] if i.get("tipo_aplicacion") != "ABONO"]
    assert pago_items[0]["application_status"] == "WOULD_APPLY"


def test_dry_run_abono_missing_asiento_blocks_group(monkeypatch):
    out = _abono_manifest_output()
    out["credit_items"][0]["asiento_pdf_paths"] = []
    files = _abono_dry_run_files(abono_output=out)
    g = MockGraphDryRun(files)
    result = asyncio.run(
        run_amortization_fill_dry_run(g, report_date_iso="2026-05-22", historical_file_path="HIST/cartera.xlsx")
    )
    assert result["abono_groups_not_reconciled"] == 1
    assert result["can_apply"] is False
    assert result["abono_groups_missing_accounting_pdf"] >= 1


def test_dry_run_pago_and_abono_independent(monkeypatch):
    fecha = date(2026, 5, 22)
    pago = {
        "id_pago": "P1",
        "cliente": "EQUINORTE",
        "credito": "CREDITO # 258",
        "tipo_aplicacion": "PAGO",
        "requiere_extracto": True,
        "asiento_pdf_path": "clientes/EQUINORTE/CREDITO # 258/ASIENTOS CONTABLES CRED 258/asiento.pdf",
        "extracto_pdf_path": "x.pdf",
        "output_relative_path": "OUT/pago.pdf",
    }
    abono = _abono_manifest_output(id_pago="AB9", monto_banco=50_000_000.0)
    abono["credit_items"][0]["asiento_pdf_paths"] = []
    files = _abono_dry_run_files(abono_output=abono, pago_manifest_output=pago, fecha=fecha)
    g = MockGraphDryRun(files)
    monkeypatch.setattr(
        "app.application.use_cases.amortization_fill_dry_run.extract_text_from_pdf",
        lambda _b: _accounting_text(),
    )
    monkeypatch.setattr(
        "app.application.services.abono_dry_run.extract_text_from_pdf",
        lambda _b: _accounting_text(),
    )
    result = asyncio.run(run_amortization_fill_dry_run(g, report_date_iso=fecha.isoformat()))
    pago_items = [i for i in result["items"] if i.get("tipo_aplicacion") != "ABONO"]
    assert pago_items[0]["application_status"] == "WOULD_APPLY"
    assert result["abono_groups_not_reconciled"] == 1
    assert result["can_apply"] is False
