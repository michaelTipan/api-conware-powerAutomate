"""Merge incompleto: no PDF parcial, manifest PARTIAL, gates Dry-run/Apply."""

import asyncio
import json
from datetime import date
from unittest.mock import AsyncMock, patch

import pytest

from app.application.services.merge_group_validation import (
    MANIFEST_STATUS_PARTIAL,
    expected_creditos_for_id_pago,
    validate_merge_group_completeness,
)
from app.application.services.merge_manifest_gate import (
    MERGE_INCOMPLETE_NOT_APPLICABLE,
    evaluate_merge_incomplete_block,
)
from app.application.use_cases.amortization_fill_apply import run_amortization_fill_apply
from app.application.use_cases.amortization_fill_dry_run import run_amortization_fill_dry_run
from app.application.use_cases.merge_composite_validado_pdfs import merge_composite_validado_pdfs
from tests.test_merge_composite_control_workbook import (
    _MergeGraph,
    _bank_bytes,
    _hist_workbook_bytes,
    _tiny_pdf,
)
from app.application.sharepoint_resolution import encode_graph_drive_path


@pytest.fixture(autouse=True)
def env_merge(monkeypatch):
    import app.application.sharepoint_resolution as spr
    import app.application.use_cases.payment_validation_process_control as pc
    import app.application.use_cases.merge_composite_validado_pdfs as m

    monkeypatch.setenv("GRAPH_MERGE_COMPOSITE_OUTPUT_FOLDER_PATH", "OUT/PDFS")
    monkeypatch.setenv("GRAPH_SHAREPOINT_SITE_SEARCH", "TEST_SITE")
    monkeypatch.setenv("GRAPH_SHAREPOINT_DRIVE_NAME", "TEST_DRIVE")
    monkeypatch.setenv("GRAPH_SHAREPOINT_FILE_PATH", "bank/report.xlsx")
    monkeypatch.setenv("GRAPH_PAYMENT_VALIDATION_CONTROL_PATH", "CTL")
    monkeypatch.setenv("GRAPH_PAYMENT_VALIDATION_LOGS_PATH", "LOGS")

    async def _fake_read(_g, _s, _d, *, bank_code: str):
        from types import SimpleNamespace

        return SimpleNamespace(
            estado_proceso="PENDIENTE_ASIENTOS",
            is_active=True,
            process_key=f"payment-validation|{bank_code}|2026-05-12",
            historical_file_path="",
            email_pdf_path="",
            merge_manifest_path="",
            merge_idempotency_key="",
        )

    async def _fake_update(_g, _s, _d, *, bank_code: str, updates: dict):
        return True

    async def _fake_resolve(_g, _site_search: str, _drive_name: str, _path: str):
        return {
            "site_id": "s1",
            "drive_id": "d1",
            "path_encoded": encode_graph_drive_path("bank/report.xlsx"),
            "file_path": "bank/report.xlsx",
        }

    monkeypatch.setattr(pc, "read_process_control_snapshot", _fake_read)
    monkeypatch.setattr(pc, "update_process_control_row2", _fake_update)
    monkeypatch.setattr(spr, "resolve_sharepoint_path", _fake_resolve)
    monkeypatch.setattr(m, "resolve_sharepoint_path", _fake_resolve)


def test_expected_creditos_from_rows():
    rows = [
        {"credito_digits": "265", "excel_row": 2},
        {"credito_digits": "258", "excel_row": 3},
    ]
    assert expected_creditos_for_id_pago(rows) == ("258", "265")


def test_validate_group_incomplete_when_missing_credit():
    rows = [{"credito_digits": "258"}, {"credito_digits": "265"}]
    credit_items = [{"credito": "265", "asiento_pdf_paths": ["a.pdf"], "extracto_pdf_paths": ["e.pdf"]}]
    result = validate_merge_group_completeness(
        id_pago="G1",
        tipo_aplicacion="PAGO",
        group_rows=rows,
        credit_items=credit_items,
        pre_skips=["reason=asiento_contable_not_found | credit_number_expected=258"],
    )
    assert not result.is_complete
    assert result.missing_creditos == ("258",)


def test_manifest_gate_blocks_partial():
    manifest = {
        "manifest_status": "PARTIAL",
        "eligible_for_dry_run": False,
        "incomplete_groups_count": 1,
        "outputs": [],
        "incomplete_groups": [{"id_pago": "G1", "missing_creditos": ["258"]}],
        "skipped": ["x"],
    }
    block = evaluate_merge_incomplete_block(manifest, estado_proceso="MERGE_PARCIAL")
    assert block is not None
    assert block["error_code"] in (
        MERGE_INCOMPLETE_NOT_APPLICABLE,
        "MERGE_GROUP_PENDING_INPUTS",
    )


def test_dry_run_blocked_on_partial_manifest(monkeypatch):
    fecha = date(2026, 5, 22)
    manifest = {
        "report_date_iso": fecha.isoformat(),
        "manifest_status": "PARTIAL",
        "eligible_for_dry_run": False,
        "incomplete_groups_count": 1,
        "outputs": [],
        "incomplete_groups": [{"id_pago": "8489b9e6", "missing_creditos": ["258"]}],
        "skipped": ["asiento missing"],
    }
    manifest_key = f"LOGS/merge_manifest_{fecha.isoformat()}.json"

    class G:
        files = {manifest_key: json.dumps(manifest).encode("utf-8"), "CTL/dummy.xlsx": b"x"}

        async def get(self, endpoint, params=None):
            return {"value": []}

        async def get_bytes(self, endpoint, params=None):
            from urllib.parse import unquote

            path = unquote(endpoint.split("/root:/", 1)[1].rsplit(":/content", 1)[0])
            return self.files[path]

    monkeypatch.setenv("GRAPH_IBR_DIARIO_PATH", "CTL/IBR_DIARIO.xlsx")

    async def run():
        with patch(
            "app.application.use_cases.amortization_fill_dry_run.resolve_sharepoint_path",
            new_callable=AsyncMock,
            return_value={"site_id": "s", "drive_id": "d"},
        ):
            return await run_amortization_fill_dry_run(
                G(),
                report_date_iso=fecha.isoformat(),
                merge_manifest_path=manifest_key,
                bank_code="banco_bogota",
            )

    out = asyncio.run(run())
    assert out["status"] == "blocked"
    assert out["can_apply"] is False
    assert out["items"] == []


def test_merge_two_credits_one_missing_no_pdf_upload(monkeypatch):
    hist = "HIST/hist.xlsx"
    email = "EMAIL/mail.pdf"
    p258 = "clientes/GEO/CREDITO # 258/Extracto 258.pdf"
    p265 = "clientes/GEO/CREDITO # 265/Extracto 265.pdf"
    d265 = "clientes/GEO/CREDITO # 265/ASIENTOS CONTABLES CRED 265"
    a265 = f"{d265}/Asiento 265.pdf"

    g = _MergeGraph()
    g.initial["bank/report.xlsx"] = _bank_bytes()
    g.initial[hist] = _hist_workbook_bytes(
        [
            ["VALIDAR", "", "8489b9e6", "GEO", "258", "clientes/GEO/CREDITO # 258/ASIENTOS"],
            ["VALIDAR", "", "8489b9e6", "GEO", "265", d265],
        ]
    )
    g.initial[email] = _tiny_pdf()
    g.initial[p258] = _tiny_pdf()
    g.initial[p265] = _tiny_pdf()
    g.initial[a265] = _tiny_pdf()
    g.children[d265] = [{"name": "Asiento 265.pdf", "file": {}}]

    ctx = {
        "site_id": "s1",
        "drive_id": "d1",
        "path_encoded": encode_graph_drive_path("bank/report.xlsx"),
        "file_path": "bank/report.xlsx",
    }
    ret = iter([[p258], [p265]])

    async def fake_collect(_gr, _si, _dr, _cell):
        return next(ret)

    async def run():
        with (
            patch(
                "app.application.use_cases.merge_composite_validado_pdfs.resolve_sharepoint_from_env",
                new_callable=AsyncMock,
                return_value=ctx,
            ),
            patch(
                "app.application.use_cases.merge_composite_validado_pdfs._collect_pdf_paths_from_ruta_cell",
                new_callable=AsyncMock,
                side_effect=fake_collect,
            ),
        ):
            return await merge_composite_validado_pdfs(
                g,
                bank_code="banco_bogota",
                historical_file_path=hist,
                email_pdf_path=email,
            )

    r = asyncio.run(run())
    assert r.outputs_count == 0
    assert r.incomplete_groups_count == 1
    assert r.merge_control_status == "MERGE_PARCIAL"
    assert r.manifest_status == MANIFEST_STATUS_PARTIAL
    assert r.payment_skipped_count >= 1
    assert not any(p.startswith("OUT/PDFS/") for p in g.uploaded)
