"""Phase 5B: apply con control oficial por banco."""

from __future__ import annotations

import asyncio
import json
from datetime import date
from unittest.mock import AsyncMock, patch

import pytest

from app.application.use_cases.amortization_fill_apply import (
    VERIFICATION_FAILED,
    run_amortization_fill_apply,
)


class _Graph:
    def __init__(self, files: dict[str, bytes] | None = None) -> None:
        self.files = files or {}

    async def get(self, endpoint: str, params=None):
        return {"value": []}

    async def get_bytes(self, endpoint: str, params=None):
        from urllib.parse import unquote

        import httpx

        path = unquote(endpoint.split("/root:/", 1)[1].rsplit(":/content", 1)[0])
        if path in self.files:
            return self.files[path]
        raise httpx.HTTPStatusError(
            "404",
            request=httpx.Request("GET", endpoint),
            response=httpx.Response(404),
        )


class _Snap:
    def __init__(
        self,
        *,
        estado: str,
        merge_manifest: str = "",
        historical: str = "",
        process_key: str = "",
        apply_idem: str = "",
        is_active: bool = True,
    ) -> None:
        self.estado_proceso = estado
        self.is_active = is_active
        self.merge_manifest_path = merge_manifest
        self.historical_file_path = historical
        self.process_key = process_key
        self.apply_idempotency_key = apply_idem


def _manifest_bytes(fecha: date) -> bytes:
    return json.dumps(
        {
            "report_date_iso": fecha.isoformat(),
            "historico_excel_path": "HIST/h.xlsx",
            "manifest_status": "COMPLETE",
            "eligible_for_dry_run": True,
            "incomplete_groups_count": 0,
            "outputs": [
                {
                    "id_pago": "p1",
                    "status": "COMPLETE",
                    "expected_creditos": ["258"],
                    "creditos_seleccionados": ["258"],
                    "credit_items": [{"credito": "258", "asiento_pdf_paths": ["a.pdf"]}],
                    "output_relative_path": "OUT/x.pdf",
                    "eligible_for_dry_run": True,
                }
            ],
            "skipped": [],
        }
    ).encode("utf-8")


@pytest.fixture(autouse=True)
def _env(monkeypatch):
    monkeypatch.setenv("GRAPH_SHAREPOINT_SITE_SEARCH", "site")
    monkeypatch.setenv("GRAPH_SHAREPOINT_DRIVE_NAME", "drive")
    monkeypatch.setenv("GRAPH_PAYMENT_VALIDATION_LOGS_PATH", "LOGS")


def _patch_apply_success(monkeypatch, *, dry_run_extra: dict | None = None):
    """Evita escritura real en tablas; dry-run vacío OK."""

    async def fake_dry_run(graph, **kwargs):
        base = {
            "status": "ok",
            "mode": "dry_run",
            "manifest_path": kwargs.get("merge_manifest_path") or "LOGS/m.json",
            "historical_file_path": kwargs.get("historical_file_path"),
            "items": [],
            "summary": {"errors": 0, "skipped_idempotent": 0},
        }
        if dry_run_extra:
            base.update(dry_run_extra)
        return base

    monkeypatch.setattr(
        "app.application.use_cases.amortization_fill_apply.run_amortization_fill_dry_run",
        fake_dry_run,
    )
    monkeypatch.setattr(
        "app.application.use_cases.amortization_fill_apply._writable_planned_items",
        lambda _dry: {},
    )


def test_apply_auto_detects_bogota_when_only_bogota_ready(monkeypatch):
    fecha = date(2026, 6, 1)
    manifest_rel = f"LOGS/merge_manifest_banco_bogota_{fecha.isoformat()}.json"
    g = _Graph({manifest_rel: _manifest_bytes(fecha)})
    _patch_apply_success(monkeypatch)
    captured: list[dict] = []

    async def fake_update(_g, _s, _d, *, bank_code: str, updates: dict):
        captured.append(dict(updates))
        return True

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_dry_run.resolve_sharepoint_path",
                new_callable=AsyncMock,
                return_value={"site_id": "s", "drive_id": "d"},
            ),
            patch(
                "app.application.use_cases.amortization_fill_dry_run.read_process_control_snapshot",
                new_callable=AsyncMock,
                side_effect=lambda _g, _s, _d, *, bank_code: _Snap(
                    estado="CONSOLIDADO" if bank_code == "banco_bogota" else "FINALIZADO",
                    merge_manifest=manifest_rel if bank_code == "banco_bogota" else "",
                    historical="HIST/h.xlsx" if bank_code == "banco_bogota" else "",
                    process_key=f"payment-validation|{bank_code}|{fecha.isoformat()}",
                ),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.update_process_control_row2",
                side_effect=fake_update,
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.read_process_control_snapshot",
                new_callable=AsyncMock,
                side_effect=lambda _g, _s, _d, *, bank_code: _Snap(
                    estado="CONSOLIDADO",
                    merge_manifest=manifest_rel,
                    historical="HIST/h.xlsx",
                    process_key=f"payment-validation|banco_bogota|{fecha.isoformat()}",
                ),
            ),
        ):
            return await run_amortization_fill_apply(g)

    out = asyncio.run(run())
    assert out["bank_code"] == "banco_bogota"
    assert out["bank_code_source"] == "auto_detected"
    assert out["merge_manifest_source"] == "control"
    assert out["historical_file_source"] == "control"
    assert any(u.get("EstadoProceso") == "APLICANDO_AMORTIZACION" for u in captured)
    assert any(u.get("EstadoProceso") == "AMORTIZACION_APLICADA" for u in captured)


def test_apply_auto_detects_bancolombia_when_only_bancolombia_ready(monkeypatch):
    fecha = date(2026, 6, 1)
    manifest_rel = f"LOGS/merge_manifest_banco_bancolombia_{fecha.isoformat()}.json"
    g = _Graph({manifest_rel: _manifest_bytes(fecha)})
    _patch_apply_success(monkeypatch)

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_dry_run.resolve_sharepoint_path",
                new_callable=AsyncMock,
                return_value={"site_id": "s", "drive_id": "d"},
            ),
            patch(
                "app.application.use_cases.amortization_fill_dry_run.read_process_control_snapshot",
                new_callable=AsyncMock,
                side_effect=lambda _g, _s, _d, *, bank_code: _Snap(
                    estado="CONSOLIDADO" if bank_code == "banco_bancolombia" else "PENDIENTE_ASIENTOS",
                    merge_manifest=manifest_rel if bank_code == "banco_bancolombia" else "",
                    historical="HIST/h.xlsx" if bank_code == "banco_bancolombia" else "",
                    process_key=f"payment-validation|{bank_code}|{fecha.isoformat()}",
                ),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.read_process_control_snapshot",
                new_callable=AsyncMock,
                return_value=_Snap(estado="CONSOLIDADO", merge_manifest=manifest_rel, historical="HIST/h.xlsx"),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.update_process_control_row2",
                new_callable=AsyncMock,
                return_value=True,
            ),
        ):
            return await run_amortization_fill_apply(g)

    out = asyncio.run(run())
    assert out["bank_code"] == "banco_bancolombia"
    assert out["bank_code_source"] == "auto_detected"


def test_apply_fails_when_no_ready_process(monkeypatch):
    g = _Graph({})

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_dry_run.resolve_sharepoint_path",
                new_callable=AsyncMock,
                return_value={"site_id": "s", "drive_id": "d"},
            ),
            patch(
                "app.application.use_cases.amortization_fill_dry_run.read_process_control_snapshot",
                new_callable=AsyncMock,
                return_value=_Snap(estado="FINALIZADO"),
            ),
        ):
            with pytest.raises(ValueError, match="NO_READY_PROCESS"):
                await run_amortization_fill_apply(g)

    asyncio.run(run())


def test_apply_fails_when_multiple_ready_processes(monkeypatch):
    g = _Graph({})

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_dry_run.resolve_sharepoint_path",
                new_callable=AsyncMock,
                return_value={"site_id": "s", "drive_id": "d"},
            ),
            patch(
                "app.application.use_cases.amortization_fill_dry_run.read_process_control_snapshot",
                new_callable=AsyncMock,
                return_value=_Snap(
                    estado="CONSOLIDADO",
                    merge_manifest="LOGS/x.json",
                    historical="HIST/h.xlsx",
                ),
            ),
        ):
            with pytest.raises(ValueError, match="MULTIPLE_READY_PROCESSES"):
                await run_amortization_fill_apply(g)

    asyncio.run(run())


def test_apply_with_bank_code_override(monkeypatch):
    g = _Graph({})
    _patch_apply_success(monkeypatch)

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_apply._drive_context",
                new_callable=AsyncMock,
                return_value=("s", "d"),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply._resolve_amortization_inputs",
                new_callable=AsyncMock,
                return_value=(
                    "LOGS/m.json",
                    None,
                    "HIST/h.xlsx",
                    "banco_bancolombia",
                    "Bancolombia",
                    "pk-bc",
                    "CTL/bancolombia.xlsx",
                    [],
                    "control",
                    "control",
                ),
            ) as mock_resolve,
            patch(
                "app.application.use_cases.amortization_fill_apply.read_process_control_snapshot",
                new_callable=AsyncMock,
                return_value=_Snap(estado="CONSOLIDADO", merge_manifest="LOGS/m.json", historical="HIST/h.xlsx"),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.update_process_control_row2",
                new_callable=AsyncMock,
                return_value=True,
            ),
        ):
            out = await run_amortization_fill_apply(g, bank_code="banco_bancolombia")
            mock_resolve.assert_awaited_once()
            assert mock_resolve.await_args.kwargs["bank_code"] == "banco_bancolombia"
            return out

    out = asyncio.run(run())
    assert out["bank_code"] == "banco_bancolombia"
    assert out["bank_code_source"] == "body"


def test_apply_uses_control_paths_when_body_omits_manifest_and_hist(monkeypatch):
    g = _Graph({})
    _patch_apply_success(monkeypatch)

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_apply._drive_context",
                new_callable=AsyncMock,
                return_value=("s", "d"),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply._resolve_amortization_inputs",
                new_callable=AsyncMock,
                return_value=(
                    "LOGS/from_control.json",
                    None,
                    "HIST/from_control.xlsx",
                    "banco_bogota",
                    "Banco de Bogotá",
                    "pk",
                    "CTL/bogota.xlsx",
                    [],
                    "control",
                    "control",
                ),
            ) as mock_resolve,
            patch(
                "app.application.use_cases.amortization_fill_apply.read_process_control_snapshot",
                new_callable=AsyncMock,
                return_value=_Snap(estado="CONSOLIDADO"),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.update_process_control_row2",
                new_callable=AsyncMock,
                return_value=True,
            ),
        ):
            out = await run_amortization_fill_apply(g)
            mock_resolve.assert_awaited_once()
            assert mock_resolve.await_args.kwargs["merge_manifest_path"] is None
            assert mock_resolve.await_args.kwargs["historical_file_path"] is None
            return out

    out = asyncio.run(run())
    assert out["merge_manifest_source"] == "control"
    assert out["historical_file_source"] == "control"
    assert out["merge_manifest_path"] == "LOGS/from_control.json"
    assert out["historical_file_path"] == "HIST/from_control.xlsx"


def test_apply_keeps_manual_overrides(monkeypatch):
    g = _Graph({})
    _patch_apply_success(monkeypatch)

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_apply._drive_context",
                new_callable=AsyncMock,
                return_value=("s", "d"),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply._resolve_amortization_inputs",
                new_callable=AsyncMock,
                return_value=(
                    "LOGS/manual.json",
                    "2026-06-01",
                    "HIST/manual.xlsx",
                    "banco_bogota",
                    "Banco de Bogotá",
                    "",
                    "",
                    [],
                    "body",
                    "body",
                ),
            ),
        ):
            return await run_amortization_fill_apply(
                g,
                merge_manifest_path="LOGS/manual.json",
                historical_file_path="HIST/manual.xlsx",
                report_date_iso="2026-06-01",
            )

    out = asyncio.run(run())
    assert out["merge_manifest_source"] == "body"
    assert out["historical_file_source"] == "body"


def test_apply_already_applied_skips_writes():
    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_apply._drive_context",
                new_callable=AsyncMock,
                return_value=("s", "d"),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply._resolve_amortization_inputs",
                new_callable=AsyncMock,
                return_value=(
                    "LOGS/manifest.json",
                    None,
                    "HIST/h.xlsx",
                    "banco_bogota",
                    "Banco de Bogotá",
                    "payment-validation|banco_bogota|2026-06-01",
                    "CTL/bogota.xlsx",
                    [],
                    "control",
                    "control",
                ),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.read_process_control_snapshot",
                new_callable=AsyncMock,
                return_value=_Snap(
                    estado="AMORTIZACION_APLICADA",
                    merge_manifest="LOGS/manifest.json",
                    historical="HIST/h.xlsx",
                    apply_idem="payment-validation|banco_bogota|2026-06-01",
                    process_key="payment-validation|banco_bogota|2026-06-01",
                ),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.run_amortization_fill_dry_run",
                new_callable=AsyncMock,
            ) as mock_dry,
        ):
            out = await run_amortization_fill_apply(_Graph())  # type: ignore[arg-type]
            mock_dry.assert_not_called()
            assert out["already_applied"] is True
            assert out["file_action"] == "reused"
            assert out["apply_wrote_changes"] is False
            assert out["process_control_estado"] == "AMORTIZACION_APLICADA"

    asyncio.run(run())


def test_apply_sets_amortizacion_parcial_on_partial_status(monkeypatch):
    g = _Graph({})
    captured: list[dict] = []

    async def fake_dry_run(graph, **kwargs):
        return {
            "status": "ok",
            "mode": "dry_run",
            "manifest_path": "LOGS/m.json",
            "items": [],
            "summary": {"errors": 0},
        }

    monkeypatch.setattr(
        "app.application.use_cases.amortization_fill_apply.run_amortization_fill_dry_run",
        fake_dry_run,
    )
    monkeypatch.setattr(
        "app.application.use_cases.amortization_fill_apply._writable_planned_items",
        lambda _dry: {
            "T/ok.xlsx": [{"tabla_amortizacion_path": "T/ok.xlsx"}],
            "T/fail.xlsx": [{"tabla_amortizacion_path": "T/fail.xlsx"}],
        },
    )

    async def fake_apply_one_table(_graph, _s, _d, tabla_path, *_a, **_k):
        if "ok" in tabla_path:
            return {
                "items": [{"apply_status": "applied"}],
                "uploaded": True,
                "upload_status": "uploaded",
                "verification_status": "ok",
            }
        return {
            "items": [],
            "uploaded": False,
            "upload_status": "failed",
            "verification_status": VERIFICATION_FAILED,
        }

    async def fake_update(_g, _s, _d, *, bank_code: str, updates: dict):
        captured.append(dict(updates))
        return True

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_apply._drive_context",
                new_callable=AsyncMock,
                return_value=("s", "d"),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply._resolve_amortization_inputs",
                new_callable=AsyncMock,
                return_value=(
                    "LOGS/m.json",
                    None,
                    "HIST/h.xlsx",
                    "banco_bogota",
                    "Banco de Bogotá",
                    "pk",
                    "CTL/b.xlsx",
                    [],
                    "control",
                    "control",
                ),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.read_process_control_snapshot",
                new_callable=AsyncMock,
                return_value=_Snap(estado="CONSOLIDADO"),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.update_process_control_row2",
                side_effect=fake_update,
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply._apply_one_table",
                side_effect=fake_apply_one_table,
            ),
        ):
            return await run_amortization_fill_apply(g)

    out = asyncio.run(run())
    assert out["status"] == "partial"
    assert out["process_control_estado"] == "AMORTIZACION_PARCIAL"
    assert any(u.get("EstadoProceso") == "AMORTIZACION_PARCIAL" for u in captured)
    parcial = [u for u in captured if u.get("EstadoProceso") == "AMORTIZACION_PARCIAL"][-1]
    assert parcial["LastStepStatus"] == "COMPLETED_WITH_WARNINGS"


def test_apply_sets_amortizacion_aplicada_on_ok(monkeypatch):
    g = _Graph({})
    captured: list[dict] = []

    async def fake_dry_run(graph, **kwargs):
        return {
            "status": "ok",
            "mode": "dry_run",
            "manifest_path": "LOGS/m.json",
            "items": [],
            "summary": {"errors": 0, "skipped_idempotent": 2},
        }

    monkeypatch.setattr(
        "app.application.use_cases.amortization_fill_apply.run_amortization_fill_dry_run",
        fake_dry_run,
    )
    monkeypatch.setattr(
        "app.application.use_cases.amortization_fill_apply._writable_planned_items",
        lambda _dry: {"T/x.xlsx": [{"credit_id": "c1"}]},
    )

    async def fake_apply_one_table(*_a, **_k):
        return {
            "items": [{"apply_status": "applied"}],
            "uploaded": True,
            "upload_status": "uploaded",
            "verification_status": "ok",
        }

    async def fake_update(_g, _s, _d, *, bank_code: str, updates: dict):
        captured.append(dict(updates))
        return True

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_apply._drive_context",
                new_callable=AsyncMock,
                return_value=("s", "d"),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply._resolve_amortization_inputs",
                new_callable=AsyncMock,
                return_value=(
                    "LOGS/m.json",
                    None,
                    "HIST/h.xlsx",
                    "banco_bogota",
                    "Banco de Bogotá",
                    "pk-ok",
                    "CTL/b.xlsx",
                    [],
                    "control",
                    "control",
                ),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.read_process_control_snapshot",
                new_callable=AsyncMock,
                return_value=_Snap(estado="CONSOLIDADO"),
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply.update_process_control_row2",
                side_effect=fake_update,
            ),
            patch(
                "app.application.use_cases.amortization_fill_apply._apply_one_table",
                side_effect=fake_apply_one_table,
            ),
        ):
            return await run_amortization_fill_apply(g, job_id="job-apply-1")

    out = asyncio.run(run())
    assert out["status"] == "ok"
    assert out["process_control_estado"] == "AMORTIZACION_APLICADA"
    assert out["apply_wrote_changes"] is True
    final = [u for u in captured if u.get("EstadoProceso") == "AMORTIZACION_APLICADA"]
    assert final
    assert final[-1]["ApplyIdempotencyKey"] == "pk-ok"
    assert final[-1]["ApplyJobId"] == "job-apply-1"
    assert final[-1]["LastCompletedStep"] == "APPLY"
