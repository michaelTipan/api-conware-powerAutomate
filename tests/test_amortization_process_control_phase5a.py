from __future__ import annotations

import asyncio
import json
from datetime import date
from unittest.mock import AsyncMock, patch

import pytest

from app.application.use_cases.amortization_fill_dry_run import run_amortization_fill_dry_run


class _Graph:
    def __init__(self, files: dict[str, bytes]) -> None:
        self.files = files

    async def get(self, endpoint: str, params=None):
        return {"value": []}

    async def get_bytes(self, endpoint: str, params=None):
        from urllib.parse import unquote

        path = unquote(endpoint.split("/root:/", 1)[1].rsplit(":/content", 1)[0])
        if path in self.files:
            return self.files[path]
        import httpx

        raise httpx.HTTPStatusError(
            "404",
            request=httpx.Request("GET", endpoint),
            response=httpx.Response(404),
        )


def _manifest_bytes(fecha: date) -> bytes:
    payload = {
        "report_date_iso": fecha.isoformat(),
        "historico_excel_path": "HIST/h.xlsx",
        "manifest_status": "COMPLETE",
        "eligible_for_dry_run": True,
        "incomplete_groups_count": 0,
        "outputs": [
            {
                "id_pago": "p1",
                "status": "COMPLETE",
                "cliente": "C",
                "credito": "258",
                "tipo_aplicacion": "PAGO",
                "expected_creditos": ["258"],
                "creditos_seleccionados": ["258"],
                "asiento_pdf_path": "a.pdf",
                "extracto_pdf_path": "e.pdf",
                "output_relative_path": "OUT/x.pdf",
                "eligible_for_dry_run": True,
                "credit_items": [
                    {
                        "credito": "258",
                        "asiento_pdf_paths": ["a.pdf"],
                        "extracto_pdf_paths": ["e.pdf"],
                    }
                ],
            }
        ],
        "skipped": [],
    }
    return json.dumps(payload).encode("utf-8")


@pytest.fixture(autouse=True)
def _env(monkeypatch):
    monkeypatch.setenv("GRAPH_SHAREPOINT_SITE_SEARCH", "site")
    monkeypatch.setenv("GRAPH_SHAREPOINT_DRIVE_NAME", "drive")
    monkeypatch.setenv("GRAPH_PAYMENT_VALIDATION_LOGS_PATH", "LOGS")


def test_dry_run_auto_detects_bogota_when_only_bogota_ready(monkeypatch):
    fecha = date(2026, 6, 1)
    manifest_rel = f"LOGS/merge_manifest_banco_bogota_{fecha.isoformat()}.json"
    g = _Graph({manifest_rel: _manifest_bytes(fecha)})
    ctx = {"site_id": "s", "drive_id": "d"}

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_dry_run.resolve_sharepoint_path",
                new_callable=AsyncMock,
                return_value=ctx,
            ),
            patch(
                "app.application.use_cases.amortization_fill_dry_run.read_process_control_snapshot",
                new_callable=AsyncMock,
                side_effect=lambda _g, _s, _d, *, bank_code: type(
                    "Snap",
                    (),
                    {
                        "estado_proceso": "CONSOLIDADO" if bank_code == "banco_bogota" else "FINALIZADO",
                        "is_active": True,
                        "merge_manifest_path": manifest_rel if bank_code == "banco_bogota" else "",
                        "historical_file_path": "HIST/h.xlsx" if bank_code == "banco_bogota" else "",
                        "process_key": f"payment-validation|{bank_code}|{fecha.isoformat()}",
                    },
                )(),
            ),
            patch(
                "app.application.use_cases.amortization_fill_dry_run.update_process_control_row2",
                new_callable=AsyncMock,
                return_value=True,
            ),
        ):
            return await run_amortization_fill_dry_run(g)

    out = asyncio.run(run())
    assert out["bank_code"] == "banco_bogota"
    assert out["bank_code_source"] == "auto_detected"
    assert out["merge_manifest_source"] == "control"
    assert out["dry_run_wrote_changes"] is False


def test_dry_run_fails_when_no_ready_process(monkeypatch):
    fecha = date(2026, 6, 1)
    g = _Graph({})
    ctx = {"site_id": "s", "drive_id": "d"}

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_dry_run.resolve_sharepoint_path",
                new_callable=AsyncMock,
                return_value=ctx,
            ),
            patch(
                "app.application.use_cases.amortization_fill_dry_run.read_process_control_snapshot",
                new_callable=AsyncMock,
                return_value=type(
                    "Snap",
                    (),
                    {
                        "estado_proceso": "FINALIZADO",
                        "is_active": True,
                        "merge_manifest_path": "",
                        "historical_file_path": "",
                        "process_key": "",
                    },
                )(),
            ),
        ):
            with pytest.raises(ValueError, match="NO_READY_PROCESS"):
                await run_amortization_fill_dry_run(g)

    asyncio.run(run())


def test_dry_run_fails_when_multiple_ready_processes(monkeypatch):
    fecha = date(2026, 6, 1)
    g = _Graph({})
    ctx = {"site_id": "s", "drive_id": "d"}

    async def run():
        with (
            patch(
                "app.application.use_cases.amortization_fill_dry_run.resolve_sharepoint_path",
                new_callable=AsyncMock,
                return_value=ctx,
            ),
            patch(
                "app.application.use_cases.amortization_fill_dry_run.read_process_control_snapshot",
                new_callable=AsyncMock,
                return_value=type(
                    "Snap",
                    (),
                    {
                        "estado_proceso": "CONSOLIDADO",
                        "is_active": True,
                        "merge_manifest_path": "LOGS/x.json",
                        "historical_file_path": "HIST/h.xlsx",
                        "process_key": "k",
                    },
                )(),
            ),
        ):
            with pytest.raises(ValueError, match="MULTIPLE_READY_PROCESSES"):
                await run_amortization_fill_dry_run(g)

    asyncio.run(run())

