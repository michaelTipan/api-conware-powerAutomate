"""Amortización sin body: lee MergeManifestPath desde el control oficial por banco."""

import asyncio
import json
from datetime import date
from io import BytesIO
from unittest.mock import AsyncMock, patch

import openpyxl
import pytest
from openpyxl import Workbook

from app.application.use_cases.amortization_fill_dry_run import run_amortization_fill_dry_run


def _manifest_bytes(fecha: date) -> bytes:
    payload = {
        "report_date_iso": fecha.isoformat(),
        "historico_excel_path": "HIST/h.xlsx",
        "outputs": [],
    }
    return json.dumps(payload).encode("utf-8")


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


def test_amortization_dry_run_resolves_manifest_from_merge_control(monkeypatch):
    monkeypatch.setenv("GRAPH_SHAREPOINT_SITE_SEARCH", "site")
    monkeypatch.setenv("GRAPH_SHAREPOINT_DRIVE_NAME", "drive")
    monkeypatch.setenv("GRAPH_PAYMENT_VALIDATION_LOGS_PATH", "LOGS")
    fecha = date(2026, 5, 20)
    manifest_rel = f"LOGS/merge_manifest_banco_bogota_{fecha.isoformat()}.json"
    g = _Graph(
        {
            manifest_rel: _manifest_bytes(fecha),
        }
    )
    ctx = {"site_id": "s", "drive_id": "d"}

    async def run():
        with patch(
            "app.application.use_cases.amortization_fill_dry_run.resolve_sharepoint_path",
            new_callable=AsyncMock,
            return_value=ctx,
        ):
            with patch(
                "app.application.use_cases.amortization_fill_dry_run.read_process_control_snapshot",
                new_callable=AsyncMock,
                    side_effect=lambda _g, _s, _d, *, bank_code: type(
                        "Snap",
                        (),
                        {
                            "estado_proceso": "CONSOLIDADO"
                            if bank_code == "banco_bogota"
                            else "FINALIZADO",
                            "is_active": True,
                            "merge_manifest_path": manifest_rel
                            if bank_code == "banco_bogota"
                            else "",
                            "historical_file_path": "HIST/h.xlsx"
                            if bank_code == "banco_bogota"
                            else "",
                            "process_key": f"payment-validation|{bank_code}|2026-05-20",
                        },
                    )(),
            ):
                with patch(
                    "app.application.use_cases.amortization_fill_dry_run.update_process_control_row2",
                    new_callable=AsyncMock,
                    return_value=True,
                ):
                    return await run_amortization_fill_dry_run(g)

    out = asyncio.run(run())
    assert out["merge_manifest_source"] == "control"
    assert out["manifest_path"] == manifest_rel


def test_amortization_fails_when_control_not_ready(monkeypatch):
    monkeypatch.setenv("GRAPH_SHAREPOINT_SITE_SEARCH", "site")
    monkeypatch.setenv("GRAPH_SHAREPOINT_DRIVE_NAME", "drive")
    g = _Graph(
        {
        }
    )
    ctx = {"site_id": "s", "drive_id": "d"}

    async def run():
        with patch(
            "app.application.use_cases.amortization_fill_dry_run.resolve_sharepoint_path",
            new_callable=AsyncMock,
            return_value=ctx,
        ):
            with patch(
                "app.application.use_cases.amortization_fill_dry_run.read_process_control_snapshot",
                new_callable=AsyncMock,
                return_value=type(
                    "Snap",
                    (),
                    {
                        "estado_proceso": "PENDIENTE_ASIENTOS",
                        "is_active": True,
                        "merge_manifest_path": "",
                        "historical_file_path": "HIST/h.xlsx",
                        "process_key": "payment-validation|banco_bogota|2026-05-20",
                    },
                )(),
            ):
                with pytest.raises(ValueError, match="NO_READY_PROCESS"):
                    await run_amortization_fill_dry_run(g)

    asyncio.run(run())
