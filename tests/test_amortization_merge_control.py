"""Amortización sin body: lee MergeManifestPath desde control_merge_pdfs.xlsx."""

import asyncio
import json
from datetime import date
from io import BytesIO
from unittest.mock import AsyncMock, patch

import openpyxl
import pytest
from openpyxl import Workbook

from app.application.use_cases.amortization_fill_dry_run import run_amortization_fill_dry_run
from app.application.use_cases.setup_merge_control_workbook import MERGE_CONTROL_COLUMNS, SHEET_NAME


def _control_bytes(*, estado: str, manifest: str, hist: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(list(MERGE_CONTROL_COLUMNS))
    row = {c: "" for c in MERGE_CONTROL_COLUMNS}
    row["EstadoProceso"] = estado
    row["IsActive"] = False
    row["HistoricalFilePath"] = hist
    row["MergeManifestPath"] = manifest
    ws.append([row[c] for c in MERGE_CONTROL_COLUMNS])
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


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
    monkeypatch.setenv("GRAPH_MERGE_CONTROL_WORKBOOK_PATH", "CTL/control.xlsx")
    monkeypatch.setenv("GRAPH_PAYMENT_VALIDATION_LOGS_PATH", "LOGS")
    fecha = date(2026, 5, 20)
    manifest_rel = f"LOGS/merge_manifest_{fecha.isoformat()}.json"
    g = _Graph(
        {
            "CTL/control.xlsx": _control_bytes(
                estado="CONSOLIDADO",
                manifest=manifest_rel,
                hist="HIST/h.xlsx",
            ),
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
            return await run_amortization_fill_dry_run(g)

    out = asyncio.run(run())
    assert out["resolved_from_merge_control"] is True
    assert out["manifest_path"] == manifest_rel


def test_amortization_fails_when_control_not_ready(monkeypatch):
    monkeypatch.setenv("GRAPH_SHAREPOINT_SITE_SEARCH", "site")
    monkeypatch.setenv("GRAPH_SHAREPOINT_DRIVE_NAME", "drive")
    monkeypatch.setenv("GRAPH_MERGE_CONTROL_WORKBOOK_PATH", "CTL/control.xlsx")
    g = _Graph(
        {
            "CTL/control.xlsx": _control_bytes(
                estado="PENDIENTE_ASIENTOS",
                manifest="",
                hist="HIST/h.xlsx",
            ),
        }
    )
    ctx = {"site_id": "s", "drive_id": "d"}

    async def run():
        with patch(
            "app.application.use_cases.amortization_fill_dry_run.resolve_sharepoint_path",
            new_callable=AsyncMock,
            return_value=ctx,
        ):
            with pytest.raises(ValueError, match="merge_control_amortization_not_ready"):
                await run_amortization_fill_dry_run(g)

    asyncio.run(run())
