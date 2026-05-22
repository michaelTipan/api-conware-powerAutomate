"""
Setup idempotente de IBR_DIARIO.xlsx (hoja IBR: Inicio, Fin, Valor).
"""

from __future__ import annotations

import io
import logging
import os
from typing import Any

import httpx
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from app.application.sharepoint_resolution import encode_graph_drive_path, resolve_sharepoint_path
from app.application.use_cases.setup_merge_control_workbook import (
    MERGE_CONTROL_FOLDER_RELATIVE_PATH,
    MergeControlSetupError,
    ensure_merge_control_folder_path,
)
from app.application.use_cases.setup_payment_followup_workbooks import followup_workbooks_folder_relative_path
from app.domain.exceptions import GraphConfigError
from app.domain.ports.graph import GraphApiPort

logger = logging.getLogger(__name__)

IBR_SHEET_NAME = "IBR"
IBR_COLUMNS: tuple[str, ...] = ("Inicio", "Fin", "Valor")
TABLE_IBR = "tblIbrDiario"
FILENAME_IBR = "IBR_DIARIO.xlsx"

_FILL_HEADER = PatternFill(fill_type="solid", fgColor="002060")
_FONT_HEADER = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
_FONT_BODY = Font(name="Calibri", size=11)
_BORDER = Border(
    left=Side(style="thin", color="C8C8C8"),
    right=Side(style="thin", color="C8C8C8"),
    top=Side(style="thin", color="C8C8C8"),
    bottom=Side(style="thin", color="C8C8C8"),
)
_ALIGN_HEADER = Alignment(vertical="center", horizontal="center", wrap_text=True)
_ALIGN_BODY = Alignment(vertical="center", horizontal="left", wrap_text=True)


class IbrWorkbookSetupError(Exception):
    def __init__(
        self,
        *,
        http_status: int,
        user_message: str,
        next_action: str,
        technical_message: str,
    ) -> None:
        self.http_status = http_status
        self.user_message = user_message
        self.next_action = next_action
        self.technical_message = technical_message
        super().__init__(technical_message)


def ibr_workbook_relative_path() -> str:
    p = os.getenv("GRAPH_IBR_DIARIO_PATH", "").strip()
    if p:
        return p
    folder = followup_workbooks_folder_relative_path().strip().rstrip("/")
    return f"{folder}/{FILENAME_IBR}"


def _content_endpoint(site_id: str, drive_id: str, file_path: str) -> str:
    return f"/sites/{site_id}/drives/{drive_id}/root:/{encode_graph_drive_path(file_path)}:/content"


def _item_endpoint(site_id: str, drive_id: str, file_path: str) -> str:
    return f"/sites/{site_id}/drives/{drive_id}/root:/{encode_graph_drive_path(file_path)}:"


async def _file_exists(graph: GraphApiPort, site_id: str, drive_id: str, path: str) -> bool:
    try:
        await graph.get_bytes(_content_endpoint(site_id, drive_id, path))
        return True
    except httpx.HTTPStatusError as exc:
        if exc.response is not None and exc.response.status_code == 404:
            return False
        raise


async def _drive_item_web_url(graph: GraphApiPort, site_id: str, drive_id: str, path: str) -> str | None:
    try:
        meta = await graph.get(_item_endpoint(site_id, drive_id, path))
        wu = meta.get("webUrl")
        return str(wu).strip() if isinstance(wu, str) and wu.strip() else None
    except httpx.HTTPStatusError as exc:
        if exc.response is not None and exc.response.status_code == 404:
            return None
        raise


def _build_new_ibr_workbook_bytes() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = IBR_SHEET_NAME
    _apply_ibr_sheet_structure(ws, preserve_data_from_row=2)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _apply_ibr_sheet_structure(ws: Any, *, preserve_data_from_row: int) -> list[str]:
    """Asegura encabezados y tabla; preserva filas >= preserve_data_from_row."""
    warnings: list[str] = []
    ncols = len(IBR_COLUMNS)
    existing_headers: list[str] = []
    max_c = ws.max_column or 0
    for c in range(1, max(max_c, ncols) + 1):
        existing_headers.append(str(ws.cell(1, c).value or "").strip())

    present = {h for h in existing_headers if h}
    for col_name in IBR_COLUMNS:
        if col_name not in present:
            next_col = len([h for h in existing_headers if h]) + 1
            while len(existing_headers) < next_col:
                existing_headers.append("")
            existing_headers[next_col - 1] = col_name
            warnings.append(f"column_added:{col_name}")

    for col_idx, name in enumerate(IBR_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.fill = _FILL_HEADER
        cell.font = _FONT_HEADER
        cell.alignment = _ALIGN_HEADER
        cell.border = _BORDER

    if ws.max_row < 2:
        for col_idx in range(1, ncols + 1):
            cell = ws.cell(row=2, column=col_idx, value=None)
            cell.font = _FONT_BODY
            cell.alignment = _ALIGN_BODY
            cell.border = _BORDER

    last_col = get_column_letter(ncols)
    last_row = max(ws.max_row or 2, 2)
    ws.auto_filter.ref = f"A1:{last_col}{last_row}"
    ws.freeze_panes = "A2"

    if TABLE_IBR in ws.tables:
        del ws.tables[TABLE_IBR]
    ref = f"A1:{last_col}{last_row}"
    tab = Table(displayName=TABLE_IBR, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)
    try:
        ws.sheet_view.showGridLines = False
    except Exception:
        pass
    try:
        ws.sheet_properties.tabColor = Color(rgb="002060")
    except Exception:
        pass
    return warnings


def _repair_existing_workbook(data: bytes) -> tuple[bytes, list[str], bool]:
    """Añade hoja/columnas IBR sin borrar otras hojas ni filas de datos."""
    warnings: list[str] = []
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    modified = False
    try:
        if IBR_SHEET_NAME not in wb.sheetnames:
            ws = wb.create_sheet(IBR_SHEET_NAME)
            warnings.append(f"sheet_added:{IBR_SHEET_NAME}")
            modified = True
        else:
            ws = wb[IBR_SHEET_NAME]

        w = _apply_ibr_sheet_structure(ws, preserve_data_from_row=2)
        if w:
            warnings.extend(w)
            modified = True

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue(), warnings, modified
    finally:
        closer = getattr(wb, "close", None)
        if callable(closer):
            closer()


async def setup_ibr_workbook(graph: GraphApiPort) -> dict[str, Any]:
    site_search = os.getenv("GRAPH_SHAREPOINT_SITE_SEARCH", "").strip()
    drive_name = os.getenv("GRAPH_SHAREPOINT_DRIVE_NAME", "").strip()
    if not site_search:
        raise GraphConfigError("Missing environment variable: GRAPH_SHAREPOINT_SITE_SEARCH")

    folder_rel = followup_workbooks_folder_relative_path()
    base = await resolve_sharepoint_path(graph, site_search, drive_name, folder_rel)
    site_id = base["site_id"]
    drive_id = base["drive_id"]
    file_path = ibr_workbook_relative_path()

    try:
        await ensure_merge_control_folder_path(graph, site_id, drive_id, folder_rel)
    except MergeControlSetupError as exc:
        raise IbrWorkbookSetupError(
            http_status=exc.http_status,
            user_message=exc.user_message,
            next_action=exc.next_action,
            technical_message=exc.technical_message,
        ) from exc

    warnings: list[str] = []
    exists = await _file_exists(graph, site_id, drive_id, file_path)

    try:
        if not exists:
            payload = _build_new_ibr_workbook_bytes()
            resp = await graph.put_bytes(
                _content_endpoint(site_id, drive_id, file_path),
                payload,
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            file_url = None
            if isinstance(resp, dict):
                wu = resp.get("webUrl")
                file_url = str(wu).strip() if isinstance(wu, str) and wu.strip() else None
            if not file_url:
                file_url = await _drive_item_web_url(graph, site_id, drive_id, file_path)
            return {
                "status": "success",
                "file_path": file_path,
                "created": True,
                "repaired": False,
                "file_url": file_url,
                "sheet_name": IBR_SHEET_NAME,
                "columns": list(IBR_COLUMNS),
                "warnings": warnings,
            }

        raw = await graph.get_bytes(_content_endpoint(site_id, drive_id, file_path))
        payload, repair_warnings, modified = _repair_existing_workbook(raw)
        warnings.extend(repair_warnings)
        file_url = await _drive_item_web_url(graph, site_id, drive_id, file_path)
        if modified:
            resp = await graph.put_bytes(
                _content_endpoint(site_id, drive_id, file_path),
                payload,
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            if isinstance(resp, dict):
                wu = resp.get("webUrl")
                if isinstance(wu, str) and wu.strip():
                    file_url = wu.strip()

        return {
            "status": "success",
            "file_path": file_path,
            "created": False,
            "repaired": modified,
            "file_url": file_url,
            "sheet_name": IBR_SHEET_NAME,
            "columns": list(IBR_COLUMNS),
            "warnings": warnings,
        }
    except httpx.HTTPStatusError as exc:
        code = exc.response.status_code if exc.response else 0
        body_txt = (exc.response.text if exc.response else "") or ""
        raise IbrWorkbookSetupError(
            http_status=502 if code != 403 else 403,
            user_message="Error al preparar IBR_DIARIO.xlsx en SharePoint.",
            next_action="Verifique permisos o bloqueos en 00 CONTROL.",
            technical_message=f"Graph HTTP {code}: {body_txt[:2000]}",
        ) from exc
