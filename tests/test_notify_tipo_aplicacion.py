"""Notify Fase 3: PAGO + ABONO en correo y contadores."""

import asyncio
from datetime import date
from io import BytesIO
from unittest.mock import AsyncMock, patch

import pytest
from openpyxl import Workbook, load_workbook

from app.application.services.review_schema import (
    DistribucionAbonosCols,
    DistribucionCols,
    ReviewSheets,
    ValidarAbono,
    ValidarPago,
)
from app.application.use_cases.send_validar_extractos_notification import (
    _abono_table_from_groups,
    _build_html,
    send_validar_extractos_notification_email,
)
from app.application.services.historical_application_rows import (
    group_abono_rows_for_email,
    read_validated_abono_rows,
)
from tests.test_finalize_validation import make_abono_row, make_distrib_row


def _bank_xlsx() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append(["Fecha", "Concepto", "Crédito", "Monto"])
    ws.append([date(2026, 6, 3), "pago", "258", "1000"])
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


def _historico_pago_abono_bytes() -> bytes:
    r, _ = make_distrib_row(
        id_pago="P1",
        validar_pago=ValidarPago.SI,
        ruta_pdf_internal="clientes/CLI/extractos/e1.pdf",
    )
    a1, _ = make_abono_row(
        id_pago="AB123",
        cliente="EQUINORTE",
        monto_banco=2_000_000,
        fecha_banco=date(2026, 6, 3),
        validar_abono=ValidarAbono.SI,
        credito="258",
        credito_normalizado="258",
    )
    a2, _ = make_abono_row(
        id_pago="AB123",
        cliente="EQUINORTE",
        monto_banco=2_000_000,
        fecha_banco=date(2026, 6, 3),
        validar_abono=ValidarAbono.SI,
        credito="265",
        credito_normalizado="265",
    )
    wb = Workbook()
    ws = wb.active
    ws.title = ReviewSheets.DISTRIBUCION
    ws.append(DistribucionCols.HEADERS)
    ws.append(r)
    ws_ab = wb.create_sheet(ReviewSheets.DISTRIBUCION_ABONOS)
    ws_ab.append(list(DistribucionAbonosCols.HEADERS) + [DistribucionAbonosCols.RUTA_ASIENTOS_CONTABLES])
    for row in (a1, a2):
        ws_ab.append(list(row) + [f"clientes/EQUINORTE/CREDITO# {row[2]}/ASIENTOS CONTABLES"])
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


def test_abono_table_single_row_per_id_pago():
    wb = load_workbook(BytesIO(_historico_pago_abono_bytes()), data_only=True)
    groups = group_abono_rows_for_email(read_validated_abono_rows(wb))
    headers, rows = _abono_table_from_groups(groups)
    assert len(rows) == 1
    assert rows[0][0] == "AB123"
    assert "258" in rows[0][-1] and "265" in rows[0][-1]
    assert rows[0][3] == "2.000.000"


def test_build_html_includes_abono_section():
    html = _build_html(
        intro="<p>Hola</p>",
        bank_headers=["Fecha"],
        bank_rows=[["03/06/2026"]],
        abono_headers=["ID Pago", "Créditos seleccionados"],
        abono_rows=[["AB123", "258, 265"]],
    )
    assert "Abonos (sin extracto)" in html
    assert "AB123" in html
    assert "258, 265" in html


class _NotifyGraph:
    def __init__(self, hist_bytes: bytes, extract_pdf: bytes = b"%PDF"):
        self.hist_bytes = hist_bytes
        self.extract_pdf = extract_pdf
        self.sent: list[dict] = []

    async def get(self, endpoint, params=None):
        if endpoint == "/sites":
            return {"value": [{"id": "s1"}]}
        if endpoint == "/sites/s1/drives":
            return {"value": [{"id": "d1", "name": "DRIVE"}]}
        return {"value": []}

    async def get_bytes(self, endpoint, params=None):
        ep = str(endpoint)
        if "HIST/" in ep or "cartera_validada" in ep:
            return self.hist_bytes
        if "extractos" in ep or "e1.pdf" in ep:
            return self.extract_pdf
        if ".pdf" in ep.lower():
            return b"%PDF-1.4 mock"
        return _bank_xlsx()

    async def post_json(self, endpoint, body):
        self.sent.append(body)
        return {}, 202

    async def put_bytes(self, endpoint, content, content_type=None):
        return {}


@pytest.fixture(autouse=True)
def _notify_env(monkeypatch):
    monkeypatch.setenv("GRAPH_SHAREPOINT_SITE_SEARCH", "SITIO")
    monkeypatch.setenv("GRAPH_SHAREPOINT_DRIVE_NAME", "DRIVE")
    monkeypatch.setenv("GRAPH_SHAREPOINT_FILE_PATH", "banco.xlsx")
    monkeypatch.setenv("GRAPH_BANK_PAYMENTS_FILE_PATH", "banco.xlsx")
    monkeypatch.setenv("GRAPH_VALIDAR_NOTIFY_ATTACH_PDFS", "false")
    monkeypatch.setenv("GRAPH_VALIDAR_NOTIFY_EXPORT_EMAIL_PDF", "false")


def test_notify_abono_without_ruta_does_not_fail(monkeypatch):
    r, _ = make_distrib_row(validar_pago=ValidarPago.NO, estado="ATRASADO")
    a1, _ = make_abono_row(id_pago="AB1", validar_abono=ValidarAbono.SI)
    wb = Workbook()
    ws = wb.active
    ws.title = ReviewSheets.DISTRIBUCION
    ws.append(DistribucionCols.HEADERS)
    ws.append(r)
    ws_ab = wb.create_sheet(ReviewSheets.DISTRIBUCION_ABONOS)
    ws_ab.append(list(DistribucionAbonosCols.HEADERS) + [DistribucionAbonosCols.RUTA_ASIENTOS_CONTABLES])
    ws_ab.append(list(a1) + ["clientes/CLI/ASIENTOS"])
    bio = BytesIO()
    wb.save(bio)
    g = _NotifyGraph(bio.getvalue())

    async def _correos(*_a, **_k):
        return "sender@test.com", ["dest@test.com"]

    async def _resolve(*_a, **_k):
        return {"site_id": "s1", "drive_id": "d1", "path_encoded": "x", "file_path": "banco.xlsx"}

    async def _from_env(*_a, **_k):
        return {"site_id": "s1", "drive_id": "d1"}

    with (
        patch(
            "app.application.use_cases.send_validar_extractos_notification._load_sender_and_recipients_from_correos_xlsx",
            new=AsyncMock(side_effect=_correos),
        ),
        patch(
            "app.application.use_cases.send_validar_extractos_notification.resolve_sharepoint_path",
            new=AsyncMock(side_effect=_resolve),
        ),
        patch(
            "app.application.use_cases.send_validar_extractos_notification.resolve_sharepoint_from_env",
            new=AsyncMock(side_effect=_from_env),
        ),
        patch(
            "app.application.use_cases.payment_validation_process_control.update_process_control_row2",
            new=AsyncMock(return_value=True),
        ),
    ):
        result = asyncio.run(
            send_validar_extractos_notification_email(
                g,
                historical_file_path="HIST/cartera_validada.xlsx",
                bank_code="banco_bogota",
            )
        )
    assert result.abono_groups_included == 1
    assert result.extracts_not_required_count == 1
    assert "Abonos (sin extracto)" in g.sent[0]["message"]["body"]["content"]


def test_notify_counters_pago_and_abono(monkeypatch):
    g = _NotifyGraph(_historico_pago_abono_bytes())

    async def _correos(*_a, **_k):
        return "sender@test.com", ["dest@test.com"]

    async def _resolve(*_a, **_k):
        return {"site_id": "s1", "drive_id": "d1", "path_encoded": "x", "file_path": "banco.xlsx"}

    async def _from_env(*_a, **_k):
        return {"site_id": "s1", "drive_id": "d1"}

    with (
        patch(
            "app.application.use_cases.send_validar_extractos_notification._load_sender_and_recipients_from_correos_xlsx",
            new=AsyncMock(side_effect=_correos),
        ),
        patch(
            "app.application.use_cases.send_validar_extractos_notification.resolve_sharepoint_path",
            new=AsyncMock(side_effect=_resolve),
        ),
        patch(
            "app.application.use_cases.send_validar_extractos_notification.resolve_sharepoint_from_env",
            new=AsyncMock(side_effect=_from_env),
        ),
        patch(
            "app.application.use_cases.payment_validation_process_control.update_process_control_row2",
            new=AsyncMock(return_value=True),
        ),
        patch(
            "app.application.use_cases.send_validar_extractos_notification._collect_pdf_paths_from_ruta_cell",
            new=AsyncMock(return_value=["clientes/CLI/extractos/e1.pdf"]),
        ),
        patch(
            "app.application.use_cases.send_validar_extractos_notification._pdf_attachments_from_drive_paths",
            new=AsyncMock(return_value=([{"name": "e1.pdf"}], [])),
        ),
    ):
        monkeypatch.setenv("GRAPH_VALIDAR_NOTIFY_ATTACH_PDFS", "true")
        result = asyncio.run(
            send_validar_extractos_notification_email(
                g,
                historical_file_path="HIST/cartera_validada.xlsx",
                bank_code="banco_bogota",
            )
        )
    assert result.payment_groups_included >= 1
    assert result.abono_groups_included == 1
    assert result.abono_credit_rows_included == 2
    assert result.extracts_attached_count == 1
    assert result.movement_groups_included == result.payment_groups_included + 1
