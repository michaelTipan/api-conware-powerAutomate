"""Tests movimiento de asientos a PROCESADOS tras apply."""

import asyncio
from datetime import date
from urllib.parse import unquote

import httpx
import pytest

from app.application.services.accounting_pdf_processed_move import (
    PROCESADOS_FOLDER_NAME,
    build_processed_asiento_filename,
    move_used_accounting_pdf_to_processed,
    process_used_accounting_pdfs_after_apply,
    _parent_asientos_folder,
)
from tests.test_amortization_fill_apply import MockGraphApply


class _MoveGraph(MockGraphApply):
    pass


@pytest.fixture(autouse=True)
def env_sharepoint(monkeypatch):
    monkeypatch.setenv("GRAPH_SHAREPOINT_SITE_SEARCH", "TEST")
    monkeypatch.setenv("GRAPH_SHAREPOINT_DRIVE_NAME", "")


def test_parent_asientos_folder_resolves_cred_folder():
    path = "clientes/E/CREDITO # 258/ASIENTOS CONTABLES CRED 258/Asiento.pdf"
    assert _parent_asientos_folder(path) == "clientes/E/CREDITO # 258/ASIENTOS CONTABLES CRED 258"


def test_build_filename_with_event_suffix():
    name = build_processed_asiento_filename(
        payment_date_iso="2026-04-23",
        bank_code="banco_bogota",
        credito="CREDITO # 265",
        id_pago="851eb0f6",
        event_index=2,
        use_event_suffix=True,
    )
    assert name == "asiento_2026-04-23_banco_bogota_credito-265_pago-851eb0f6_evento-2.pdf"


def test_move_creates_procesados_and_renames(monkeypatch):
    fecha = date(2026, 4, 23)
    asiento = (
        "clientes/EQUINORTE/CREDITO # 258/ASIENTOS CONTABLES CRED 258/Asiento EQUINORTE # 258.pdf"
    )
    g = _MoveGraph({asiento: b"%PDF"})
    site, drive = "s1", "d1"

    rec = asyncio.run(
        move_used_accounting_pdf_to_processed(
            g,
            site,
            drive,
            asiento_pdf_path=asiento,
            bank_code="banco_bogota",
            id_pago="851eb0f6",
            credito="CREDITO # 258",
            payment_date_iso=fecha.isoformat(),
            event_index=1,
            use_event_suffix=False,
            existing_destinations=set(),
        )
    )
    assert rec.status == "moved"
    assert PROCESADOS_FOLDER_NAME in rec.destination_path
    assert "ASIENTOS CONTABLES CRED 258" in rec.processed_folder_path
    assert asiento not in g.files
    assert (
        "clientes/EQUINORTE/CREDITO # 258/ASIENTOS CONTABLES CRED 258/PROCESADOS/"
        "asiento_2026-04-23_banco_bogota_credito-258_pago-851eb0f6.pdf"
    ) in g.files
    assert g.post_json_calls
    assert g.patch_calls


def test_already_moved_when_source_missing_dest_present():
    dest = (
        "clientes/E/CREDITO # 258/ASIENTOS CONTABLES CRED 258/PROCESADOS/"
        "asiento_2026-04-23_banco_bogota_credito-258_pago-851eb0f6.pdf"
    )
    g = _MoveGraph({dest: b"%PDF"})
    rec = asyncio.run(
        move_used_accounting_pdf_to_processed(
            g,
            "s1",
            "d1",
            asiento_pdf_path=(
                "clientes/E/CREDITO # 258/ASIENTOS CONTABLES CRED 258/Asiento.pdf"
            ),
            bank_code="banco_bogota",
            id_pago="851eb0f6",
            credito="CREDITO # 258",
            payment_date_iso="2026-04-23",
            event_index=1,
            use_event_suffix=False,
            existing_destinations=set(),
        )
    )
    assert rec.status == "already_moved"


def test_warning_when_neither_source_nor_dest():
    g = _MoveGraph({})
    rec = asyncio.run(
        move_used_accounting_pdf_to_processed(
            g,
            "s1",
            "d1",
            asiento_pdf_path="clientes/E/CREDITO # 258/ASIENTOS CONTABLES CRED 258/x.pdf",
            bank_code="banco_bogota",
            id_pago="851eb0f6",
            credito="CREDITO # 258",
            payment_date_iso="2026-04-23",
            event_index=1,
            use_event_suffix=False,
            existing_destinations=set(),
        )
    )
    assert rec.status == "warning"
    assert rec.reason == "ACCOUNTING_PDF_NOT_FOUND_FOR_MOVE"


def test_process_batch_uses_event_index_for_two_asientos_same_pago():
    asiento_a = "clientes/E/CREDITO # 265/ASIENTOS CONTABLES CRED 265/a.pdf"
    asiento_b = "clientes/E/CREDITO # 265/ASIENTOS CONTABLES CRED 265/b.pdf"
    g = _MoveGraph({asiento_a: b"a", asiento_b: b"b"})
    items = [
        {
            "apply_status": "APPLIED",
            "tabla_amortizacion_path": "TABLAS/t.xlsx",
            "asiento_pdf_path": asiento_a,
            "id_pago": "851eb0f6",
            "credito": "CREDITO # 265",
            "event_index": 1,
            "payment_date_iso": "2026-04-23",
        },
        {
            "apply_status": "APPLIED",
            "tabla_amortizacion_path": "TABLAS/t.xlsx",
            "asiento_pdf_path": asiento_b,
            "id_pago": "851eb0f6",
            "credito": "CREDITO # 265",
            "event_index": 2,
            "payment_date_iso": "2026-04-23",
        },
    ]
    out = asyncio.run(
        process_used_accounting_pdfs_after_apply(
            g,
            "s1",
            "d1",
            items=items,
            bank_code="banco_bogota",
            verified_tabla_paths={"TABLAS/t.xlsx"},
        )
    )
    assert out["accounting_pdfs_moved_count"] == 2
    names = [m["destination_path"] for m in out["accounting_pdfs_moves"]]
    assert any("_evento-1" in n for n in names)
    assert any("_evento-2" in n for n in names)


def test_skips_non_applied_items():
    g = _MoveGraph({"clientes/E/CREDITO # 258/ASIENTOS CONTABLES CRED 258/x.pdf": b"x"})
    out = asyncio.run(
        process_used_accounting_pdfs_after_apply(
            g,
            "s1",
            "d1",
            items=[
                {
                    "apply_status": "ERROR",
                    "tabla_amortizacion_path": "TABLAS/t.xlsx",
                    "asiento_pdf_path": "clientes/E/CREDITO # 258/ASIENTOS CONTABLES CRED 258/x.pdf",
                    "id_pago": "x",
                    "credito": "CREDITO # 258",
                }
            ],
            bank_code="banco_bogota",
            verified_tabla_paths={"TABLAS/t.xlsx"},
        )
    )
    assert out["accounting_pdfs_processed_count"] == 0
