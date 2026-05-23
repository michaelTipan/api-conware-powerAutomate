"""Tests credit_items, deduplicación de extractos y orden del PDF consolidado."""

import asyncio
import json
from unittest.mock import AsyncMock, patch

import pytest

from app.application.sharepoint_resolution import encode_graph_drive_path
from app.application.use_cases.merge_composite_validado_pdfs import (
    _build_consolidated_pdf_parts,
    _finalize_credit_item,
    merge_composite_validado_pdfs,
    normalize_sharepoint_path,
    unique_paths_preserve_order,
)
from tests.test_merge_composite_control_workbook import (
    _MergeGraph,
    _bank_bytes,
    _control_row_bytes,
    _hist_workbook_bytes,
    _tiny_pdf,
)


def test_unique_paths_preserve_order_case_insensitive():
    paths = [
        "clientes/X/Extracto 265.pdf",
        "clientes/X/EXTRACTO 265.PDF",
        " clientes/X/Extracto 265.pdf ",
    ]
    out = unique_paths_preserve_order(paths)
    assert len(out) == 1
    assert normalize_sharepoint_path(out[0]) == normalize_sharepoint_path(paths[0])


def test_finalize_credit_item_warns_multiple_extracts():
    item = _finalize_credit_item(
        "265",
        ["a/asiento1.pdf", "a/asiento2.pdf"],
        ["e/ex1.pdf", "e/ex2.pdf"],
    )
    assert len(item["extracto_pdf_paths"]) == 2
    assert "MULTIPLE_EXTRACTS_FOR_CREDIT" in item["warnings"]
    assert len(item["asiento_pdf_paths"]) == 2


def test_build_pdf_order_asientos_before_single_extract(monkeypatch):
    """Simula orden: asiento1, asiento2, extracto (no extracto entre asientos)."""
    credit_items = [
        _finalize_credit_item(
            "265",
            ["c/CRED 265/Asiento 1 265.pdf", "c/CRED 265/Asiento 2 265.pdf"],
            ["c/CRED 265/Extracto 265.pdf", "c/CRED 265/Extracto 265.pdf"],
        )
    ]
    assert len(credit_items[0]["extracto_pdf_paths"]) == 1

    class G:
        async def get_bytes(self, *a, **k):
            return b"%PDF"

    parts, labels, skips = asyncio.run(
        _build_consolidated_pdf_parts(G(), "s", "d", "P1", b"email", "EMAIL/x.pdf", credit_items)
    )
    assert not skips
    assert labels[0] == "email:EMAIL/x.pdf"
    assert labels[1].startswith("asiento:")
    assert labels[2].startswith("asiento:")
    assert labels[3].startswith("extracto:")
    assert labels.count("extracto:c/CRED 265/Extracto 265.pdf") == 1


def test_merge_two_credits_credit_items_and_no_duplicate_extract_265(monkeypatch):
    monkeypatch.setenv("GRAPH_MERGE_CONTROL_WORKBOOK_PATH", "CTL/control.xlsx")
    monkeypatch.setenv("GRAPH_MERGE_COMPOSITE_OUTPUT_FOLDER_PATH", "OUT/PDFS")
    monkeypatch.setenv("GRAPH_PAYMENT_VALIDATION_LOGS_PATH", "LOGS")
    hist = "HIST/hist.xlsx"
    email = "EMAIL/mail.pdf"
    p258 = "clientes/GEO/CREDITO # 258/Extracto 258.pdf"
    p265 = "clientes/GEO/CREDITO # 265/Extracto 265.pdf"
    d258 = "clientes/GEO/CREDITO # 258/ASIENTOS CONTABLES CRED 258"
    d265 = "clientes/GEO/CREDITO # 265/ASIENTOS CONTABLES CRED 265"
    a258 = f"{d258}/Asiento cuota credito 258.pdf"
    a265_1 = f"{d265}/Asiento 1 credito 265.pdf"
    a265_2 = f"{d265}/Asiento 2 credito 265.pdf"

    pdf = _tiny_pdf()
    g = _MergeGraph()
    g.initial["CTL/control.xlsx"] = _control_row_bytes(
        estado="PENDIENTE_ASIENTOS",
        is_active=True,
        hist=hist,
        email=email,
    )
    g.initial["bank/report.xlsx"] = _bank_bytes()
    g.initial[hist] = _hist_workbook_bytes(
        [
            ["VALIDAR", "", "G1", "GEO", "258", d258],
            ["VALIDAR", "", "G1", "GEO", "265", d265],
        ]
    )
    g.initial[email] = pdf
    g.initial[p258] = pdf
    g.initial[p265] = pdf
    g.initial[a258] = pdf
    g.initial[a265_1] = pdf
    g.initial[a265_2] = pdf
    g.children[d258] = [{"name": "Asiento cuota credito 258.pdf", "file": {}}]
    g.children[d265] = [
        {"name": "Asiento 1 credito 265.pdf", "file": {}},
        {"name": "Asiento 2 credito 265.pdf", "file": {}},
    ]

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
            return await merge_composite_validado_pdfs(g)

    r = asyncio.run(run())
    out = r.outputs[0]
    assert len(out.credit_items) == 2
    by_cred = {str(ci["credito"]): ci for ci in out.credit_items}
    assert len(by_cred["258"]["asiento_pdf_paths"]) == 1
    assert len(by_cred["265"]["asiento_pdf_paths"]) == 2
    assert len(by_cred["265"]["extracto_pdf_paths"]) == 1
    assert out.extracto_pdf_path.count(p265) == 1
    assert p265 not in out.extracto_pdf_path or out.extracto_pdf_path.count("265") <= 2
    manifest_key = next(k for k in g.uploaded if k.endswith(".json"))
    data = json.loads(g.uploaded[manifest_key].decode("utf-8"))
    m_ci = data["outputs"][0]["credit_items"]
    assert len(m_ci) == 2
    m265 = next(x for x in m_ci if x["credito"] == "265")
    assert len(m265["extracto_pdf_paths"]) == 1
    extract_labels = [lb for lb in out.sources_summary.split(" | ") if lb.startswith("extracto:")]
    assert len([lb for lb in extract_labels if "265" in lb]) == 1


def test_merge_force_rebuild_uploads_when_pdf_exists(monkeypatch):
    monkeypatch.setenv("GRAPH_MERGE_CONTROL_WORKBOOK_PATH", "CTL/control.xlsx")
    monkeypatch.setenv("GRAPH_MERGE_COMPOSITE_OUTPUT_FOLDER_PATH", "OUT/PDFS")
    hist = "HIST/hist.xlsx"
    email = "EMAIL/mail.pdf"
    extract = "clientes/ACME/CREDITO# 264/Extracto.pdf"
    asiento_dir = "clientes/ACME/CREDITO# 264/ASIENTOS CONTABLES CRED 264"
    asiento_rel = f"{asiento_dir}/asiento_264.pdf"

    g = _MergeGraph()
    g.initial["CTL/control.xlsx"] = _control_row_bytes(
        estado="PENDIENTE_ASIENTOS",
        is_active=True,
        hist=hist,
        email=email,
    )
    g.initial["bank/report.xlsx"] = _bank_bytes()
    g.initial[hist] = _hist_workbook_bytes(
        [["VALIDAR", "", "G1", "ACME", "264", asiento_dir]]
    )
    g.initial[email] = _tiny_pdf()
    g.initial[extract] = _tiny_pdf()
    g.initial[asiento_rel] = _tiny_pdf()
    g.children[asiento_dir] = [{"name": "asiento_264.pdf", "file": {}}]

    ctx = {
        "site_id": "s1",
        "drive_id": "d1",
        "path_encoded": encode_graph_drive_path("bank/report.xlsx"),
        "file_path": "bank/report.xlsx",
    }

    async def fake_collect(_g, _s, _d, _cell):
        return [extract]

    async def run(force: bool):
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
            return await merge_composite_validado_pdfs(g, force_rebuild=force)

    r1 = asyncio.run(run(False))
    out_rel = r1.outputs[0].output_relative_path
    g.initial[out_rel] = _tiny_pdf()

    g.uploaded.clear()
    r2 = asyncio.run(run(True))
    assert r2.outputs[0].bytes_written > 0
    assert out_rel in g.uploaded
