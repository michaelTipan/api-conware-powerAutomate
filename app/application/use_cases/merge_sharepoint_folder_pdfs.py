"""
Une todos los PDF de una carpeta en SharePoint en un solo archivo y lo sube a la misma carpeta.

Ruta de carpeta temporal (hasta parametrizar por API): ver MERGE_PDF_FOLDER_RELATIVE_PATH.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import Any
from urllib.parse import parse_qs, urlparse

from pypdf import PdfReader, PdfWriter

from app.application.sharepoint_resolution import encode_graph_drive_path, resolve_sharepoint_from_env
from app.application.use_cases.validate_payment_report import _graph_download_by_path
from app.domain.ports.graph import GraphApiPort

logger = logging.getLogger(__name__)

# Temporal: misma biblioteca que GRAPH_SHAREPOINT_*; ruta relativa al root del drive.
MERGE_PDF_FOLDER_RELATIVE_PATH = (
    "INFORMACION CREDITOS-CLIENTES/00 COMWARE - CARGA TRANSACCIONES BANCO"
)
# Salida estándar: "INGRESO DE PLATA DEL {año}.pdf" (año calendario al ejecutar).
MERGE_PDF_OUTPUT_NAME_TEMPLATE = "INGRESO DE PLATA DEL {year}.pdf"
# Nombre anterior; no se mezcla como fuente si sigue en la carpeta.
MERGE_PDF_LEGACY_OUTPUT_NAME = "PDFS_UNIFICADOS.pdf"


def _merge_output_basename(year: int | None = None) -> str:
    y = year if year is not None else datetime.now().year
    return MERGE_PDF_OUTPUT_NAME_TEMPLATE.format(year=y)


def _excluded_source_names_upper(output_basename: str) -> set[str]:
    return {
        output_basename.upper(),
        MERGE_PDF_LEGACY_OUTPUT_NAME.upper(),
    }


@dataclass(frozen=True)
class MergeFolderPdfsResult:
    folder_path: str
    output_relative_path: str
    source_pdf_names: tuple[str, ...]
    source_count: int
    output_size_bytes: int


def _endpoint_and_params_from_next_link(next_link: str) -> tuple[str, dict[str, str] | None]:
    parsed = urlparse(next_link)
    path = parsed.path or ""
    if "/v1.0/" in path:
        endpoint = path.split("/v1.0/", 1)[1].lstrip("/")
    else:
        endpoint = path.lstrip("/")
    if not parsed.query:
        return endpoint, None
    flat = {k: v[0] for k, v in parse_qs(parsed.query).items() if v}
    return endpoint, flat or None


async def _list_drive_folder_children(
    graph: GraphApiPort,
    site_id: str,
    drive_id: str,
    folder_path: str,
) -> list[dict[str, Any]]:
    encoded = encode_graph_drive_path(folder_path.strip().strip("/"))
    path = f"/sites/{site_id}/drives/{drive_id}/root:/{encoded}:/children"
    items: list[dict[str, Any]] = []
    endpoint = path.lstrip("/")
    params: dict[str, str] | None = {"$top": "200"}
    while True:
        resp = await graph.get(endpoint, params=params)
        items.extend(resp.get("value") or [])
        next_link = resp.get("@odata.nextLink")
        if not next_link:
            break
        endpoint, params = _endpoint_and_params_from_next_link(next_link)
    return items


async def merge_pdfs_in_configured_folder(graph: GraphApiPort) -> MergeFolderPdfsResult:
    """
    Lista PDF en MERGE_PDF_FOLDER_RELATIVE_PATH, los concatena en orden por nombre de archivo,
    excluye el PDF de salida del año en curso y el legado PDFS_UNIFICADOS.pdf, y sube
    ``INGRESO DE PLATA DEL {año}.pdf``.
    """
    ctx = await resolve_sharepoint_from_env(graph)
    site_id = ctx["site_id"]
    drive_id = ctx["drive_id"]
    folder = MERGE_PDF_FOLDER_RELATIVE_PATH.strip().strip("/")
    output_basename = _merge_output_basename()
    excluded = _excluded_source_names_upper(output_basename)

    children = await _list_drive_folder_children(graph, site_id, drive_id, folder)

    pdf_names: list[str] = []
    for item in children:
        name = (item.get("name") or "").strip()
        if not name:
            continue
        if "folder" in item:
            continue
        if "file" not in item:
            continue
        if not name.lower().endswith(".pdf"):
            continue
        if name.upper() in excluded:
            continue
        pdf_names.append(name)

    pdf_names.sort()

    if not pdf_names:
        raise ValueError(
            f"No hay archivos .pdf en la carpeta (excl. salida {output_basename!r}): {folder!r}"
        )

    writer = PdfWriter()
    for pdf_name in pdf_names:
        rel_path = f"{folder}/{pdf_name}"
        try:
            raw = await _graph_download_by_path(graph, site_id, drive_id, rel_path)
            reader = PdfReader(BytesIO(raw))
            for page in reader.pages:
                writer.add_page(page)
        except Exception as exc:
            raise ValueError(f"No se pudo leer o unir el PDF {pdf_name!r}: {exc}") from exc

    out_buf = BytesIO()
    writer.write(out_buf)
    merged_bytes = out_buf.getvalue()
    out_rel = f"{folder}/{output_basename}"
    encoded = encode_graph_drive_path(out_rel)

    logger.info(
        "merge_pdfs_in_configured_folder: subiendo PDF unificado",
        extra={
            "folder": folder,
            "sources": pdf_names,
            "output": out_rel,
            "bytes": len(merged_bytes),
        },
    )

    await graph.put_bytes(
        f"/sites/{site_id}/drives/{drive_id}/root:/{encoded}:/content",
        merged_bytes,
        content_type="application/pdf",
    )

    return MergeFolderPdfsResult(
        folder_path=folder,
        output_relative_path=out_rel,
        source_pdf_names=tuple(pdf_names),
        source_count=len(pdf_names),
        output_size_bytes=len(merged_bytes),
    )
