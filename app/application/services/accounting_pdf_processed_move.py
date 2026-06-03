"""
Mueve PDFs de asientos contables a PROCESADOS dentro de la carpeta ASIENTOS CONTABLES CRED del crédito.
"""

from __future__ import annotations

import logging
import re
from collections import Counter
from dataclasses import dataclass
from typing import Any, Literal

import httpx

from app.application.sharepoint_resolution import encode_graph_drive_path
from app.domain.ports.graph import GraphApiPort

logger = logging.getLogger(__name__)

PROCESADOS_FOLDER_NAME = "PROCESADOS"
ASIENTOS_FOLDER_MARKER = "ASIENTOS CONTABLES CRED"

AccountingPdfMoveStatus = Literal[
    "moved", "already_moved", "skipped", "warning", "error"
]


@dataclass(frozen=True)
class AccountingPdfMoveRecord:
    source_path: str
    processed_folder_path: str
    destination_path: str
    status: AccountingPdfMoveStatus
    reason: str
    bank_code: str
    credito: str
    id_pago: str
    event_index: int

    def as_dict(self) -> dict[str, Any]:
        return {
            "source_path": self.source_path,
            "processed_folder_path": self.processed_folder_path,
            "destination_path": self.destination_path,
            "status": self.status,
            "reason": self.reason,
            "bank_code": self.bank_code,
            "credito": self.credito,
            "id_pago": self.id_pago,
            "event_index": self.event_index,
        }


def _normalize_rel_path(path: str) -> str:
    return path.strip().strip("/").replace("\\", "/")


def _sanitize_filename_token(value: str, *, fallback: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value or "").strip())
    cleaned = re.sub(r"[^\w.\-]+", "_", cleaned)
    cleaned = cleaned.strip("._-")
    return cleaned or fallback


def _credito_slug(credito: str) -> str:
    digits = re.search(r"\d+", str(credito or ""))
    if digits:
        return digits.group(0)
    return _sanitize_filename_token(credito, fallback="credito")


def _parent_asientos_folder(asiento_pdf_path: str) -> str | None:
    """
    Carpeta que contiene el PDF (ASIENTOS CONTABLES CRED {n}), no PROCESADOS.
    """
    rel = _normalize_rel_path(asiento_pdf_path)
    if not rel or "/" not in rel:
        return None
    parent = rel.rsplit("/", 1)[0]
    parts = parent.split("/")
    if any(p.casefold() == PROCESADOS_FOLDER_NAME.casefold() for p in parts):
        idx = next(
            i
            for i, p in enumerate(parts)
            if p.casefold() == PROCESADOS_FOLDER_NAME.casefold()
        )
        if idx <= 0:
            return None
        parent = "/".join(parts[:idx])
    if ASIENTOS_FOLDER_MARKER.casefold() not in parent.casefold():
        return None
    return parent


def _is_under_procesados(asiento_pdf_path: str) -> bool:
    parts = _normalize_rel_path(asiento_pdf_path).split("/")
    return any(p.casefold() == PROCESADOS_FOLDER_NAME.casefold() for p in parts)


def build_processed_asiento_filename(
    *,
    payment_date_iso: str,
    bank_code: str,
    credito: str,
    id_pago: str,
    event_index: int,
    use_event_suffix: bool,
    collision_suffix: int = 0,
) -> str:
    date_part = _sanitize_filename_token(payment_date_iso, fallback="sin-fecha")
    bank_part = _sanitize_filename_token(bank_code, fallback="banco")
    cred_part = _credito_slug(credito)
    pago_part = _sanitize_filename_token(id_pago, fallback="pago")
    stem = f"asiento_{date_part}_{bank_part}_credito-{cred_part}_pago-{pago_part}"
    if use_event_suffix and event_index > 0:
        stem += f"_evento-{event_index}"
    if collision_suffix > 0:
        stem += f"_{collision_suffix}"
    return f"{stem}.pdf"


def _item_endpoint(site_id: str, drive_id: str, rel_path: str) -> str:
    enc = encode_graph_drive_path(rel_path)
    return f"/sites/{site_id}/drives/{drive_id}/root:/{enc}:"


async def _drive_item_exists(
    graph: GraphApiPort, site_id: str, drive_id: str, rel_path: str
) -> bool:
    try:
        await graph.get(_item_endpoint(site_id, drive_id, rel_path))
        return True
    except httpx.HTTPStatusError as exc:
        if exc.response is not None and exc.response.status_code == 404:
            return False
        raise


async def _ensure_procesados_folder(
    graph: GraphApiPort, site_id: str, drive_id: str, asientos_parent: str
) -> str:
    processed_folder = f"{asientos_parent}/{PROCESADOS_FOLDER_NAME}"
    if await _drive_item_exists(graph, site_id, drive_id, processed_folder):
        return processed_folder
    enc_parent = encode_graph_drive_path(asientos_parent)
    children_ep = (
        f"/sites/{site_id}/drives/{drive_id}/root:/{enc_parent}:/children"
    )
    try:
        await graph.post_json(
            children_ep,
            {
                "name": PROCESADOS_FOLDER_NAME,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail",
            },
        )
    except httpx.HTTPStatusError as exc:
        if exc.response is not None and exc.response.status_code in (409, 200):
            pass
        elif await _drive_item_exists(graph, site_id, drive_id, processed_folder):
            pass
        else:
            raise
    return processed_folder


def _allocate_destination_path(
    processed_folder: str,
    *,
    payment_date_iso: str,
    bank_code: str,
    credito: str,
    id_pago: str,
    event_index: int,
    use_event_suffix: bool,
    existing_destinations: set[str],
) -> str:
    for collision in range(0, 50):
        name = build_processed_asiento_filename(
            payment_date_iso=payment_date_iso,
            bank_code=bank_code,
            credito=credito,
            id_pago=id_pago,
            event_index=event_index,
            use_event_suffix=use_event_suffix,
            collision_suffix=collision,
        )
        dest = f"{processed_folder}/{name}"
        key = dest.casefold()
        if key not in existing_destinations:
            existing_destinations.add(key)
            return dest
    raise ValueError("destination_name_exhausted")


async def move_used_accounting_pdf_to_processed(
    graph: GraphApiPort,
    site_id: str,
    drive_id: str,
    *,
    asiento_pdf_path: str,
    bank_code: str,
    id_pago: str,
    credito: str,
    payment_date_iso: str,
    event_index: int,
    use_event_suffix: bool,
    existing_destinations: set[str],
) -> AccountingPdfMoveRecord:
    source = _normalize_rel_path(asiento_pdf_path)
    bank_code = (bank_code or "").strip()
    id_pago = (id_pago or "").strip()
    credito = (credito or "").strip()
    event_index = int(event_index or 0)

    if not source:
        return AccountingPdfMoveRecord(
            source_path="",
            processed_folder_path="",
            destination_path="",
            status="skipped",
            reason="missing_asiento_pdf_path",
            bank_code=bank_code,
            credito=credito,
            id_pago=id_pago,
            event_index=event_index,
        )

    asientos_parent = _parent_asientos_folder(source)
    if not asientos_parent:
        return AccountingPdfMoveRecord(
            source_path=source,
            processed_folder_path="",
            destination_path="",
            status="warning",
            reason="ASIENTO_NOT_UNDER_ASIENTOS_CONTABLES_CRED_FOLDER",
            bank_code=bank_code,
            credito=credito,
            id_pago=id_pago,
            event_index=event_index,
        )

    processed_folder = f"{asientos_parent}/{PROCESADOS_FOLDER_NAME}"

    if _is_under_procesados(source):
        return AccountingPdfMoveRecord(
            source_path=source,
            processed_folder_path=processed_folder,
            destination_path=source,
            status="already_moved",
            reason="source_already_under_PROCESADOS",
            bank_code=bank_code,
            credito=credito,
            id_pago=id_pago,
            event_index=event_index,
        )

    try:
        processed_folder = await _ensure_procesados_folder(
            graph, site_id, drive_id, asientos_parent
        )
    except Exception as exc:
        logger.warning("No se pudo crear PROCESADOS en %s: %s", asientos_parent, exc)
        return AccountingPdfMoveRecord(
            source_path=source,
            processed_folder_path=processed_folder,
            destination_path="",
            status="error",
            reason=f"PROCESADOS_FOLDER_CREATE_FAILED|{exc!s}"[:500],
            bank_code=bank_code,
            credito=credito,
            id_pago=id_pago,
            event_index=event_index,
        )

    dest_path = _allocate_destination_path(
        processed_folder,
        payment_date_iso=payment_date_iso,
        bank_code=bank_code,
        credito=credito,
        id_pago=id_pago,
        event_index=event_index,
        use_event_suffix=use_event_suffix,
        existing_destinations=existing_destinations,
    )
    dest_name = dest_path.rsplit("/", 1)[-1]

    source_exists = await _drive_item_exists(graph, site_id, drive_id, source)
    dest_exists = await _drive_item_exists(graph, site_id, drive_id, dest_path)

    if not source_exists and dest_exists:
        return AccountingPdfMoveRecord(
            source_path=source,
            processed_folder_path=processed_folder,
            destination_path=dest_path,
            status="already_moved",
            reason="source_missing_destination_present",
            bank_code=bank_code,
            credito=credito,
            id_pago=id_pago,
            event_index=event_index,
        )

    if not source_exists and not dest_exists:
        return AccountingPdfMoveRecord(
            source_path=source,
            processed_folder_path=processed_folder,
            destination_path=dest_path,
            status="warning",
            reason="ACCOUNTING_PDF_NOT_FOUND_FOR_MOVE",
            bank_code=bank_code,
            credito=credito,
            id_pago=id_pago,
            event_index=event_index,
        )

    if source_exists and dest_exists:
        alt_dest = _allocate_destination_path(
            processed_folder,
            payment_date_iso=payment_date_iso,
            bank_code=bank_code,
            credito=credito,
            id_pago=id_pago,
            event_index=event_index,
            use_event_suffix=use_event_suffix,
            existing_destinations=existing_destinations,
        )
        dest_path = alt_dest
        dest_name = dest_path.rsplit("/", 1)[-1]
        dest_exists = await _drive_item_exists(graph, site_id, drive_id, dest_path)
        if dest_exists:
            return AccountingPdfMoveRecord(
                source_path=source,
                processed_folder_path=processed_folder,
                destination_path=dest_path,
                status="warning",
                reason="destination_exists_no_overwrite",
                bank_code=bank_code,
                credito=credito,
                id_pago=id_pago,
                event_index=event_index,
            )

    parent_ref_path = f"/drive/root:/{processed_folder}"
    try:
        await graph.patch_json(
            _item_endpoint(site_id, drive_id, source),
            {
                "parentReference": {
                    "driveId": drive_id,
                    "path": parent_ref_path,
                },
                "name": dest_name,
            },
        )
        return AccountingPdfMoveRecord(
            source_path=source,
            processed_folder_path=processed_folder,
            destination_path=dest_path,
            status="moved",
            reason="",
            bank_code=bank_code,
            credito=credito,
            id_pago=id_pago,
            event_index=event_index,
        )
    except Exception as exc:
        logger.warning("Fallo moviendo asiento %s -> %s: %s", source, dest_path, exc)
        return AccountingPdfMoveRecord(
            source_path=source,
            processed_folder_path=processed_folder,
            destination_path=dest_path,
            status="error",
            reason=str(exc)[:500],
            bank_code=bank_code,
            credito=credito,
            id_pago=id_pago,
            event_index=event_index,
        )


def _payment_date_iso_for_item(item: dict[str, Any], dry_run: dict[str, Any] | None) -> str:
    raw = str(item.get("payment_date_iso") or "").strip()
    if raw:
        return raw
    if dry_run:
        return str(dry_run.get("report_date_iso") or "").strip()
    return ""


def _event_suffix_flags(items: list[dict[str, Any]]) -> dict[tuple[str, str], bool]:
    counts: Counter[tuple[str, str]] = Counter()
    for it in items:
        key = (
            str(it.get("id_pago") or "").strip(),
            _credito_slug(str(it.get("credito") or "")),
        )
        if key[0]:
            counts[key] += 1
    return {k: counts[k] > 1 for k in counts}


async def process_used_accounting_pdfs_after_apply(
    graph: GraphApiPort,
    site_id: str,
    drive_id: str,
    *,
    items: list[dict[str, Any]],
    bank_code: str,
    verified_tabla_paths: set[str],
    dry_run: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    Mueve asientos de eventos APPLIED cuyas tablas pasaron verificación post-upload.
    """
    eligible: list[dict[str, Any]] = []
    for it in items:
        if not isinstance(it, dict):
            continue
        if it.get("apply_status") != "APPLIED":
            continue
        tabla = str(it.get("tabla_amortizacion_path") or "").strip()
        if tabla and tabla not in verified_tabla_paths:
            continue
        asiento = str(it.get("asiento_pdf_path") or "").strip()
        if not asiento:
            continue
        eligible.append(it)

    suffix_flags = _event_suffix_flags(eligible)
    existing_destinations: set[str] = set()
    moves: list[AccountingPdfMoveRecord] = []

    for it in eligible:
        id_pago = str(it.get("id_pago") or "").strip()
        credito = str(it.get("credito") or "").strip()
        event_index = int(it.get("event_index") or 0)
        use_suffix = suffix_flags.get((id_pago, _credito_slug(credito)), False)
        rec = await move_used_accounting_pdf_to_processed(
            graph,
            site_id,
            drive_id,
            asiento_pdf_path=str(it.get("asiento_pdf_path") or ""),
            bank_code=bank_code,
            id_pago=id_pago,
            credito=credito,
            payment_date_iso=_payment_date_iso_for_item(it, dry_run),
            event_index=event_index,
            use_event_suffix=use_suffix,
            existing_destinations=existing_destinations,
        )
        moves.append(rec)

    moved = sum(1 for m in moves if m.status == "moved")
    already = sum(1 for m in moves if m.status == "already_moved")
    warnings = sum(1 for m in moves if m.status == "warning")
    errors = sum(1 for m in moves if m.status == "error")
    skipped = sum(1 for m in moves if m.status == "skipped")

    return {
        "accounting_pdfs_processed_count": len(moves),
        "accounting_pdfs_moved_count": moved,
        "accounting_pdfs_already_moved_count": already,
        "accounting_pdfs_move_warnings_count": warnings,
        "accounting_pdfs_move_errors_count": errors,
        "accounting_pdfs_move_skipped_count": skipped,
        "accounting_pdfs_moves": [m.as_dict() for m in moves],
    }


def empty_accounting_pdf_move_summary() -> dict[str, Any]:
    return {
        "accounting_pdfs_processed_count": 0,
        "accounting_pdfs_moved_count": 0,
        "accounting_pdfs_already_moved_count": 0,
        "accounting_pdfs_move_warnings_count": 0,
        "accounting_pdfs_move_errors_count": 0,
        "accounting_pdfs_move_skipped_count": 0,
        "accounting_pdfs_moves": [],
    }
