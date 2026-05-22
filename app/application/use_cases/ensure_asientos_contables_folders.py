"""
Crea subcarpetas bajo cada carpeta de crédito de cada cliente
(GRAPH_CLIENTS_BASE_PATH / {cliente} / {credito}/): por defecto ``ASIENTOS CONTABLES`` y ``EXTRACTOS``.
Si una carpeta ya existe, no hace nada y continúa con las demás.
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass, field
from typing import Any
from urllib.parse import parse_qs, urlparse

import httpx

from app.application.sharepoint_resolution import encode_graph_drive_path, resolve_sharepoint_path
from app.domain.exceptions import GraphConfigError
from app.domain.ports.graph import GraphApiPort

logger = logging.getLogger(__name__)

ASIENTOS_FOLDER_NAME = "ASIENTOS CONTABLES"
EXTRACTOS_FOLDER_NAME = "EXTRACTOS"


def _dedupe_names_preserve_order(names: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for n in names:
        s = (n or "").strip()
        if not s:
            continue
        key = s.casefold()
        if key not in seen:
            seen.add(key)
            out.append(s)
    return out


def _credit_subfolder_names_to_ensure() -> list[str]:
    """
    Lista de nombres de carpeta a crear bajo cada crédito.
    Si ``GRAPH_CREDIT_SUBFOLDERS_TO_ENSURE`` está definida (coma), solo usa eso.
    Si no: ``GRAPH_ASIENTOS_CONTABLES_FOLDER_NAME`` + ``GRAPH_EXTRACTOS_FOLDER_NAME`` (defectos arriba).
    """
    raw = os.getenv("GRAPH_CREDIT_SUBFOLDERS_TO_ENSURE", "").strip()
    if raw:
        return _dedupe_names_preserve_order([p.strip() for p in raw.split(",") if p.strip()])
    asientos = (
        os.getenv("GRAPH_ASIENTOS_CONTABLES_FOLDER_NAME", ASIENTOS_FOLDER_NAME).strip()
        or ASIENTOS_FOLDER_NAME
    )
    extractos = (
        os.getenv("GRAPH_EXTRACTOS_FOLDER_NAME", EXTRACTOS_FOLDER_NAME).strip()
        or EXTRACTOS_FOLDER_NAME
    )
    return _dedupe_names_preserve_order([asientos, extractos])


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


def _folder_children(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        it
        for it in items
        if "folder" in it
        and "name" in it
        and not str(it.get("name", "")).startswith("~$")
    ]


def _has_named_folder(children: list[dict[str, Any]], name: str) -> bool:
    target = name.casefold()
    for it in children:
        if "folder" not in it:
            continue
        n = str(it.get("name", "")).strip()
        if n.casefold() == target:
            return True
    return False


async def _create_child_folder(
    graph: GraphApiPort,
    site_id: str,
    drive_id: str,
    parent_rel_path: str,
    folder_name: str,
) -> None:
    enc = encode_graph_drive_path(parent_rel_path.strip().strip("/"))
    endpoint = f"/sites/{site_id}/drives/{drive_id}/root:/{enc}:/children"
    body: dict[str, Any] = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "fail",
    }
    _resp, code = await graph.post_json(endpoint, body)
    if code not in (200, 201):
        raise RuntimeError(f"Graph creó carpeta con código inesperado: {code}")


@dataclass
class EnsureAsientosContablesResult:
    clients_base_path: str
    clients_scanned: int
    credit_folders_scanned: int
    folders_created: int
    folders_already_present: int
    subfolder_names: tuple[str, ...] = ()
    errors: list[dict[str, str]] = field(default_factory=list)


async def ensure_asientos_contables_folders(graph: GraphApiPort) -> EnsureAsientosContablesResult:
    site_search = os.getenv("GRAPH_SHAREPOINT_SITE_SEARCH", "").strip()
    drive_name = os.getenv("GRAPH_SHAREPOINT_DRIVE_NAME", "").strip()
    clients_path = os.getenv("GRAPH_CLIENTS_BASE_PATH", "").strip().strip("/")

    if not site_search:
        raise GraphConfigError("Missing environment variable: GRAPH_SHAREPOINT_SITE_SEARCH")
    if not clients_path:
        raise GraphConfigError("Missing environment variable: GRAPH_CLIENTS_BASE_PATH")

    ctx = await resolve_sharepoint_path(graph, site_search, drive_name, clients_path)
    site_id = ctx["site_id"]
    drive_id = ctx["drive_id"]

    subfolder_names = _credit_subfolder_names_to_ensure()
    if not subfolder_names:
        raise ValueError(
            "No hay nombres de subcarpetas para crear: define GRAPH_CREDIT_SUBFOLDERS_TO_ENSURE "
            "o GRAPH_ASIENTOS_CONTABLES_FOLDER_NAME / GRAPH_EXTRACTOS_FOLDER_NAME."
        )

    created = 0
    already = 0
    errors: list[dict[str, str]] = []

    try:
        client_items = _folder_children(await _list_drive_folder_children(graph, site_id, drive_id, clients_path))
    except Exception as exc:
        raise ValueError(
            f"No se pudo listar la ruta base de clientes {clients_path!r}: {exc}"
        ) from exc

    credits_scanned = 0

    for client_it in client_items:
        client_name = str(client_it.get("name", "")).strip()
        if not client_name:
            continue
        client_rel = f"{clients_path}/{client_name}"
        try:
            credit_items = _folder_children(await _list_drive_folder_children(graph, site_id, drive_id, client_rel))
        except Exception as exc:
            errors.append({"path": client_rel, "error": str(exc)})
            logger.warning("ensure credit subfolders: no se listó cliente %s: %s", client_rel, exc)
            continue

        for credit_it in credit_items:
            credit_name = str(credit_it.get("name", "")).strip()
            if not credit_name:
                continue
            credit_rel = f"{client_rel}/{credit_name}"
            credits_scanned += 1
            try:
                inner = await _list_drive_folder_children(graph, site_id, drive_id, credit_rel)
            except Exception as exc:
                errors.append({"path": credit_rel, "error": str(exc)})
                logger.warning("ensure credit subfolders: no se listó crédito %s: %s", credit_rel, exc)
                continue

            for folder_name in subfolder_names:
                if _has_named_folder(inner, folder_name):
                    already += 1
                    continue
                try:
                    await _create_child_folder(graph, site_id, drive_id, credit_rel, folder_name)
                    created += 1
                    logger.info("ensure credit subfolders: creada %s/%s", credit_rel, folder_name)
                except httpx.HTTPStatusError as exc:
                    code = exc.response.status_code if exc.response is not None else 0
                    if code == 409:
                        already += 1
                    else:
                        errors.append({"path": f"{credit_rel}/{folder_name}", "error": str(exc)})
                        logger.warning(
                            "ensure credit subfolders: no se creó %s/%s: %s",
                            credit_rel,
                            folder_name,
                            exc,
                        )
                except Exception as exc:
                    errors.append({"path": f"{credit_rel}/{folder_name}", "error": str(exc)})
                    logger.warning(
                        "ensure credit subfolders: no se creó %s/%s: %s",
                        credit_rel,
                        folder_name,
                        exc,
                    )

    return EnsureAsientosContablesResult(
        clients_base_path=clients_path,
        clients_scanned=len(client_items),
        credit_folders_scanned=credits_scanned,
        folders_created=created,
        folders_already_present=already,
        subfolder_names=tuple(subfolder_names),
        errors=errors,
    )
