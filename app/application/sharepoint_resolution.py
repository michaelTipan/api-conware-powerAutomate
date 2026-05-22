"""Resolucion de sitio/drive/ruta SharePoint desde configuracion (entorno)."""

import os
from urllib.parse import quote

from app.domain.exceptions import GraphConfigError
from app.domain.ports.graph import GraphApiPort


def encode_graph_drive_path(relative_path: str) -> str:
    trimmed = relative_path.strip().strip("/")
    if not trimmed:
        raise GraphConfigError("Path must not be empty.")
    parts = [p for p in trimmed.split("/") if p]
    return "/".join(quote(part, safe="") for part in parts)


async def resolve_sharepoint_path(client: GraphApiPort, site_search: str, drive_name: str, path: str) -> dict[str, str]:
    if not path:
        raise GraphConfigError("Missing sharepoint path")

    path_encoded = encode_graph_drive_path(path)

    if not site_search:
        raise GraphConfigError("Missing site_search")

    sites = await client.get("/sites", params={"search": site_search})
    values = sites.get("value") or []
    if not values:
        raise GraphConfigError(f"No SharePoint site found for search: {site_search!r}")
    site_id = values[0]["id"]

    drives_resp = await client.get(f"/sites/{site_id}/drives")
    drives = drives_resp.get("value") or []
    if not drives:
        raise GraphConfigError(f"No drives found for site id {site_id!r}")
    if drive_name:
        match = next((d for d in drives if d.get("name") == drive_name), None)
        if not match:
            raise GraphConfigError(
                f"No drive named {drive_name!r}; available: "
                f"{[d.get('name') for d in drives]}"
            )
        drive_id = match["id"]
    else:
        drive_id = drives[0]["id"]

    return {
        "site_id": site_id,
        "drive_id": drive_id,
        "path_encoded": path_encoded,
        "file_path": path,
    }


async def resolve_sharepoint_from_env(client: GraphApiPort) -> dict[str, str]:
    site_search = os.getenv("GRAPH_SHAREPOINT_SITE_SEARCH", "").strip()
    file_path = os.getenv("GRAPH_SHAREPOINT_FILE_PATH", "").strip()
    drive_name = os.getenv("GRAPH_SHAREPOINT_DRIVE_NAME", "").strip()
    
    if not file_path:
        raise GraphConfigError("Missing environment variable: GRAPH_SHAREPOINT_FILE_PATH")
    if not site_search:
        raise GraphConfigError("Missing environment variable: GRAPH_SHAREPOINT_SITE_SEARCH")

    return await resolve_sharepoint_path(client, site_search, drive_name, file_path)
