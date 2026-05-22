import time
from base64 import b64decode, b64encode

from app.application.sharepoint_resolution import resolve_sharepoint_from_env
from app.domain.services.excel_column import parse_excel_names
from app.domain.ports.graph import GraphApiPort
from app.models import ParseExcelResponse


async def resolve_configured_item(graph: GraphApiPort) -> dict:
    ctx = await resolve_sharepoint_from_env(graph)
    item = await graph.get(
        f"/sites/{ctx['site_id']}/drives/{ctx['drive_id']}/root:/{ctx['path_encoded']}:"
    )
    return {"resolved": ctx, "item": item}


async def download_configured_file_base64(graph: GraphApiPort) -> dict:
    ctx = await resolve_sharepoint_from_env(graph)
    content = await graph.get_bytes(
        f"/sites/{ctx['site_id']}/drives/{ctx['drive_id']}/root:/{ctx['path_encoded']}:/content"
    )
    return {
        "resolved": ctx,
        "content_base64": b64encode(content).decode("utf-8"),
    }


async def upload_configured_file(graph: GraphApiPort, content_base64: str) -> dict:
    ctx = await resolve_sharepoint_from_env(graph)
    file_bytes = b64decode(content_base64)
    endpoint = (
        f"/sites/{ctx['site_id']}/drives/{ctx['drive_id']}/root:/{ctx['path_encoded']}:/content"
    )
    return await graph.put_bytes(endpoint, file_bytes)


async def parse_configured_excel(
    graph: GraphApiPort,
    column_letter: str = "C",
    run_id: str = "",
) -> ParseExcelResponse:
    started_at = time.perf_counter()
    ctx = await resolve_sharepoint_from_env(graph)
    content = await graph.get_bytes(
        f"/sites/{ctx['site_id']}/drives/{ctx['drive_id']}/root:/{ctx['path_encoded']}:/content"
    )
    excel_base64 = b64encode(content).decode("utf-8")
    file_name = ctx["file_path"].rsplit("/", maxsplit=1)[-1]
    unique_names, duplicates, total_rows, values_in_row_order = parse_excel_names(
        excel_base64,
        column_letter,
    )
    elapsed_ms = int((time.perf_counter() - started_at) * 1000)
    return ParseExcelResponse(
        status="ok",
        run_id=run_id,
        file_name=file_name,
        column_letter=column_letter.upper(),
        total_rows=total_rows,
        non_empty_count=len(values_in_row_order),
        unique_count=len(unique_names),
        duplicate_count=len(duplicates),
        unique_values=unique_names,
        values_in_row_order=values_in_row_order,
        duplicates=duplicates,
        elapsed_ms=elapsed_ms,
    )
