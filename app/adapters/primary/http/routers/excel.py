from fastapi import APIRouter, HTTPException

from app.adapters.primary.http.deps import GraphClientDep
from app.application.use_cases.excel import run_parse_excel
from app.application.use_cases.sharepoint_from_env import parse_configured_excel
from app.domain.exceptions import GraphConfigError
from app.models import ParseExcelRequest, ParseExcelResponse

router = APIRouter(tags=["excel"])


@router.post("/parse-excel", response_model=ParseExcelResponse)
async def parse_excel(payload: ParseExcelRequest) -> ParseExcelResponse:
    try:
        return run_parse_excel(payload)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid Excel payload: {exc}") from exc


@router.post("/validate-excel", response_model=ParseExcelResponse, deprecated=True)
async def validate_excel_alias(payload: ParseExcelRequest) -> ParseExcelResponse:
    """Alias por compatibilidad; preferir POST /parse-excel."""
    return await parse_excel(payload)


@router.post(
    "/graph/sharepoint/parse-excel-from-env",
    response_model=ParseExcelResponse,
    tags=["sharepoint"],
)
async def graph_sharepoint_parse_excel_from_env(
    graph: GraphClientDep,
    column_letter: str = "C",
    run_id: str = "",
) -> ParseExcelResponse:
    try:
        return await parse_configured_excel(graph, column_letter, run_id)
    except GraphConfigError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"SharePoint/Excel error: {exc}") from exc
