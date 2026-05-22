from fastapi import APIRouter, HTTPException

from app.adapters.primary.http.deps import GraphClientDep
from app.domain.exceptions import GraphConfigError

router = APIRouter(prefix="/graph", tags=["graph"])


@router.get("/users")
async def graph_users(graph: GraphClientDep, top: int = 10) -> dict:
    try:
        return await graph.get("/users", params={"$top": top})
    except GraphConfigError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Graph request failed: {exc}") from exc


@router.get("/me")
async def graph_me(graph: GraphClientDep) -> dict:
    try:
        return await graph.get("/me")
    except GraphConfigError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Graph request failed: {exc}") from exc
