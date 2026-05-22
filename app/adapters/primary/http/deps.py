from typing import Annotated

from fastapi import Depends

from app.adapters.secondary.ms_graph_client import MsGraphClient

_client: MsGraphClient | None = None


def init_graph_client(client: MsGraphClient) -> None:
    global _client
    _client = client


def get_graph_client() -> MsGraphClient:
    if _client is None:
        raise RuntimeError("Graph client not initialized (call init_graph_client in app factory).")
    return _client


GraphClientDep = Annotated[MsGraphClient, Depends(get_graph_client)]
