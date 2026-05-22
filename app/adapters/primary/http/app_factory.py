from fastapi import FastAPI

from app.adapters.primary.http.deps import init_graph_client
from app.adapters.primary.http.routers import excel, graph, health, sharepoint, payment_validation
from app.adapters.secondary.ms_graph_client import MsGraphClient
from app.logging_config import configure_logging


def create_app() -> FastAPI:
    configure_logging()
    app = FastAPI(title="Excel Parser API", version="2.0.0")
    init_graph_client(MsGraphClient())
    app.include_router(health.router)
    app.include_router(excel.router)
    app.include_router(graph.router)
    app.include_router(sharepoint.router)
    app.include_router(payment_validation.router)
    return app
