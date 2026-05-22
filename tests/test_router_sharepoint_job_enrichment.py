"""GET jobs de SharePoint (notify/merge) devuelven payload enriquecido."""

import asyncio

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

from app.adapters.primary.http.deps import init_graph_client
from app.adapters.primary.http.routers import sharepoint as sharepoint_mod
from app.adapters.primary.http.routers.sharepoint import router


class _MockGraph:
    async def get(self, *a, **k):
        return {}

    async def get_bytes(self, *a, **k):
        return b""

    async def put_bytes(self, *a, **k):
        return {}

    async def delete(self, *a, **k):
        return None

    async def post_json(self, *a, **k):
        return {}, 202


@pytest.fixture  # type: ignore[name-defined]
def client():
    app = FastAPI()
    init_graph_client(_MockGraph())
    app.include_router(router)
    sharepoint_mod._validation_jobs.clear()
    yield TestClient(app, raise_server_exceptions=False)
    sharepoint_mod._validation_jobs.clear()


def test_merge_get_job_enriched_completed(client):
    jid = "m1"
    asyncio.run(
        sharepoint_mod._set_job(
            jid,
            {
                "job_id": jid,
                "type": "merge_composite_validado_pdfs",
                "status": "completed",
                "elapsed_ms": 100.0,
                "result": {
                    "status": "ok",
                    "message": "Ejecutado con éxito",
                    "outputs": [{"id_pago": "a", "output_relative_path": "out/x.pdf"}],
                    "skipped": [],
                },
                "error": None,
            },
        )
    )
    res = client.get(f"/graph/sharepoint/merge-composite-validado-pdfs/jobs/{jid}")
    assert res.status_code == 200
    b = res.json()
    assert b["severity"] == "success"
    assert "user_message" in b


def test_notify_get_job_enriched_failed_string_stored(client):
    jid = "n1"
    asyncio.run(
        sharepoint_mod._set_job(
            jid,
            {
                "job_id": jid,
                "type": "notify_validar_extractos",
                "status": "failed",
                "result": None,
                "error": "Faltan columnas requeridas: Estado línea",
            },
        )
    )
    res = client.get(f"/graph/sharepoint/notify-validar-extractos-email/jobs/{jid}")
    assert res.status_code == 200
    b = res.json()
    assert b["severity"] == "error"
    assert isinstance(b["error"], dict)
    assert b["error"]["message"] == "Faltan columnas requeridas: Estado línea"
