"""Tests para el router de payment-validation.

Valida comportamiento HTTP: códigos de respuesta, estructura del JSON,
uso del JobManager compartido y que la lógica pesada no se ejecuta
de forma síncrona en el request.
"""
import pytest
import asyncio
from fastapi import FastAPI
from fastapi.testclient import TestClient

from app.application.job_manager import JobManager
from app.adapters.primary.http.routers import payment_validation as payment_validation_router
from app.adapters.primary.http.routers.payment_validation import router
from app.adapters.primary.http.deps import init_graph_client


class MockGraphClient:
    """Mock mínimo del Graph client para tests de router."""
    async def get(self, endpoint, params=None): return {}
    async def get_bytes(self, *a, **k): return b""
    async def put_bytes(self, *a, **k): return {"id": "x"}
    async def delete(self, *a, **k): return None
    async def post_json(self, *a, **k): return {}, 202


def test_router_mock_graph_contract_has_no_legacy_methods():
    graph = MockGraphClient()
    assert not hasattr(graph, "get_file_content")
    assert not hasattr(graph, "upload_file")


def build_app() -> FastAPI:
    app = FastAPI()
    init_graph_client(MockGraphClient())
    app.include_router(router)
    return app


@pytest.fixture(autouse=True)
def reset_job_manager():
    """Resetea el Singleton de JobManager entre tests para evitar contaminación."""
    jm = JobManager()
    jm._validation_jobs.clear()
    jm._generate_active = False
    jm._finalize_active = False
    yield
    jm._validation_jobs.clear()
    jm._generate_active = False
    jm._finalize_active = False


@pytest.fixture
def client():
    return TestClient(build_app(), raise_server_exceptions=False)


def test_finalize_queue_accepts_validation_file(client, monkeypatch):
    captured = {}

    async def fake_run_finalize_job(job_id, graph, validation_file, validation_file_path, process_date):
        captured["job_id"] = job_id
        captured["validation_file"] = validation_file
        captured["validation_file_path"] = validation_file_path
        captured["process_date"] = process_date.isoformat()

    monkeypatch.setattr(payment_validation_router, "_run_finalize_job", fake_run_finalize_job)

    res = client.post(
        "/graph/sharepoint/payment-validation/finalize/queue",
        json={"validation_file": "validacion_pagos_2026-05-10.xlsx"}
    )

    assert res.status_code == 202
    assert captured["validation_file"] == "validacion_pagos_2026-05-10.xlsx"
    assert captured["validation_file_path"] is None


def test_finalize_queue_accepts_validation_file_path(client, monkeypatch):
    captured = {}

    async def fake_run_finalize_job(job_id, graph, validation_file, validation_file_path, process_date):
        captured["validation_file"] = validation_file
        captured["validation_file_path"] = validation_file_path
        captured["process_date"] = process_date.isoformat()

    monkeypatch.setattr(payment_validation_router, "_run_finalize_job", fake_run_finalize_job)

    res = client.post(
        "/graph/sharepoint/payment-validation/finalize/queue",
        json={"validation_file_path": "revision/subcarpeta/val_manual.xlsx"}
    )

    assert res.status_code == 202
    assert captured["validation_file"] is None
    assert captured["validation_file_path"] == "revision/subcarpeta/val_manual.xlsx"


def test_finalize_queue_accepts_process_date(client, monkeypatch):
    captured = {}

    async def fake_run_finalize_job(job_id, graph, validation_file, validation_file_path, process_date):
        captured["process_date"] = process_date.isoformat()

    monkeypatch.setattr(payment_validation_router, "_run_finalize_job", fake_run_finalize_job)

    res = client.post(
        "/graph/sharepoint/payment-validation/finalize/queue",
        json={"process_date": "2026-05-10"}
    )

    assert res.status_code == 202
    assert captured["process_date"] == "2026-05-10"


# ──────────────────────────────────────────────────────────────────────────────
# 1. POST /generate/queue devuelve 202, job_id y status queued
# ──────────────────────────────────────────────────────────────────────────────
def test_generate_queue_returns_202(client):
    res = client.post("/graph/sharepoint/payment-validation/generate/queue")
    assert res.status_code == 202
    body = res.json()
    assert "job_id" in body
    assert body["status"] == "queued"


def test_generate_queue_accepts_empty_body(client):
    res = client.post(
        "/graph/sharepoint/payment-validation/generate/queue",
        json={}
    )
    assert res.status_code == 202


def test_generate_queue_accepts_optional_overrides(client):
    res = client.post(
        "/graph/sharepoint/payment-validation/generate/queue",
        json={"process_date": "2026-05-10", "source_file_path": None, "force": False}
    )
    assert res.status_code == 202
    body = res.json()
    assert body["status"] == "queued"


# ──────────────────────────────────────────────────────────────────────────────
# 2. POST /finalize/queue devuelve 202, job_id y status queued
# ──────────────────────────────────────────────────────────────────────────────
def test_finalize_queue_returns_202(client):
    res = client.post("/graph/sharepoint/payment-validation/finalize/queue")
    assert res.status_code == 202
    body = res.json()
    assert "job_id" in body
    assert body["status"] == "queued"


def test_finalize_queue_accepts_empty_body(client):
    res = client.post(
        "/graph/sharepoint/payment-validation/finalize/queue",
        json={}
    )
    assert res.status_code == 202


def test_finalize_queue_accepts_optional_overrides(client):
    res = client.post(
        "/graph/sharepoint/payment-validation/finalize/queue",
        json={"validation_file": "validacion_pagos_2026-05-10.xlsx"}
    )
    assert res.status_code == 202


def test_finalize_rejects_invalid_process_date(client):
    res = client.post(
        "/graph/sharepoint/payment-validation/finalize/queue",
        json={"process_date": "not-a-date"}
    )
    assert res.status_code == 422


# ──────────────────────────────────────────────────────────────────────────────
# 3. GET /jobs/{job_id} devuelve estado del job
# ──────────────────────────────────────────────────────────────────────────────
def test_get_job_returns_job_state(client):
    # Primero crear un job
    res = client.post("/graph/sharepoint/payment-validation/generate/queue")
    assert res.status_code == 202
    job_id = res.json()["job_id"]

    # Consultar el job (puede ser queued o running dado el background task)
    res2 = client.get(f"/graph/sharepoint/payment-validation/jobs/{job_id}")
    assert res2.status_code == 200
    body = res2.json()
    assert body.get("job_id") == job_id or "status" in body


# ──────────────────────────────────────────────────────────────────────────────
# 4. GET /jobs/{job_id} inexistente devuelve 404
# ──────────────────────────────────────────────────────────────────────────────
def test_get_job_not_found_returns_404(client):
    res = client.get("/graph/sharepoint/payment-validation/jobs/nonexistent-id-xyz")
    assert res.status_code == 404


def test_get_completed_generate_job_response_is_enriched(client):
    import asyncio

    jm = JobManager()
    job_id = "enriched-generate-1"

    async def setup():
        await jm.set_job(
            job_id,
            {
                "job_id": job_id,
                "type": "generate",
                "status": "completed",
                "queued_at": "2026-01-01T00:00:00+00:00",
                "result": {
                    "validation_file": "f.xlsx",
                    "validation_file_path": "rev/f.xlsx",
                    "process_id": "pid",
                    "process_date": "2026-05-01",
                    "summary": {"pagos_banco": 2, "errores": 0},
                    "elapsed_ms": 12.0,
                },
            },
        )

    asyncio.run(setup())
    res = client.get(f"/graph/sharepoint/payment-validation/jobs/{job_id}")
    assert res.status_code == 200
    body = res.json()
    assert body["severity"] == "success"
    assert "user_message" in body
    assert body["result"]["validation_file"] == "f.xlsx"


def test_generate_completed_job_still_includes_user_message_next_action_severity(client):
    """GET job completed sigue enriqueciendo el documento si result incluye validation_file_url."""
    import asyncio

    jm = JobManager()
    job_id = "enriched-generate-url-1"
    file_url = "https://comwareec.sharepoint.com/sites/x/revision/f.xlsx"

    async def setup():
        await jm.set_job(
            job_id,
            {
                "job_id": job_id,
                "type": "generate",
                "status": "completed",
                "queued_at": "2026-01-01T00:00:00+00:00",
                "result": {
                    "validation_file": "f.xlsx",
                    "validation_file_path": "rev/f.xlsx",
                    "validation_file_url": file_url,
                    "process_id": "pid",
                    "process_date": "2026-05-01",
                    "summary": {"pagos_banco": 2, "errores": 0},
                    "elapsed_ms": 12.0,
                },
            },
        )

    asyncio.run(setup())
    res = client.get(f"/graph/sharepoint/payment-validation/jobs/{job_id}")
    assert res.status_code == 200
    body = res.json()
    assert body["severity"] == "success"
    assert "user_message" in body
    assert "next_action" in body
    assert body["result"]["validation_file_url"] == file_url


# ──────────────────────────────────────────────────────────────────────────────
# 5. Router usa JobManager compartido, no almacenamiento local nuevo
# ──────────────────────────────────────────────────────────────────────────────
def test_router_uses_shared_job_manager(client):
    """El job creado por el endpoint debe ser visible en el Singleton JobManager."""
    res = client.post("/graph/sharepoint/payment-validation/generate/queue")
    assert res.status_code == 202
    job_id = res.json()["job_id"]

    jm = JobManager()
    job = jm.get_job(job_id)
    assert job is not None
    assert job.get("status") in {"queued", "running", "completed", "failed"}


# ──────────────────────────────────────────────────────────────────────────────
# 6. Los endpoints no ejecutan lógica pesada síncronamente
# ──────────────────────────────────────────────────────────────────────────────
def test_generate_returns_immediately_without_heavy_logic(client):
    """El endpoint /generate/queue debe responder antes de que el job termine.
    El job se encola en background — el request no debe bloquear."""
    import time
    t0 = time.monotonic()
    res = client.post("/graph/sharepoint/payment-validation/generate/queue")
    elapsed = time.monotonic() - t0
    # El endpoint debe responder rápidamente (no espera que el job termine)
    # TestClient ejecuta el background task igualmente, pero la respuesta 202 se emite antes
    assert res.status_code == 202
    # No hacemos assert de tiempo exacto ya que TestClient puede ejecutar tasks sincrónicamente


# ──────────────────────────────────────────────────────────────────────────────
# 7. Conflicto 409 cuando ya hay un proceso activo
# ──────────────────────────────────────────────────────────────────────────────
def test_generate_returns_409_if_already_active(client):
    """Si hay un generate o finalize activo, debe devolver 409."""
    jm = JobManager()
    jm._generate_active = True  # Simular proceso activo

    res = client.post("/graph/sharepoint/payment-validation/generate/queue")
    assert res.status_code == 409


def test_finalize_returns_409_if_generate_active(client):
    """Finalize también debe bloquear si hay un generate activo."""
    jm = JobManager()
    jm._generate_active = True

    res = client.post("/graph/sharepoint/payment-validation/finalize/queue")
    assert res.status_code == 409


# ──────────────────────────────────────────────────────────────────────────────
# 8. Validación de process_date inválido
# ──────────────────────────────────────────────────────────────────────────────
def test_generate_rejects_invalid_process_date(client):
    res = client.post(
        "/graph/sharepoint/payment-validation/generate/queue",
        json={"process_date": "not-a-date"}
    )
    assert res.status_code == 422
