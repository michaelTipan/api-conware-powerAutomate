"""Tests HTTP del dry-run de amortización (payment-validation router)."""

import asyncio

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

from app.application.job_manager import JobManager
from app.adapters.primary.http.routers import payment_validation as payment_validation_router
from app.adapters.primary.http.routers.payment_validation import router
from app.adapters.primary.http.deps import init_graph_client
from app.application.job_status_enrichment import enrich_job_for_http_response


class MockGraphClient:
    """Mock Graph: solo lectura; registra put_bytes para asegurar que dry-run no escribe."""

    def __init__(self) -> None:
        self.put_calls: list[tuple] = []

    async def get(self, endpoint, params=None):
        return {"value": [{"id": "site"}, {"id": "drive", "name": "Doc"}]}

    async def get_bytes(self, *args, **kwargs):
        return b""

    async def put_bytes(self, *args, **kwargs):
        self.put_calls.append((args, kwargs))
        return {"id": "x"}

    async def delete(self, *args, **kwargs):
        return None

    async def post_json(self, *args, **kwargs):
        return {}, 202


def build_app() -> FastAPI:
    app = FastAPI()
    init_graph_client(MockGraphClient())
    app.include_router(router)
    return app


@pytest.fixture(autouse=True)
def reset_job_manager():
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


@pytest.fixture
def mock_graph():
    return MockGraphClient()


def test_amortization_dry_run_queue_returns_202(client):
    res = client.post(
        "/graph/sharepoint/payment-validation/amortization/dry-run/queue",
        json={"report_date_iso": "2026-05-15"},
    )
    assert res.status_code == 202
    body = res.json()
    assert body["status"] == "queued"
    assert body.get("job_id")


def test_amortization_dry_run_queue_rejects_missing_params(client):
    res = client.post(
        "/graph/sharepoint/payment-validation/amortization/dry-run/queue",
        json={},
    )
    assert res.status_code == 422
    detail = res.json()["detail"]
    assert detail["error_code"] == "amortization_dry_run_params_required"


def test_amortization_dry_run_job_runs_use_case_with_params(client, monkeypatch):
    captured: dict = {}
    jm = JobManager()

    async def fake_run(
        job_id: str,
        graph,
        *,
        report_date_iso: str | None,
        merge_manifest_path: str | None,
        historical_file_path: str | None,
    ) -> None:
        captured["report_date_iso"] = report_date_iso
        captured["merge_manifest_path"] = merge_manifest_path
        captured["historical_file_path"] = historical_file_path
        await jm.set_job(
            job_id,
            {
                "status": "completed",
                "finished_at": "2026-05-15T12:00:00+00:00",
                "result": {
                    "status": "ok",
                    "mode": "dry_run",
                    "manifest_path": "LOGS/merge_manifest_2026-05-15.json",
                    "items": [],
                    "summary": {"total_events": 0, "total": 0},
                },
            },
        )

    monkeypatch.setattr(
        payment_validation_router,
        "_run_amortization_dry_run_job",
        fake_run,
    )

    res = client.post(
        "/graph/sharepoint/payment-validation/amortization/dry-run/queue",
        json={
            "report_date_iso": "2026-05-15",
            "merge_manifest_path": None,
            "historical_file_path": "HIST/cartera.xlsx",
        },
    )
    job_id = res.json()["job_id"]
    assert captured["report_date_iso"] == "2026-05-15"
    assert captured["merge_manifest_path"] is None
    assert captured["historical_file_path"] == "HIST/cartera.xlsx"

    res2 = client.get(f"/graph/sharepoint/payment-validation/jobs/{job_id}")
    assert res2.status_code == 200
    body = res2.json()
    assert body["status"] == "completed"
    assert body["result"]["mode"] == "dry_run"


def test_amortization_dry_run_does_not_put_to_sharepoint(monkeypatch):
    mock_graph = MockGraphClient()
    app = FastAPI()
    init_graph_client(mock_graph)
    app.include_router(router)
    test_client = TestClient(app, raise_server_exceptions=False)

    async def fake_run_use_case(graph, **kwargs):
        return {
            "status": "ok",
            "mode": "dry_run",
            "manifest_path": "LOGS/x.json",
            "items": [],
            "summary": {"total": 0},
        }

    monkeypatch.setattr(
        payment_validation_router,
        "run_amortization_fill_dry_run",
        fake_run_use_case,
    )

    test_client.post(
        "/graph/sharepoint/payment-validation/amortization/dry-run/queue",
        json={"report_date_iso": "2026-05-15"},
    )
    assert mock_graph.put_calls == []


def test_amortization_dry_run_completed_job_enrichment(client):
    jm = JobManager()
    job_id = "dry-run-enriched-1"
    result = {
        "status": "ok",
        "mode": "dry_run",
        "manifest_path": "LOGS/merge_manifest_2026-05-15.json",
        "manifest_outputs_count": 1,
        "items": [
            {
                "id_pago": "P1",
                "event_index": 1,
                "asiento_pdf_path": "c/a1.pdf",
                "application_status": "WOULD_APPLY",
            },
            {
                "id_pago": "P1",
                "event_index": 2,
                "asiento_pdf_path": "c/a2.pdf",
                "application_status": "WOULD_APPLY",
            },
        ],
        "summary": {"total_events": 2, "total": 2, "would_apply": 2, "errors": 0},
    }

    async def setup():
        await jm.set_job(
            job_id,
            {
                "job_id": job_id,
                "type": "amortization_dry_run",
                "status": "completed",
                "result": result,
            },
        )

    asyncio.run(setup())
    res = client.get(f"/graph/sharepoint/payment-validation/jobs/{job_id}")
    body = res.json()
    assert body["severity"] == "success"
    assert "no se modificó ninguna tabla" in body["user_message"].lower()
    assert body["result"]["summary"]["total_events"] == 2
    assert len(body["result"]["items"]) == 2


def test_amortization_dry_run_enrichment_helper_multi_events():
    raw = {
        "job_id": "x",
        "type": "amortization_dry_run",
        "status": "completed",
        "result": {
            "status": "ok",
            "mode": "dry_run",
            "items": [
                {"event_index": 1, "asiento_pdf_path": "a.pdf"},
                {"event_index": 2, "asiento_pdf_path": "b.pdf"},
            ],
            "summary": {"total_events": 2, "errors": 0},
        },
    }
    out = enrich_job_for_http_response(raw)
    assert out["severity"] == "success"
    assert "análisis preliminar" in out["user_message"].lower()


def test_amortization_dry_run_failed_enrichment_manifest_not_found():
    raw = {
        "job_id": "x",
        "type": "amortization_dry_run",
        "status": "failed",
        "error": {
            "type": "ValueError",
            "message": "merge_manifest_not_found|LOGS/merge_manifest_2026-05-15.json",
        },
    }
    out = enrich_job_for_http_response(raw)
    assert out["severity"] == "error"
    assert out["error"]["error_code"] == "merge_manifest_not_found"
