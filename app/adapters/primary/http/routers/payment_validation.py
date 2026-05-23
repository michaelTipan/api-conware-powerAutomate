import logging
import asyncio
from datetime import date, datetime, timezone
from time import perf_counter
from typing import Any

from fastapi import APIRouter, BackgroundTasks, HTTPException
from pydantic import BaseModel

from app.adapters.primary.http.deps import GraphClientDep
from app.application.job_status_enrichment import enrich_job_for_http_response
from app.application.job_manager import JobManager
from app.application.use_cases.payment_validation_generate import generate_payment_validation
from app.application.use_cases.payment_validation_finalize import finalize_payment_validation
from app.application.use_cases.setup_ibr_workbook import (
    IbrWorkbookSetupError,
    setup_ibr_workbook,
)
from app.application.use_cases.setup_payment_followup_workbooks import (
    PaymentFollowupSetupError,
    setup_payment_followup_workbooks,
)
from app.application.use_cases.setup_merge_control_workbook import (
    MergeControlSetupError,
    setup_merge_control_workbook,
)
from app.application.use_cases.amortization_fill_apply import run_amortization_fill_apply
from app.application.use_cases.amortization_fill_dry_run import run_amortization_fill_dry_run
from app.domain.exceptions import GraphConfigError

router = APIRouter(prefix="/graph/sharepoint/payment-validation", tags=["payment-validation"])
logger = logging.getLogger(__name__)


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


# ─── Request Bodies ──────────────────────────────────────────────────────────

class GenerateRequest(BaseModel):
    process_date: str | None = None
    source_file_path: str | None = None
    force: bool = False

class FinalizeRequest(BaseModel):
    validation_file: str | None = None
    validation_file_path: str | None = None
    process_date: str | None = None


class PaymentFollowupSetupRequest(BaseModel):
    force_recreate: bool = False


class IbrWorkbookSetupRequest(BaseModel):
    force_recreate: bool = False


class AmortizationDryRunRequest(BaseModel):
    report_date_iso: str | None = None
    merge_manifest_path: str | None = None
    historical_file_path: str | None = None


# ─── Background Tasks ─────────────────────────────────────────────────────────

async def _run_generate_job(job_id: str, graph: GraphClientDep, process_date: date) -> None:
    jm = JobManager()
    await jm.set_job(job_id, {
        "status": "running",
        "started_at": _utc_now_iso(),
        "updated_at": _utc_now_iso(),
    })
    logger.info("job %s: generate_payment_validation iniciado", job_id)
    started = perf_counter()
    try:
        result = await generate_payment_validation(graph, process_date)
        elapsed_ms = round((perf_counter() - started) * 1000, 2)
        await jm.set_job(job_id, {
            "status": "completed",
            "finished_at": _utc_now_iso(),
            "updated_at": _utc_now_iso(),
            "result": {
                **result,
                "process_date": process_date.isoformat(),
                "elapsed_ms": elapsed_ms,
            },
        })
        logger.info("job %s: completado en %.2fms", job_id, elapsed_ms)
    except Exception as exc:
        await jm.set_job(job_id, {
            "status": "failed",
            "finished_at": _utc_now_iso(),
            "updated_at": _utc_now_iso(),
            "error": {"type": type(exc).__name__, "message": str(exc)},
        })
        logger.error("job %s: falló con %s: %s", job_id, type(exc).__name__, exc)
    finally:
        jm.finish_generate()


async def _run_finalize_job(
    job_id: str,
    graph: GraphClientDep,
    validation_file: str | None,
    validation_file_path: str | None,
    process_date: date,
) -> None:
    jm = JobManager()
    await jm.set_job(job_id, {
        "status": "running",
        "started_at": _utc_now_iso(),
        "updated_at": _utc_now_iso(),
    })
    logger.info("job %s: finalize_payment_validation iniciado", job_id)
    started = perf_counter()
    try:
        result = await finalize_payment_validation(
            graph,
            validation_file=validation_file,
            validation_file_path=validation_file_path,
            process_date=process_date,
        )
        elapsed_ms = round((perf_counter() - started) * 1000, 2)
        await jm.set_job(job_id, {
            "status": "completed",
            "finished_at": _utc_now_iso(),
            "updated_at": _utc_now_iso(),
            "result": {**result, "elapsed_ms": elapsed_ms},
        })
        logger.info("job %s: completado en %.2fms", job_id, elapsed_ms)
    except Exception as exc:
        await jm.set_job(job_id, {
            "status": "failed",
            "finished_at": _utc_now_iso(),
            "updated_at": _utc_now_iso(),
            "error": {"type": type(exc).__name__, "message": str(exc)},
        })
        logger.error("job %s: falló con %s: %s", job_id, type(exc).__name__, exc)
    finally:
        jm.finish_finalize()


async def _run_amortization_dry_run_job(
    job_id: str,
    graph: GraphClientDep,
    *,
    report_date_iso: str | None,
    merge_manifest_path: str | None,
    historical_file_path: str | None,
) -> None:
    jm = JobManager()
    await jm.set_job(
        job_id,
        {
            "status": "running",
            "started_at": _utc_now_iso(),
            "updated_at": _utc_now_iso(),
        },
    )
    logger.info("job %s: amortization_dry_run iniciado", job_id)
    started = perf_counter()
    try:
        result = await run_amortization_fill_dry_run(
            graph,
            report_date_iso=report_date_iso,
            merge_manifest_path=merge_manifest_path,
            historical_file_path=historical_file_path,
        )
        elapsed_ms = round((perf_counter() - started) * 1000, 2)
        await jm.set_job(
            job_id,
            {
                "status": "completed",
                "finished_at": _utc_now_iso(),
                "updated_at": _utc_now_iso(),
                "result": {**result, "elapsed_ms": elapsed_ms},
                "error": None,
            },
        )
        logger.info("job %s: amortization_dry_run completado en %.2fms", job_id, elapsed_ms)
    except ValueError as exc:
        msg = str(exc)
        code = msg.split("|", 1)[0].strip() if "|" in msg else msg.strip()
        await jm.set_job(
            job_id,
            {
                "status": "failed",
                "finished_at": _utc_now_iso(),
                "updated_at": _utc_now_iso(),
                "result": None,
                "error": {
                    "type": "ValueError",
                    "message": msg,
                    "error_code": code,
                },
            },
        )
        logger.warning("job %s: amortization_dry_run falló (validación): %s", job_id, msg)
    except GraphConfigError as exc:
        await jm.set_job(
            job_id,
            {
                "status": "failed",
                "finished_at": _utc_now_iso(),
                "updated_at": _utc_now_iso(),
                "result": None,
                "error": {
                    "type": "GraphConfigError",
                    "message": str(exc),
                    "error_code": "graph_config_error",
                },
            },
        )
        logger.error("job %s: amortization_dry_run config: %s", job_id, exc)
    except Exception as exc:
        await jm.set_job(
            job_id,
            {
                "status": "failed",
                "finished_at": _utc_now_iso(),
                "updated_at": _utc_now_iso(),
                "result": None,
                "error": {"type": type(exc).__name__, "message": str(exc)},
            },
        )
        logger.exception("job %s: amortization_dry_run falló: %s", job_id, exc)


async def _run_amortization_apply_job(
    job_id: str,
    graph: GraphClientDep,
    *,
    report_date_iso: str | None,
    merge_manifest_path: str | None,
    historical_file_path: str | None,
) -> None:
    jm = JobManager()
    await jm.set_job(
        job_id,
        {
            "status": "running",
            "started_at": _utc_now_iso(),
            "updated_at": _utc_now_iso(),
        },
    )
    logger.info("job %s: amortization_apply iniciado", job_id)
    started = perf_counter()
    try:
        result = await run_amortization_fill_apply(
            graph,
            report_date_iso=report_date_iso,
            merge_manifest_path=merge_manifest_path,
            historical_file_path=historical_file_path,
        )
        elapsed_ms = round((perf_counter() - started) * 1000, 2)
        await jm.set_job(
            job_id,
            {
                "status": "completed",
                "finished_at": _utc_now_iso(),
                "updated_at": _utc_now_iso(),
                "result": {**result, "elapsed_ms": elapsed_ms},
                "error": None,
            },
        )
        logger.info("job %s: amortization_apply completado en %.2fms", job_id, elapsed_ms)
    except ValueError as exc:
        msg = str(exc)
        code = msg.split("|", 1)[0].strip() if "|" in msg else msg.strip()
        await jm.set_job(
            job_id,
            {
                "status": "failed",
                "finished_at": _utc_now_iso(),
                "updated_at": _utc_now_iso(),
                "result": None,
                "error": {
                    "type": "ValueError",
                    "message": msg,
                    "error_code": code,
                },
            },
        )
    except GraphConfigError as exc:
        await jm.set_job(
            job_id,
            {
                "status": "failed",
                "finished_at": _utc_now_iso(),
                "updated_at": _utc_now_iso(),
                "result": None,
                "error": {
                    "type": "GraphConfigError",
                    "message": str(exc),
                    "error_code": "graph_config_error",
                },
            },
        )
    except Exception as exc:
        await jm.set_job(
            job_id,
            {
                "status": "failed",
                "finished_at": _utc_now_iso(),
                "updated_at": _utc_now_iso(),
                "result": None,
                "error": {"type": type(exc).__name__, "message": str(exc)},
            },
        )
        logger.exception("job %s: amortization_apply falló: %s", job_id, exc)


# ─── Endpoints ────────────────────────────────────────────────────────────────

@router.post("/generate/queue", status_code=202)
async def queue_generate(
    graph: GraphClientDep,
    background_tasks: BackgroundTasks,
    body: GenerateRequest = None,
) -> dict[str, Any]:
    """
    Encola la generación del Excel de revisión de pagos.
    Power Automate debe llamar este endpoint y luego consultar /jobs/{job_id}.
    """
    jm = JobManager()
    if not jm.try_start_generate():
        raise HTTPException(
            status_code=409,
            detail="Ya existe un proceso generate o finalize activo. Consulta /jobs/{job_id}."
        )

    body = body or GenerateRequest()
    try:
        pd_str = body.process_date or datetime.now(timezone.utc).strftime("%Y-%m-%d")
        process_date = date.fromisoformat(pd_str)
    except ValueError:
        jm.finish_generate()
        raise HTTPException(status_code=422, detail="process_date inválido. Usar formato YYYY-MM-DD.")

    # Crear el job en estado queued
    import uuid
    job_id = str(uuid.uuid4())
    await jm.set_job(job_id, {
        "job_id": job_id,
        "type": "generate",
        "status": "queued",
        "queued_at": _utc_now_iso(),
        "updated_at": _utc_now_iso(),
    })

    background_tasks.add_task(_run_generate_job, job_id, graph, process_date)
    logger.info("job %s: generate encolado", job_id)

    return {"job_id": job_id, "status": "queued"}


@router.post("/finalize/queue", status_code=202)
async def queue_finalize(
    graph: GraphClientDep,
    background_tasks: BackgroundTasks,
    body: FinalizeRequest = None,
) -> dict[str, Any]:
    """
    Encola la finalización del flujo de validación de pagos.
    Power Automate debe llamar este endpoint después de que la secretaria
    complete el Excel de revisión.

    Nota: el endpoint legacy `validate-payment-report` sigue coexistiendo
    temporalmente en `sharepoint.py` y NO se consolida todavía.
    """
    jm = JobManager()
    if not jm.try_start_finalize():
        raise HTTPException(
            status_code=409,
            detail="Ya existe un proceso generate o finalize activo. Consulta /jobs/{job_id}."
        )

    body = body or FinalizeRequest()
    validation_file = body.validation_file or None
    validation_file_path = body.validation_file_path or None

    try:
        pd_str = body.process_date or datetime.now(timezone.utc).strftime("%Y-%m-%d")
        process_date = date.fromisoformat(pd_str)
    except ValueError:
        jm.finish_finalize()
        raise HTTPException(status_code=422, detail="process_date inválido. Usar formato YYYY-MM-DD.")

    import uuid
    job_id = str(uuid.uuid4())
    await jm.set_job(job_id, {
        "job_id": job_id,
        "type": "finalize",
        "status": "queued",
        "queued_at": _utc_now_iso(),
        "updated_at": _utc_now_iso(),
    })

    background_tasks.add_task(
        _run_finalize_job,
        job_id,
        graph,
        validation_file,
        validation_file_path,
        process_date,
    )
    logger.info("job %s: finalize encolado", job_id)

    return {"job_id": job_id, "status": "queued"}


@router.post("/amortization/dry-run/queue", status_code=202)
async def queue_amortization_dry_run(
    graph: GraphClientDep,
    background_tasks: BackgroundTasks,
    body: AmortizationDryRunRequest | None = None,
) -> dict[str, Any]:
    """
    Encola análisis preliminar de amortización (dry-run) contra SharePoint.
    No escribe tablas ni mueve asientos. Consulte ``GET /jobs/{job_id}`` para el resultado.
    """
    payload = body or AmortizationDryRunRequest()
    manifest_path = (payload.merge_manifest_path or "").strip() or None
    report_date = (payload.report_date_iso or "").strip() or None
    historical_path = (payload.historical_file_path or "").strip() or None

    if not manifest_path and not report_date:
        raise HTTPException(
            status_code=422,
            detail={
                "error_code": "amortization_dry_run_params_required",
                "user_message": (
                    "Debe indicar la fecha del reporte (report_date_iso) o la ruta explícita "
                    "del manifest de Merge (merge_manifest_path)."
                ),
                "next_action": (
                    "Envíe report_date_iso en formato YYYY-MM-DD (fecha del merge_manifest) o "
                    "merge_manifest_path con la ruta completa del JSON en SharePoint."
                ),
            },
        )

    if report_date:
        try:
            date.fromisoformat(report_date)
        except ValueError as exc:
            raise HTTPException(
                status_code=422,
                detail={
                    "error_code": "invalid_report_date_iso",
                    "user_message": "report_date_iso no es una fecha válida.",
                    "next_action": "Use formato YYYY-MM-DD, por ejemplo 2026-05-15.",
                },
            ) from exc

    import uuid

    job_id = str(uuid.uuid4())
    jm = JobManager()
    await jm.set_job(
        job_id,
        {
            "job_id": job_id,
            "type": "amortization_dry_run",
            "status": "queued",
            "queued_at": _utc_now_iso(),
            "updated_at": _utc_now_iso(),
            "request": {
                "report_date_iso": report_date,
                "merge_manifest_path": manifest_path,
                "historical_file_path": historical_path,
            },
        },
    )

    background_tasks.add_task(
        _run_amortization_dry_run_job,
        job_id,
        graph,
        report_date_iso=report_date,
        merge_manifest_path=manifest_path,
        historical_file_path=historical_path,
    )
    logger.info("job %s: amortization_dry_run encolado", job_id)

    return {"job_id": job_id, "status": "queued"}


@router.post("/amortization/apply/queue", status_code=202)
async def queue_amortization_apply(
    graph: GraphClientDep,
    background_tasks: BackgroundTasks,
    body: AmortizationDryRunRequest | None = None,
) -> dict[str, Any]:
    """
    Encola apply real de amortización (preflight dry-run + escritura en SharePoint).
    Consulte ``GET /jobs/{job_id}`` para el resultado.
    """
    payload = body or AmortizationDryRunRequest()
    manifest_path = (payload.merge_manifest_path or "").strip() or None
    report_date = (payload.report_date_iso or "").strip() or None
    historical_path = (payload.historical_file_path or "").strip() or None

    if not manifest_path and not report_date:
        raise HTTPException(
            status_code=422,
            detail={
                "error_code": "amortization_apply_params_required",
                "user_message": (
                    "Debe indicar report_date_iso o merge_manifest_path para apply."
                ),
            },
        )

    if report_date:
        try:
            date.fromisoformat(report_date)
        except ValueError as exc:
            raise HTTPException(
                status_code=422,
                detail={"error_code": "invalid_report_date_iso"},
            ) from exc

    import uuid

    job_id = str(uuid.uuid4())
    jm = JobManager()
    await jm.set_job(
        job_id,
        {
            "job_id": job_id,
            "type": "amortization_apply",
            "status": "queued",
            "queued_at": _utc_now_iso(),
            "updated_at": _utc_now_iso(),
            "request": {
                "report_date_iso": report_date,
                "merge_manifest_path": manifest_path,
                "historical_file_path": historical_path,
            },
        },
    )

    background_tasks.add_task(
        _run_amortization_apply_job,
        job_id,
        graph,
        report_date_iso=report_date,
        merge_manifest_path=manifest_path,
        historical_file_path=historical_path,
    )
    logger.info("job %s: amortization_apply encolado", job_id)

    return {"job_id": job_id, "status": "queued"}


@router.get("/jobs/{job_id}")
async def get_job_status(job_id: str) -> dict[str, Any]:
    """
    Consulta el estado de un job de payment-validation.
    Devuelve 404 si el job no existe.
    """
    jm = JobManager()
    job = jm.get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail=f"Job {job_id!r} no encontrado.")
    return enrich_job_for_http_response(job)


@router.post("/setup/merge-control-workbook")
async def post_setup_merge_control_workbook(graph: GraphClientDep) -> dict[str, Any]:
    """
    Crea de forma idempotente ``control_merge_pdfs.xlsx`` en la carpeta 00 CONTROL
    (no sobrescribe si ya existe). Temporal hasta integración Notify/Merge.
    """
    try:
        return await setup_merge_control_workbook(graph)
    except MergeControlSetupError as exc:
        raise HTTPException(
            status_code=exc.http_status,
            detail={
                "user_message": exc.user_message,
                "next_action": exc.next_action,
                "technical_message": exc.technical_message,
            },
        ) from exc
    except GraphConfigError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@router.post("/setup/payment-followup-workbooks")
async def post_setup_payment_followup_workbooks(
    graph: GraphClientDep,
    body: PaymentFollowupSetupRequest | None = None,
) -> dict[str, Any]:
    """
    Crea bandejas operativas ``pagos_adelantados.xlsx`` y ``pagos_incompletos.xlsx``
    (hojas Pendientes + Historico). Por defecto no sobrescribe; ``force_recreate`` reemplaza.
    """
    payload = body or PaymentFollowupSetupRequest()
    try:
        return await setup_payment_followup_workbooks(
            graph,
            force_recreate=payload.force_recreate,
        )
    except PaymentFollowupSetupError as exc:
        raise HTTPException(
            status_code=exc.http_status,
            detail={
                "user_message": exc.user_message,
                "next_action": exc.next_action,
                "technical_message": exc.technical_message,
            },
        ) from exc
    except GraphConfigError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@router.post("/setup/ibr-workbook")
async def post_setup_ibr_workbook(
    graph: GraphClientDep,
    body: IbrWorkbookSetupRequest | None = None,
) -> dict[str, Any]:
    """
    Crea o repara idempotentemente ``IBR_DIARIO.xlsx`` (hoja IBR: Inicio, Fin, Valor).
    Encabezados bloqueados; filas de datos editables. ``force_recreate`` reemplaza el archivo.
    """
    payload = body or IbrWorkbookSetupRequest()
    try:
        return await setup_ibr_workbook(graph, force_recreate=payload.force_recreate)
    except IbrWorkbookSetupError as exc:
        raise HTTPException(
            status_code=exc.http_status,
            detail={
                "user_message": exc.user_message,
                "next_action": exc.next_action,
                "technical_message": exc.technical_message,
            },
        ) from exc
    except GraphConfigError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
