import base64
import logging
import re
import time

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.encoders import jsonable_encoder
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse

from app.excel_parser import parse_excel_from_bytes
from app.models import ParseExcelJsonBody, ParseExcelResponse


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
)
logger = logging.getLogger(__name__)

app = FastAPI(title="Excel Parser API", version="2.0.0")


@app.exception_handler(RequestValidationError)
async def validation_exception_handler(
    request: Request,
    exc: RequestValidationError,
) -> JSONResponse:
    """422: body JSON inválido (p. ej. falta excel_base64). Log + detalle para depurar."""
    errs = exc.errors()
    body_preview = ""
    try:
        body_bytes = await request.body()
        body_preview = body_bytes[:500].decode("utf-8", errors="replace")
    except Exception as read_exc:
        body_preview = f"(body no disponible: {read_exc})"

    logger.warning(
        "validation_failed path=%s method=%s errors=%s body_preview500=%r",
        request.url.path,
        request.method,
        errs,
        body_preview,
    )
    return JSONResponse(
        status_code=422,
        content=jsonable_encoder(
            {
                "status": "error",
                "error": "validation_error",
                "message": "El JSON no cumple el esquema esperado.",
                "details": errs,
                "hint": "Se espera JSON con claves: excel_base64 (string obligatorio), "
                "file_name (opcional), run_id (opcional). Content-Type: application/json.",
            }
        ),
    )


def _normalize_base64_string(raw: str) -> tuple[str, bool]:
    """Quita espacios/saltos típicos de Power Automate; devuelve (texto, hubo_cambio)."""
    s = raw.strip()
    s = re.sub(r"\s+", "", s)
    changed = s != raw.strip()
    return s, changed


def _decode_excel_base64(excel_base64: str) -> bytes:
    """Decodifica Base64 con logs en cada fallo (incl. espacios y padding típicos de Power Automate)."""
    normalized, was_normalized = _normalize_base64_string(excel_base64)
    logger.info(
        "parse_excel_json: excel_base64 len_in=%s len_norm=%s normalized_whitespace=%s",
        len(excel_base64),
        len(normalized),
        was_normalized,
    )

    if not normalized:
        logger.error("parse_excel_json: excel_base64 vacío tras normalizar")
        raise ValueError("excel_base64 vacío")

    last_error: Exception | None = None
    for label, data in [
        (
            "b64decode_validate_true",
            lambda: base64.b64decode(normalized, validate=True),
        ),
        (
            "b64decode_padded",
            lambda: base64.b64decode(
                normalized + "=" * ((-len(normalized)) % 4),
                validate=False,
            ),
        ),
        (
            "urlsafe_b64decode_padded",
            lambda: base64.urlsafe_b64decode(
                normalized + "=" * ((-len(normalized)) % 4),
            ),
        ),
    ]:
        try:
            out = data()
            logger.info(
                "parse_excel_json: base64 OK method=%s -> bytes=%s",
                label,
                len(out),
            )
            return out
        except Exception as e:
            last_error = e
            logger.warning(
                "parse_excel_json: base64 fallo method=%s (%s: %s)",
                label,
                type(e).__name__,
                e,
            )
    assert last_error is not None
    logger.error(
        "parse_excel_json: todos los intentos base64 fallaron: %s",
        last_error,
    )
    raise ValueError(
        f"No se pudo decodificar excel_base64: {last_error}"
    ) from last_error


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/parse-excel", response_model=ParseExcelResponse)
async def parse_excel(
    file: UploadFile = File(..., description="Archivo Excel (.xlsx)"),
    run_id: str = Form(default=""),
) -> ParseExcelResponse:
    """Entrada: archivo Excel. Salida: todas las filas y columnas de la primera hoja."""
    raw = await file.read()
    if not raw:
        logger.error("parse_excel: archivo vacío upload filename=%r", file.filename)
        raise HTTPException(
            status_code=400,
            detail={
                "status": "error",
                "error": "empty_file",
                "message": "Archivo vacío.",
            },
        )

    started_at = time.perf_counter()
    try:
        columns, rows = parse_excel_from_bytes(raw)
    except Exception as exc:
        logger.exception(
            "parse_excel: fallo al leer Excel filename=%r bytes=%s",
            file.filename,
            len(raw),
        )
        raise HTTPException(
            status_code=400,
            detail={
                "status": "error",
                "error": "excel_parse_failed",
                "message": str(exc),
                "exception_type": type(exc).__name__,
            },
        ) from exc

    elapsed_ms = int((time.perf_counter() - started_at) * 1000)
    name = file.filename or "upload.xlsx"
    return ParseExcelResponse(
        status="ok",
        run_id=run_id,
        file_name=name,
        row_count=len(rows),
        columns=columns,
        rows=rows,
        elapsed_ms=elapsed_ms,
    )


@app.post("/parse-excel-json", response_model=ParseExcelResponse)
async def parse_excel_json(payload: ParseExcelJsonBody) -> ParseExcelResponse:
    """Misma salida que /parse-excel, pero entrada JSON con el Excel en Base64 (útil para Power Automate HTTP)."""
    try:
        raw = _decode_excel_base64(payload.excel_base64)
    except ValueError as exc:
        logger.error("parse_excel_json: ValueError en decode: %s", exc)
        raise HTTPException(
            status_code=400,
            detail={
                "status": "error",
                "error": "invalid_base64",
                "message": str(exc),
            },
        ) from exc

    if not raw:
        logger.error("parse_excel_json: bytes vacíos tras base64 decode")
        raise HTTPException(
            status_code=400,
            detail={
                "status": "error",
                "error": "empty_payload",
                "message": "excel_base64 decodifica a 0 bytes.",
            },
        )

    started_at = time.perf_counter()
    try:
        columns, rows = parse_excel_from_bytes(raw)
    except Exception as exc:
        logger.exception(
            "parse_excel_json: parse_excel_from_bytes falló file_name=%r bytes=%s",
            payload.file_name,
            len(raw),
        )
        raise HTTPException(
            status_code=400,
            detail={
                "status": "error",
                "error": "excel_parse_failed",
                "message": str(exc),
                "exception_type": type(exc).__name__,
            },
        ) from exc

    elapsed_ms = int((time.perf_counter() - started_at) * 1000)
    return ParseExcelResponse(
        status="ok",
        run_id=payload.run_id,
        file_name=payload.file_name or "upload.xlsx",
        row_count=len(rows),
        columns=columns,
        rows=rows,
        elapsed_ms=elapsed_ms,
    )
