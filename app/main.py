import base64
import logging
import re
import time

from fastapi import FastAPI, HTTPException, Request
from fastapi.encoders import jsonable_encoder
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse

from app.excel_parser import parse_excel_from_bytes
from app.models import ParseExcelJsonBody, ParseExcelResponse

# Configuración de Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Excel Parser API",
    description="API optimizada para Power Automate (JSON + Base64)",
    version="2.1.0"
)

@app.exception_handler(RequestValidationError)
async def validation_exception_handler(
    request: Request,
    exc: RequestValidationError,
) -> JSONResponse:
    """422: body JSON inválido. Log detallado para depurar errores en Power Automate."""
    errs = exc.errors()
    body_preview = ""
    try:
        body_bytes = await request.body()
        body_preview = body_bytes[:500].decode("utf-8", errors="replace")
    except Exception as read_exc:
        body_preview = f"(body no disponible: {read_exc})"

    logger.warning(
        "validation_failed path=%s errors=%s body_preview500=%r",
        request.url.path,
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
                "hint": "Se espera JSON con: excel_base64 (obligatorio), file_name y run_id.",
            }
        ),
    )

def _normalize_base64_string(raw: str) -> tuple[str, bool]:
    """Quita espacios/saltos típicos de Power Automate."""
    s = raw.strip()
    s = re.sub(r"\s+", "", s)
    changed = s != raw.strip()
    return s, changed

def _decode_excel_base64(excel_base64: str) -> bytes:
    """Decodifica Base64 con múltiples intentos de padding/formato."""
    normalized, _ = _normalize_base64_string(excel_base64)
    
    if not normalized:
        raise ValueError("excel_base64 está vacío")

    last_error: Exception | None = None
    # Intentos de decodificación: Estándar, Con Padding, URL-Safe
    for label, data in [
        ("standard", lambda: base64.b64decode(normalized, validate=True)),
        ("padded", lambda: base64.b64decode(normalized + "=" * ((-len(normalized)) % 4), validate=False)),
        ("urlsafe", lambda: base64.urlsafe_b64decode(normalized + "=" * ((-len(normalized)) % 4))),
    ]:
        try:
            return data()
        except Exception as e:
            last_error = e
            logger.debug("Falló decodificación %s: %s", label, e)

    raise ValueError(f"No se pudo decodificar excel_base64: {last_error}")

@app.get("/health")
async def health() -> dict[str, str]:
    """Endpoint de vitalidad para Render."""
    return {"status": "ok"}

@app.post("/parse-excel-json", response_model=ParseExcelResponse)
async def parse_excel_json(payload: ParseExcelJsonBody) -> ParseExcelResponse:
    """
    Endpoint principal para Power Automate.
    Recibe el Excel en Base64 dentro de un JSON.
    """
    try:
        raw = _decode_excel_base64(payload.excel_base64)
    except ValueError as exc:
        logger.error("Error decodificando Base64: %s", exc)
        raise HTTPException(
            status_code=400,
            detail={"status": "error", "error": "invalid_base64", "message": str(exc)},
        )

    if not raw:
        raise HTTPException(
            status_code=400,
            detail={"status": "error", "error": "empty_payload", "message": "Bytes vacíos."},
        )

    started_at = time.perf_counter()
    try:
        columns, rows = parse_excel_from_bytes(raw)
    except Exception as exc:
        logger.exception("Error al parsear Excel: bytes=%s", len(raw))
        raise HTTPException(
            status_code=400,
            detail={
                "status": "error",
                "error": "excel_parse_failed",
                "message": str(exc),
                "exception_type": type(exc).__name__,
            },
        )

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

