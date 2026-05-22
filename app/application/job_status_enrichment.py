"""
Enriquecimiento de documentos de job para respuestas HTTP homogéneas (Power Automate).

Solo afecta la salida de GET job; no modifica lógica de negocio ni almacenamiento interno
más allá de devolver una copia enriquecida desde los routers.
"""

from __future__ import annotations

import re
from typing import Any

ENRICHABLE_JOB_TYPES = frozenset(
    {
        "generate",
        "finalize",
        "notify_validar_extractos",
        "merge_composite_validado_pdfs",
    }
)

_UNKNOWN_USER = (
    "El proceso se detuvo por un error que el sistema no pudo clasificar. "
    "No se completó la operación."
)
_UNKNOWN_NEXT = (
    "No vuelva a ejecutar el flujo hasta revisar con soporte. "
    "Indique la fecha del reporte y copie el detalle técnico que aparece en el mismo mensaje de error."
)

_GENERATE_MESSAGES: dict[str, tuple[str, str]] = {
    "review_folder_not_empty": (
        "No se pudo generar el archivo nuevo porque en 01 REVISION todavía hay un Excel de un día anterior.",
        "Mueva o archive los archivos validacion_pagos_*.xlsx viejos en 01 REVISION y vuelva a ejecutar la generación. "
        "Deje solo el reporte del banco actualizado en su carpeta.",
    ),
    "missing_sharepoint_folder": (
        "El sistema no tiene configuradas las rutas de SharePoint para este proceso (sitio, banco o clientes).",
        "Contacte a soporte técnico. No es un error de la secretaría ni del Excel del banco.",
    ),
    "bank_headers_not_found": (
        "El archivo BANCO_BOGOTA.xlsx no tiene las columnas que el sistema espera (fecha, monto/crédito, concepto).",
        "Revise el formato del reporte del banco. Debe coincidir con la plantilla habitual. "
        "Corrija el Excel, vuelva a subirlo y ejecute de nuevo la generación.",
    ),
    "customer_not_found": (
        "En el reporte del banco hay un pago cuyo Concepto no coincide con ninguna carpeta de cliente en SharePoint.",
        "En 01 COMWARE AUTOMATIZACION - INFORMACION CREDITOS CLIENTES, cree o corrija la carpeta del cliente "
        "para que el nombre coincida con el concepto del banco. Vuelva a generar.",
    ),
    "customer_ambiguous": (
        "El concepto del banco coincide con más de una carpeta de cliente; el sistema no sabe cuál usar.",
        "En SharePoint, deje un solo nombre de carpeta por cliente (sin duplicados parecidos). Vuelva a generar.",
    ),
    "credit_folder_not_found": (
        "No se encontró ninguna carpeta de crédito asociada al cliente del pago (estructura de carpetas incompleta).",
        "Dentro de la carpeta del cliente, verifique que existan carpetas de crédito con extractos y tabla de amortización. "
        "Corrija en SharePoint y vuelva a generar.",
    ),
    "extract_not_found": (
        "Falta el PDF del extracto en la carpeta del crédito (o no se detectó con la palabra configurada, por ejemplo Extracto).",
        "Suba el extracto en la carpeta del crédito, en la subcarpeta EXTRACTOS si aplica. Vuelva a generar.",
    ),
    "extract_amount_not_found": (
        "Hay extracto PDF pero el sistema no pudo leer el valor TOTAL A PAGAR.",
        "Use un PDF legible (no escaneado borroso) o el formato de extracto habitual. Vuelva a generar.",
    ),
    "pending_installment_not_found": (
        "En la tabla de amortización del crédito no hay una cuota pendiente clara para aplicar el pago.",
        "Revise el Excel de amortización en la carpeta del crédito (cuotas pendientes marcadas). Corrija y vuelva a generar.",
    ),
    "amortization_table_not_found": (
        "En la carpeta del crédito no hay archivo de tabla de amortización.",
        "Suba la tabla de amortización del crédito en la carpeta correspondiente y vuelva a generar.",
    ),
    "amortization_sheet_not_found": (
        "La tabla de amortización existe pero no tiene la hoja o estructura que el sistema espera.",
        "Abra el Excel de amortización y ajuste hojas o encabezados según el formato usado en otros créditos que sí funcionan. "
        "Vuelva a generar.",
    ),
}

_FINALIZE_MESSAGES: dict[str, tuple[str, str]] = {
    "process_not_approved": (
        "La secretaría no marcó el archivo como listo para procesar.",
        "Abra el Excel en 01 REVISION, hoja Control, celda Procesar: ponga SI, guarde, cierre el archivo y vuelva a finalizar.",
    ),
    "missing_control_state": (
        "La hoja Control no tiene el estado del proceso (fila Estado) o el archivo fue alterado.",
        "No borre filas de Control. Si el archivo está dañado, genere uno nuevo con Generate y vuelva a llenar Distribución.",
    ),
    "invalid_control_state": (
        "El estado del proceso en Control no es EN_REVISION (ya fue cerrado o quedó en otro valor).",
        "En Control, verifique que Estado diga EN_REVISION antes de finalizar. Si ya finalizó antes, no repita el paso; use el histórico del día.",
    ),
    "empty_estado_pago": (
        "En Distribución hay filas con datos pero Estado Pago está vacío.",
        "En cada fila con pago, elija un valor de la lista: ADELANTADO, ATRASADO, INCOMPLETO, NORMAL o REVISIÓN MANUAL. "
        "Guarde y vuelva a finalizar.",
    ),
    "invalid_estado_pago": (
        "Hay un Estado Pago escrito a mano o un valor que no está en la lista permitida.",
        "Use solo la lista desplegable: ADELANTADO, ATRASADO, INCOMPLETO, NORMAL, REVISIÓN MANUAL. "
        "No escriba texto libre. Guarde y vuelva a finalizar.",
    ),
    "estado_pago_no_finalizable": (
        "Quedan filas en REVISIÓN MANUAL sin resolver; no se puede cerrar el día.",
        "Revise esas filas en Distribución: cambie el estado cuando el caso esté resuelto o ajuste Validar Pago. "
        "Guarde y vuelva a finalizar.",
    ),
    "no_validar_requires_observation": (
        "Marcó Validar Pago = NO en una fila NORMAL o INCOMPLETO pero no puso observación.",
        "En esa fila, escriba en Observación el motivo (por qué no se valida). Guarde y vuelva a finalizar.",
    ),
    "missing_valor_intereses": (
        "Falta Aplicar a extracto en una fila que debe validarse (Validar Pago = SI).",
        "Complete Aplicar a extracto (o 0 si no aplica). Revise que Saldo por asignar vaya quedando en cero. "
        "Guarde y vuelva a finalizar.",
    ),
    "missing_abono_k": (
        "Falta Mora a aplicar en una fila con Validar Pago = SI.",
        "Complete Mora a aplicar (o 0). Guarde y vuelva a finalizar.",
    ),
    "missing_mora": (
        "Faltan Otros valores en una fila con Validar Pago = SI.",
        "Complete Otros valores (o 0). Guarde y vuelva a finalizar.",
    ),
    "amount_mismatch": (
        "Los importes que la secretaría repartió no suman el monto del banco para ese ID de pago.",
        "En Distribución, para ese pago revise Aplicar a extracto, Mora a aplicar y Otros valores "
        "hasta que Saldo por asignar sea 0. Guarde y vuelva a finalizar.",
    ),
    "missing_extract_route": (
        "Una fila validada no tiene ruta o enlace al extracto, o el PDF ya no está en SharePoint.",
        "Vuelva a ejecutar Generate para regenerar enlaces, o corrija Link extracto y la carpeta del crédito. "
        "Guarde y vuelva a finalizar.",
    ),
    "missing_ruta_unidad_credito": (
        "Falta la columna o dato Ruta unidad de crédito necesario para ubicar carpetas.",
        "Ejecute de nuevo Generate con el banco actual (regenera columnas técnicas). No edite a mano columnas bloqueadas. "
        "Vuelva a finalizar.",
    ),
    "credit_number_not_resolved": (
        "No se pudo identificar el número de crédito para crear la carpeta de asientos.",
        "Revise las columnas Crédito y Ruta unidad de crédito en Distribución. Corrija nombres de carpeta en SharePoint si están mal.",
    ),
    "asientos_folder_create_failed": (
        "SharePoint no dejó crear la carpeta de asientos contables del crédito.",
        "Verifique permisos de escritura y que la ruta del crédito sea correcta. Si el Excel estaba abierto, ciérrelo y reintente.",
    ),
    "missing_ruta_asientos_contables": (
        "Al armar el histórico falta la ruta de la carpeta ASIENTOS CONTABLES.",
        "Ejecute Generate de nuevo y luego Finalize del mismo día, sin saltarse Generate.",
    ),
    "no_validation_file_found": (
        "No hay ningún Excel validacion_pagos_... en 01 REVISION para finalizar.",
        "Ejecute primero Generate del día. Cuando exista el archivo y la secretaría lo complete, ejecute Finalize.",
    ),
    "upload_failed": (
        "El proceso validó bien pero no pudo guardar el histórico o el soporte en SharePoint (red, permisos o archivo bloqueado).",
        "Cierre Excel en escritorio y en el navegador. Verifique espacio y permisos. Reintente Finalize; "
        "si persiste, contacte soporte con la hora del error.",
    ),
    "missing_control_sheet": (
        "El archivo de revisión no tiene la hoja Control.",
        "No use un Excel manual distinto. Ejecute Generate y trabaje solo sobre el archivo que genera el sistema.",
    ),
    "missing_distribucion_sheet": (
        "El archivo de revisión no tiene la hoja Distribución.",
        "Ejecute Generate y use solo el archivo que genera el sistema.",
    ),
    "missing_sheet_headers": (
        "Las tablas de Distribución u otra hoja no tienen los encabezados esperados; el archivo fue modificado de más.",
        "Vuelva a ejecutar Generate. No borre filas de encabezado ni renombre columnas.",
    ),
    "credit_folder_not_found": (
        "Al cerrar el día, no se encontró la carpeta del crédito para un extracto validado.",
        "Verifique en SharePoint que la carpeta del crédito exista y coincida con el nombre en Distribución. "
        "Corrija y vuelva a finalizar.",
    ),
    "credit_folder_ambiguous": (
        "Hay varias carpetas posibles para el mismo crédito; el sistema no puede elegir una.",
        "Deje una sola carpeta por crédito (nombres únicos en SharePoint). Vuelva a finalizar.",
    ),
    "reprogramar_requires_zero_total": (
        "Un pago ADELANTADO con Validar Pago = SI debe llevar cero en los tres importes de aplicación.",
        "Deje en 0 Aplicar a extracto, Mora a aplicar y Otros valores, o cambie Estado Pago o Validar Pago si el caso es otro. "
        "Guarde y vuelva a finalizar.",
    ),
    "validar_requires_positive_total": (
        "Marcó Validar Pago = SI pero el total aplicado es cero o negativo (NORMAL, INCOMPLETO o ATRASADO).",
        "Ingrese los montos a aplicar o cambie Validar Pago a NO con observación. Guarde y vuelva a finalizar.",
    ),
}


def _strip_exception_prefix(message: str) -> str:
    s = (message or "").strip()
    m = re.match(r"^(?:ValueError|RuntimeError|GraphConfigError|HTTPStatusError|Exception)\s*:\s*(.*)$", s, re.I | re.S)
    if m:
        return m.group(1).strip()
    return s


def _pick_table_code(table: dict[str, tuple[str, str]], raw: str) -> str | None:
    if raw in table:
        return raw
    for k in sorted(table.keys(), key=len, reverse=True):
        if k in raw:
            return k
    return None


def _error_code_from_generate_or_finalize_message(job_type: str, message: str) -> str:
    raw = _strip_exception_prefix(message)
    if "|" in raw and raw.lower().startswith("upload_failed"):
        return "upload_failed"
    if job_type == "generate":
        hit = _pick_table_code(_GENERATE_MESSAGES, raw)
        if hit:
            return hit
    if job_type == "finalize":
        for code in _FINALIZE_MESSAGES:
            if raw == code or raw.startswith(code + "|") or raw.startswith(code + " "):
                return code
        hit = _pick_table_code(_FINALIZE_MESSAGES, raw)
        if hit:
            return hit
    return raw.split("|", 1)[0].strip()[:120] or "unknown_error"


def _lookup_generate_finalize(job_type: str, code: str, full_message: str) -> tuple[str, str, str]:
    if job_type == "generate":
        table = _GENERATE_MESSAGES
    elif job_type == "finalize":
        table = _FINALIZE_MESSAGES
    else:
        return _UNKNOWN_USER, _UNKNOWN_NEXT, "unknown_error"

    if code in table:
        u, n = table[code]
        return u, n, code
    hit = _pick_table_code(table, _strip_exception_prefix(full_message))
    if hit and hit in table:
        u, n = table[hit]
        return u, n, hit
    return _UNKNOWN_USER, _UNKNOWN_NEXT, code


def _notify_merge_string_mapping(job_type: str, msg: str) -> tuple[str, str, str]:
    mlow = msg.lower()

    if job_type == "notify_validar_extractos":
        mstripped = msg.strip()
        if mstripped == "missing_historical_file_path" or mstripped.startswith(
            "missing_historical_file_path|"
        ):
            return (
                "No se recibió el archivo histórico de validación.",
                "Ejecute primero la finalización de pagos y asegúrese de pasar el historical_file_path generado por Finalize al proceso de envío de correo.",
                "missing_historical_file_path",
            )
        if mstripped.startswith("historical_file_not_found"):
            return (
                "No se encontró el archivo histórico de validación.",
                "Verifique que la finalización haya terminado correctamente y que el archivo histórico siga existiendo en SharePoint.",
                "historical_file_not_found",
            )
        if "no hay excel" in mlow and "hist" in mlow:
            return (
                "No se encontró el archivo histórico de validación.",
                "Verifique que la finalización haya terminado correctamente y que el archivo histórico siga existiendo en SharePoint.",
                "historical_file_not_found",
            )
        if mstripped == "missing_distribucion_headers" or mstripped.startswith(
            "missing_distribucion_headers|"
        ):
            return (
                "No se encontraron los encabezados requeridos en la hoja Distribución del histórico.",
                "Verifique que el archivo histórico provenga de una finalización correcta.",
                "missing_distribucion_headers",
            )
        if mstripped == "missing_distribucion_status_column" or mstripped.startswith(
            "missing_distribucion_status_column|"
        ):
            return (
                "No se encontró la columna Estado Pago / Validar Pago (o columna Estado en históricos antiguos) en el histórico.",
                "Vuelva a ejecutar la finalización con el Excel de revisión actualizado o contacte soporte.",
                "missing_distribucion_status_column",
            )
        if mstripped == "missing_distribucion_route_column" or mstripped.startswith(
            "missing_distribucion_route_column|"
        ):
            return (
                "No se encontró la columna técnica Ruta en el histórico de validación.",
                "Vuelva a ejecutar Finalize con el Excel de revisión actualizado. "
                "Si el problema persiste, contacte soporte.",
                "missing_distribucion_route_column",
            )
        if "no se encontraron encabezados" in mlow and "distribución" in mlow:
            return (
                "No se encontraron los encabezados requeridos en la hoja Distribución del histórico.",
                "Verifique que el archivo histórico provenga de una finalización correcta.",
                "missing_distribucion_headers",
            )
        if "faltan columnas" in mlow or (
            "estado" in mlow and "línea" in mlow and "ruta" in mlow and "requeridas" in mlow
        ):
            return (
                "El archivo histórico no tiene la columna Ruta o la estructura esperada.",
                "Ejecute nuevamente Finalize o revise el archivo histórico.",
                "missing_ruta_column",
            )
        if "no hay filas" in mlow and ("estado" in mlow or "línea" in mlow or "linea" in mlow):
            return (
                "No hay filas con Validar Pago=SI (o filas VALIDAR en históricos antiguos) para enviar.",
                "Revise la hoja Distribución del archivo histórico.",
                "no_validated_rows",
            )
        if "descarg" in mlow or ("no se pudo" in mlow and "pdf" in mlow):
            return (
                "No se encontró uno de los extractos a adjuntar.",
                "Verifique que el PDF exista en la ruta indicada en la columna Ruta.",
                "extract_pdf_not_found",
            )
        if "emisor" in mlow or "receptores" in mlow or "correos" in mlow or "remitente" in mlow:
            return (
                "No hay destinatarios o remitente configurados para el correo.",
                "Revise CORREOS.xlsx o la configuración de destinatarios.",
                "recipients_not_configured",
            )
        if "sendmail" in mlow or ("graph" in mlow and ("401" in msg or "403" in msg or "error" in mlow)):
            return (
                "No se pudo enviar el correo por Microsoft Graph.",
                "Revise permisos, destinatarios, buzón remitente o el detalle técnico.",
                "graph_sendmail_failed",
            )

    if job_type == "merge_composite_validado_pdfs":
        mstripped = msg.strip()
        if mstripped == "merge_control_no_pending_process":
            return (
                "No hay un proceso pendiente de consolidación de PDFs.",
                "Ejecute primero la validación y el envío de correo de extractos antes de consolidar.",
                "merge_control_no_pending_process",
            )
        if mstripped == "missing_historical_file_path":
            return (
                "No se encontró la ruta del histórico de validación.",
                "Vuelva a ejecutar el proceso de envío de correo o revise el archivo de control.",
                "missing_historical_file_path",
            )
        if mstripped == "missing_email_pdf_path":
            return (
                "No se encontró el PDF del correo de extractos.",
                "Vuelva a ejecutar el envío de correo de extractos antes de consolidar.",
                "missing_email_pdf_path",
            )
        if mstripped == "merge_control_workbook_not_found":
            return (
                "No se encontró el archivo de control de consolidación en SharePoint.",
                "Ejecute primero el setup del workbook de control o verifique GRAPH_MERGE_CONTROL_WORKBOOK_PATH.",
                "merge_control_workbook_not_found",
            )
        if mstripped == "merge_control_invalid_structure":
            return (
                "El archivo de control de consolidación no tiene el formato esperado.",
                "Vuelva a ejecutar el setup del workbook de control o restaure la plantilla.",
                "merge_control_invalid_structure",
            )
        if "no se encontró pdf de correo" in mlow:
            return (
                "Falta el PDF exportado del correo para la fecha del reporte.",
                "Ejecute primero el endpoint de correo o verifique la carpeta 05 EMAIL.",
                "email_pdf_not_found",
            )
        if mstripped == "missing_distribucion_headers" or mstripped.startswith(
            "missing_distribucion_headers|"
        ):
            return (
                "No se encontraron los encabezados requeridos en la hoja Distribución del histórico.",
                "Verifique que el archivo histórico provenga de una finalización correcta.",
                "missing_distribucion_headers",
            )
        if mstripped == "missing_distribucion_status_column" or mstripped.startswith(
            "missing_distribucion_status_column|"
        ):
            return (
                "No se encontró la columna Estado Pago / Validar Pago (o columna Estado en históricos antiguos) en el histórico.",
                "Vuelva a ejecutar la finalización con el Excel de revisión actualizado o contacte soporte.",
                "missing_distribucion_status_column",
            )
        if mstripped == "missing_distribucion_route_column" or mstripped.startswith(
            "missing_distribucion_route_column|"
        ):
            return (
                "No se encontró la columna técnica Ruta en el histórico de validación.",
                "Vuelva a ejecutar Finalize con el Excel de revisión actualizado. "
                "Si el problema persiste, contacte soporte.",
                "missing_distribucion_route_column",
            )
        if "en distribución se requieren columnas" in mlow:
            return (
                "El histórico no tiene las columnas necesarias para consolidar PDFs.",
                "Verifique que el histórico generado por Finalize tenga Estado Pago, Validar Pago, Ruta e ID Pago.",
                "missing_distribucion_columns",
            )
        if "no hay filas cuyo estado" in mlow or "estado línea contenga" in mlow or "estado linea contenga" in mlow:
            return (
                "No hay pagos marcados para consolidar.",
                "Revise que el histórico tenga filas con Validar Pago=SI (o VALIDAR en históricos antiguos) en Distribución.",
                "no_validated_rows_merge",
            )
        if "extract_routes_missing" in msg or "sin rutas de extracto" in mlow:
            return (
                "Un pago no tiene rutas de extractos para consolidar.",
                "Revise la columna Ruta en el histórico.",
                "extract_routes_missing",
            )
        if mstripped == "missing_ruta_asientos_contables" or "missing_ruta_asientos_contables" in msg:
            return (
                "Falta la ruta de la carpeta de asientos contables en el histórico.",
                "Ejecute Finalize para generar la columna RutaAsientosContables antes de consolidar.",
                "missing_ruta_asientos_contables",
            )
        if "credit_number_not_resolved" in msg:
            return (
                "No se pudo determinar el número de crédito para un extracto.",
                "Revise que la ruta del extracto incluya la carpeta CREDITO # n o que el histórico tenga una sola fila de crédito por pago.",
                "credit_number_not_resolved",
            )
        if (
            "asiento_contable_not_found" in msg
            or "asiento_contable_ambiguous" in msg
            or "asiento_contable_credit_mismatch" in msg
        ):
            return (
                "Falta el PDF de asiento contable para uno o más pagos o no coincide con el crédito.",
                "Suba un único PDF en ASIENTOS CONTABLES del crédito cuyo nombre contenga el número de crédito aislado.",
                "missing_asiento_contable_pdf",
            )
        if "extracto_download_failed" in msg or "no se descargó extracto" in mlow:
            return (
                "No se pudo descargar uno de los extractos.",
                "Verifique que la ruta del extracto exista y que la API tenga permisos.",
                "extract_pdf_download_failed",
            )
        if "asiento_download_failed" in msg or "no se descargó asiento" in mlow:
            return (
                "No se pudo descargar el PDF de asiento contable.",
                "Verifique que el archivo exista en la carpeta ASIENTOS CONTABLES del crédito.",
                "asiento_pdf_download_failed",
            )
        if "consolidated_upload_failed" in msg or "put_bytes" in mlow or "upload" in mlow or "423" in msg or "locked" in mlow:
            return (
                "No se pudo subir el PDF consolidado.",
                "Revise permisos de SharePoint, carpeta de salida o bloqueo de archivos.",
                "consolidated_upload_failed",
            )

    return _UNKNOWN_USER, _UNKNOWN_NEXT, "unknown_error"


def _build_standard_error_payload(
    job_type: str,
    *,
    exc_type: str | None,
    message: str,
) -> dict[str, Any]:
    msg = _strip_exception_prefix(message)
    if job_type in ("generate", "finalize"):
        code = _error_code_from_generate_or_finalize_message(job_type, msg)
        user, next_a, _ = _lookup_generate_finalize(job_type, code, msg)
    else:
        user, next_a, code = _notify_merge_string_mapping(job_type, msg)

    return {
        "type": exc_type or "Error",
        "message": message,
        "error_code": code,
        "technical_message": message,
        "user_message": user,
        "next_action": next_a,
        "severity": "error",
    }


def _merge_completed_enrichment(job_type: str, result: dict[str, Any]) -> tuple[str, str, str]:
    if job_type == "generate":
        return (
            "Se creó el archivo de revisión de pagos del día. Ya puede abrirlo en la carpeta 01 REVISION de SharePoint.",
            "Abra ese Excel, complete la hoja Distribución (Estado Pago y Validar Pago en cada fila) y en la hoja Control "
            "ponga Procesar = SI cuando termine. Luego ejecute de nuevo el proceso de finalización.",
            "success",
        )
    if job_type == "finalize":
        return (
            "La revisión quedó cerrada correctamente. Se guardó el histórico del día y el archivo de soporte "
            "para que la secretaría suba los asientos contables.",
            "Abra el archivo de soporte de asientos (enlaces por crédito), cargue los PDF en cada carpeta ASIENTOS CONTABLES "
            "y continúe con el envío del correo de extractos cuando corresponda.",
            "success",
        )
    if job_type == "notify_validar_extractos":
        ec = str(result.get("merge_control_error_code") or "").strip()
        wtxt = str(result.get("merge_control_warning") or "").strip()
        if ec == "merge_control_active_process_exists":
            return (
                "El correo con los extractos validados fue enviado correctamente.",
                "Ya existe un proceso pendiente de consolidación de PDFs. Termine o cancele ese proceso "
                "antes de registrar uno nuevo en el archivo de control.",
                "warning",
            )
        if ec == "missing_email_pdf_path_for_merge_control":
            return (
                "El correo con los extractos validados fue enviado correctamente.",
                "No se generó el PDF del correo necesario para registrar el proceso de consolidación en el "
                "archivo de control. Revise la exportación del PDF antes de ejecutar Merge.",
                "warning",
            )
        if ec in (
            "merge_control_workbook_not_found",
            "merge_control_invalid_structure",
            "merge_control_read_failed",
            "merge_control_write_failed",
        ):
            return (
                "El correo con los extractos validados fue enviado correctamente.",
                wtxt or "No se pudo actualizar el archivo de control de Merge. Revise la ruta, permisos o ejecute el "
                "setup del workbook.",
                "warning",
            )
        if not result.get("merge_control_updated") and wtxt and not ec:
            return (
                "El correo con los extractos validados fue enviado correctamente.",
                wtxt,
                "warning",
            )
        return (
            "El correo con los extractos validados fue enviado correctamente.",
            "Verifique que los destinatarios reciban el mensaje, cargue los asientos contables en las carpetas "
            "indicadas y luego ejecute la consolidación de PDFs.",
            "success",
        )
    if job_type == "merge_composite_validado_pdfs":
        outputs = result.get("outputs") or []
        skipped = result.get("skipped") or []
        if isinstance(outputs, list) and len(outputs) == 0 and isinstance(skipped, list) and len(skipped) > 0:
            return (
                "El proceso terminó, pero no se generó ningún PDF consolidado porque faltan archivos requeridos.",
                "Revise la lista de pagos omitidos, cargue los asientos contables o corrija las rutas faltantes y vuelva a ejecutar la consolidación.",
                "warning",
            )
        if isinstance(outputs, list) and len(outputs) > 0:
            if isinstance(skipped, list) and len(skipped) > 0:
                return (
                    "Se generaron los PDFs de soporte consolidados para los pagos que tenían todos los archivos necesarios.",
                    "Revise la carpeta de salida de PDFs consolidados y verifique la lista de pagos omitidos si existe.",
                    "warning",
                )
            return (
                "Se generaron los PDFs de soporte consolidados para los pagos que tenían todos los archivos necesarios.",
                "Revise la carpeta de salida de PDFs consolidados y verifique la lista de pagos omitidos si existe.",
                "success",
            )
        return (
            "Se generaron los PDFs de soporte consolidados para los pagos que tenían todos los archivos necesarios.",
            "Revise la carpeta de salida de PDFs consolidados y verifique la lista de pagos omitidos si existe.",
            "success",
        )

    return "", "", ""


def enrich_job_for_http_response(job: dict[str, Any]) -> dict[str, Any]:
    """
    Devuelve una copia superficial del job con user_message, next_action, severity
    y error normalizado cuando aplica.
    """
    out = dict(job)
    jt = str(out.get("type") or "")
    if jt not in ENRICHABLE_JOB_TYPES:
        return out

    status = out.get("status")
    if status == "completed":
        result = out.get("result")
        if not isinstance(result, dict):
            result = {}
        um, na, sev = _merge_completed_enrichment(jt, result)
        if um:
            out["user_message"] = um
        if na:
            out["next_action"] = na
        if sev:
            out["severity"] = sev
        return out

    if status == "failed":
        out["severity"] = "error"
        err = out.get("error")
        if isinstance(err, dict):
            msg = str(err.get("message", ""))
            exc_type = str(err.get("type", "Error"))
            if err.get("user_message") and err.get("next_action") and err.get("error_code"):
                merged = dict(err)
                merged.setdefault("technical_message", msg)
                merged.setdefault("severity", "error")
                out["error"] = merged
                return out
            payload = _build_standard_error_payload(jt, exc_type=exc_type, message=msg)
            out["error"] = {**err, **payload}
            out["error"]["message"] = err.get("message", payload["message"])
            out["error"]["type"] = err.get("type", payload["type"])
        elif isinstance(err, str):
            out["error"] = _build_standard_error_payload(jt, exc_type=None, message=err)
        else:
            out["error"] = _build_standard_error_payload(
                jt, exc_type="Error", message=str(err) if err is not None else ""
            )
        return out

    return out


def examples_for_parse_json_tests() -> tuple[dict[str, Any], dict[str, Any]]:
    """Ejemplos estables para tests de esquema flexible."""
    completed = {
        "job_id": "ex-1",
        "type": "generate",
        "status": "completed",
        "user_message": "ok",
        "next_action": "next",
        "severity": "success",
        "result": {"validation_file": "f.xlsx"},
        "queued_at": "2026-01-01T00:00:00+00:00",
    }
    failed = {
        "job_id": "ex-2",
        "type": "finalize",
        "status": "failed",
        "severity": "error",
        "error": {
            "type": "ValueError",
            "message": "amount_mismatch",
            "error_code": "amount_mismatch",
            "technical_message": "amount_mismatch",
            "user_message": "x",
            "next_action": "y",
            "severity": "error",
        },
    }
    return completed, failed
