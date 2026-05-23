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
        "amortization_dry_run",
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
                "El envío de correo no arrancó porque Power Automate no indicó qué archivo histórico usar "
                "(falta la ruta del Excel de cartera validada del día).",
                "Ejecute primero Finalize del mismo día. En el flujo de correo, envíe exactamente la ruta "
                "que devolvió Finalize en historical_file_path (carpeta 02 HISTORICO), no la URL de SharePoint.",
                "missing_historical_file_path",
            )
        if mstripped.startswith("historical_file_not_found"):
            return (
                "No se pudo abrir el archivo histórico de validación en SharePoint (no existe, fue movido o sin permiso).",
                "Confirme que Finalize terminó bien y que el archivo cartera_validada_....xlsx sigue en 02 HISTORICO. "
                "Si lo borraron o renombraron, vuelva a ejecutar Finalize y reintente el correo con la ruta nueva.",
                "historical_file_not_found",
            )
        if "no hay excel" in mlow and "hist" in mlow:
            return (
                "No se encontró el Excel histórico del día en la carpeta de histórico de validación de pagos.",
                "Verifique en 02 HISTORICO que exista cartera_validada con la fecha del reporte. "
                "Si falta, ejecute Finalize de nuevo antes del correo.",
                "historical_file_not_found",
            )
        if mstripped == "missing_distribucion_headers" or mstripped.startswith(
            "missing_distribucion_headers|"
        ) or (
            "no se encontró una hoja llamada" in mlow and "distribución" in mlow
        ):
            return (
                "El archivo histórico no tiene la hoja Distribución que el correo necesita leer.",
                "Use el histórico generado por Finalize del mismo flujo (no un Excel copiado a mano). "
                "Si el archivo es antiguo, vuelva a ejecutar Finalize y reintente el correo.",
                "missing_distribucion_headers",
            )
        if mstripped == "missing_distribucion_status_column" or mstripped.startswith(
            "missing_distribucion_status_column|"
        ):
            return (
                "En el histórico falta la columna para saber qué pagos van en el correo "
                "(Estado Pago / Validar Pago, o Estado en archivos viejos).",
                "Vuelva a ejecutar Finalize con el Excel de revisión actual del sistema. "
                "No edite manualmente los encabezados de Distribución.",
                "missing_distribucion_status_column",
            )
        if mstripped == "missing_distribucion_route_column" or mstripped.startswith(
            "missing_distribucion_route_column|"
        ):
            return (
                "En el histórico falta la columna Ruta, necesaria para adjuntar los PDF de extractos.",
                "Ejecute de nuevo Generate y Finalize del día para regenerar la columna Ruta. "
                "Luego reintente el envío de correo.",
                "missing_distribucion_route_column",
            )
        if "no se encontraron encabezados" in mlow and "distribución" in mlow:
            return (
                "La hoja Distribución del histórico no tiene la fila de encabezados que el sistema espera.",
                "No modifique la primera fila de títulos del histórico. Regenere el archivo con Finalize.",
                "missing_distribucion_headers",
            )
        if "faltan columnas" in mlow or (
            "estado" in mlow and "línea" in mlow and "ruta" in mlow and "requeridas" in mlow
        ):
            return (
                "El histórico no tiene todas las columnas necesarias (estado y ruta de extractos).",
                "Ejecute Finalize otra vez con la plantilla actual. Revise que Distribución conserve Ruta y Estado Pago.",
                "missing_ruta_column",
            )
        if "no hay filas" in mlow and ("estado" in mlow or "línea" in mlow or "linea" in mlow):
            return (
                "En el histórico no hay filas marcadas para enviar en el correo "
                "(Validar Pago = SI o estado VALIDAR según configuración).",
                "Abra el histórico en Distribución y confirme que haya pagos validados para el día. "
                "Si la secretaría no marcó filas, corrija el Excel de revisión y vuelva a Finalize.",
                "no_validated_rows",
            )
        if "no hay destinatarios" in mlow or (
            "receptores" in mlow and "vacía" in mlow
        ):
            return (
                "El correo no se envió porque no hay destinatarios válidos (columna RECEPTORES vacía o correos mal escritos).",
                "Abra 00 CONTROL / CORREOS.xlsx: columna RECEPTORES debe tener al menos un correo por fila. "
                "Si usó destinatarios fijos en Power Automate, revise el campo to del body.",
                "recipients_not_configured",
            )
        if "emisor" in mlow or "receptores" in mlow or "correos" in mlow or "remitente" in mlow:
            return (
                "Falta configurar quién envía o quién recibe el correo en CORREOS.xlsx (EMISOR y RECEPTORES).",
                "En 00 CONTROL / CORREOS.xlsx complete EMISOR (un correo) y RECEPTORES (uno o más correos). "
                "Guarde el archivo en SharePoint y reintente.",
                "recipients_not_configured",
            )
        if "no hay columna fecha" in mlow and "banco" in mlow:
            return (
                "El reporte del banco BANCO_BOGOTA.xlsx no tiene columna Fecha; el correo no puede saber el día del abono.",
                "Revise el Excel del banco en 00 COMWARE - CARGA TRANSACCIONES BANCO y agregue la columna Fecha "
                "como en días anteriores. Suba el archivo y reintente.",
                "bank_report_missing_date_column",
            )
        if "ninguna fecha válida" in mlow and "fecha" in mlow:
            return (
                "El reporte del banco tiene columna Fecha pero ninguna fecha se pudo leer (celdas vacías o formato distinto).",
                "Revise que las filas de pagos tengan fecha en formato habitual (dd/mm/aaaa o similar). "
                "Corrija BANCO_BOGOTA.xlsx y reintente el correo.",
                "bank_report_no_valid_dates",
            )
        if "banco bogotá" in mlow or "banco bogota" in mlow:
            if "filas" in mlow or "columnas" in mlow:
                return (
                    "El reporte del banco no tiene datos completos para armar la tabla del correo "
                    "(filas vacías, columnas faltantes o celdas incompletas).",
                    "Abra BANCO_BOGOTA.xlsx: cada fila del correo debe tener todos los campos llenos según la plantilla. "
                    "Suba el archivo corregido a SharePoint y reintente.",
                    "bank_report_invalid_data",
                )
        if "define graph" in mlow or "graphconfigerror" in mlow.replace(" ", ""):
            return (
                "Falta configuración del sistema para ubicar carpetas de histórico o correo.",
                "Contacte a soporte técnico (no es un error de la secretaría). Indique que falló el envío de correo de extractos.",
                "notify_config_missing",
            )
        if "descarg" in mlow or ("no se pudo" in mlow and "pdf" in mlow):
            return (
                "Uno de los PDF de extractos que debían adjuntarse al correo no se encontró o no se pudo descargar.",
                "En el histórico, columna Ruta: abra cada enlace y confirme que el PDF existe en SharePoint. "
                "Suba el extracto faltante y reintente.",
                "extract_pdf_not_found",
            )
        if "sendmail" in mlow or ("graph" in mlow and ("401" in msg or "403" in msg or "error" in mlow)):
            return (
                "Microsoft no aceptó el envío del correo (permisos del buzón, remitente o destinatarios).",
                "Verifique en CORREOS.xlsx que EMISOR sea un buzón autorizado para enviar. "
                "Confirme que los destinatarios son correos válidos. Si persiste, contacte soporte con la hora del fallo.",
                "graph_sendmail_failed",
            )

    if job_type == "merge_composite_validado_pdfs":
        mstripped = msg.strip()
        if mstripped == "merge_control_no_pending_process":
            return (
                "No se puede unir PDFs porque el sistema no tiene un proceso activo registrado tras el correo del día.",
                "Ejecute en este orden: Finalize → envío de correo de extractos (debe quedar registrado en "
                "control_merge_pdfs.xlsx con estado pendiente de asientos). Luego ejecute Unir PDFs. "
                "Si el correo de hoy no actualizó el control, reenvíe el correo o pida a soporte.",
                "merge_control_no_pending_process",
            )
        if mstripped == "missing_historical_file_path":
            return (
                "El archivo de control no tiene la ruta del histórico de validación necesaria para unir PDFs.",
                "Vuelva a ejecutar el envío de correo del día (debe registrar historical_file_path en control_merge_pdfs). "
                "Si el control está vacío o dañado, pida a soporte restaurar control_merge_pdfs.xlsx en 00 CONTROL.",
                "missing_historical_file_path",
            )
        if mstripped == "missing_email_pdf_path":
            return (
                "Falta en el control la ruta del PDF copia del correo (carpeta 05 EMAIL).",
                "Ejecute de nuevo el envío de correo y confirme que en 05 EMAIL quede el PDF del día "
                "(ABONOS BANCO BOGOTA ...). Después reintente Unir PDFs.",
                "missing_email_pdf_path",
            )
        if mstripped == "merge_control_workbook_not_found":
            return (
                "No existe control_merge_pdfs.xlsx en 00 CONTROL; sin ese archivo no puede iniciar la unión de PDFs.",
                "Pida a soporte crear el archivo de control una sola vez (setup). "
                "Luego: correo del día → Unir PDFs.",
                "merge_control_workbook_not_found",
            )
        if mstripped == "merge_control_invalid_structure":
            return (
                "control_merge_pdfs.xlsx está dañado o fue editado (falta hoja Procesos, encabezados o fila 2).",
                "No modifique ese Excel a mano. Pida a soporte restaurar la plantilla en 00 CONTROL y repita correo + Unir PDFs.",
                "merge_control_invalid_structure",
            )
        if "no se encontró pdf de correo" in mlow:
            return (
                "No se encontró en 05 EMAIL el PDF del correo para la fecha del reporte del banco.",
                "Ejecute primero el envío de correo del día y verifique que el PDF se guarde en "
                "02 COMWARE - VALIDACION PAGOS / 05 EMAIL. Luego reintente Unir PDFs.",
                "email_pdf_not_found",
            )
        if mstripped == "missing_distribucion_headers" or mstripped.startswith(
            "missing_distribucion_headers|"
        ) or (
            "no se encontró una hoja llamada" in mlow and "distribución" in mlow
        ):
            return (
                "El histórico del día no tiene la hoja Distribución que se necesita para saber qué pagos unir.",
                "Use el cartera_validada generado por Finalize del mismo flujo. Si el archivo es copia manual, "
                "vuelva a ejecutar Finalize y reintente.",
                "missing_distribucion_headers",
            )
        if mstripped == "missing_distribucion_status_column" or mstripped.startswith(
            "missing_distribucion_status_column|"
        ) or (
            "estado pago" in mlow and "requieren columnas" in mlow
        ):
            return (
                "El histórico no indica qué pagos deben consolidarse (falta Estado Pago / Validar Pago o Estado línea).",
                "Ejecute de nuevo Generate y Finalize del día. No altere encabezados de Distribución en el histórico.",
                "missing_distribucion_status_column",
            )
        if mstripped == "missing_distribucion_route_column" or mstripped.startswith(
            "missing_distribucion_route_column|"
        ) or (
            "rutaasientoscontables" in mlow.replace(" ", "")
            and "requieren" in mlow
        ):
            return (
                "Al histórico le faltan columnas obligatorias: Ruta (extractos), RutaAsientosContables e ID Pago.",
                "Regenere el histórico con Finalize actual (incluye rutas técnicas). Luego correo y Unir PDFs.",
                "missing_distribucion_route_column",
            )
        if "en distribución se requieren columnas" in mlow:
            return (
                "El histórico no tiene todas las columnas para armar los PDF unidos (estado, rutas, ID Pago).",
                "Vuelva a Finalize con la plantilla vigente. Revise Distribución: Ruta, RutaAsientosContables, ID Pago.",
                "missing_distribucion_columns",
            )
        if "no hay filas cuyo estado" in mlow or "estado línea contenga" in mlow or "estado linea contenga" in mlow:
            return (
                "No hay pagos en el histórico marcados para consolidar (ninguna fila coincide con VALIDAR / Validar Pago = SI).",
                "En el histórico, hoja Distribución, confirme que haya filas validadas para el día. "
                "Si la secretaría no cerró bien la revisión, corrija el Excel y vuelva a Finalize antes del correo y este paso.",
                "no_validated_rows_merge",
            )
        if "extract_routes_missing" in msg or "sin rutas de extracto" in mlow or "extract_routes_missing" in mlow:
            return (
                "Un pago no tiene ruta de extracto en el histórico o el PDF no está en SharePoint.",
                "En Distribución, columna Ruta: abra el enlace y confirme que el extracto exista. "
                "Si falta, corrija con Generate/Finalize o suba el PDF al crédito.",
                "extract_routes_missing",
            )
        if mstripped == "missing_ruta_asientos_contables" or "missing_ruta_asientos_contables" in msg:
            return (
                "Un pago no tiene carpeta de asientos contables registrada en el histórico (columna RutaAsientosContables).",
                "Ejecute Finalize de nuevo para regenerar esa columna. La secretaría debe haber completado la revisión antes.",
                "missing_ruta_asientos_contables",
            )
        if "credit_number_not_resolved" in msg:
            return (
                "No se pudo identificar el número de crédito de un extracto para buscar su asiento contable.",
                "Revise que la ruta del extracto pase por una carpeta CREDITO # número o que la columna Crédito "
                "en Distribución tenga el número correcto.",
                "credit_number_not_resolved",
            )
        if "asiento_folder_list_failed" in msg:
            return (
                "No se pudo abrir la carpeta de asientos contables de un crédito en SharePoint.",
                "Verifique que la ruta RutaAsientosContables del histórico sea correcta y que tenga permiso de lectura. "
                "Confirme que la carpeta ASIENTOS CONTABLES del crédito exista.",
                "asiento_folder_list_failed",
            )
        if "asiento_contable_not_found" in msg:
            return (
                "Falta al menos un PDF de asiento contable válido en la carpeta del crédito.",
                "En la carpeta ASIENTOS CONTABLES del crédito suba al menos un PDF cuyo nombre incluya el número "
                "de crédito (sin confundir con otros números, ej. 264 vs 1264). Vuelva a ejecutar Unir PDFs.",
                "missing_asiento_contable_pdf",
            )
        if "asiento_contable_credit_mismatch" in msg:
            return (
                "Hay PDF de asiento en la carpeta del crédito que no coincide con el número de crédito esperado.",
                "Renombre o retire los PDF que no correspondan al crédito; los válidos deben incluir el número de "
                "crédito en el nombre. Vuelva a ejecutar Unir PDFs.",
                "asiento_contable_credit_mismatch",
            )
        if "asiento_contable_ambiguous" in msg:
            return (
                "No se pudo determinar qué PDF de asiento usar para el crédito.",
                "Revise los archivos en la carpeta ASIENTOS CONTABLES y vuelva a ejecutar Unir PDFs.",
                "asiento_contable_ambiguous",
            )
        if "extracto_download_failed" in msg or "no se descargó extracto" in mlow:
            return (
                "No se pudo descargar el PDF de un extracto indicado en el histórico.",
                "Compruebe que el archivo siga en SharePoint en la ruta de la columna Ruta y que no esté borrado o renombrado.",
                "extract_pdf_download_failed",
            )
        if "asiento_download_failed" in msg or "no se descargó asiento" in mlow:
            return (
                "No se pudo descargar el PDF de asiento contable aunque la carpeta existe.",
                "Verifique que el archivo no esté corrupto, bloqueado o sin permisos. Vuelva a subir el asiento en "
                "ASIENTOS CONTABLES del crédito.",
                "asiento_pdf_download_failed",
            )
        if "consolidated_upload_failed" in msg or "put_bytes" in mlow or ("upload" in mlow and "423" in msg) or "locked" in mlow:
            return (
                "Se armó el PDF unido pero no se pudo guardarlo en 06 ASIENTO CONTABLES GENERADOS (permisos o archivo bloqueado).",
                "Cierre PDFs abiertos en SharePoint. Verifique espacio y permisos de escritura en la carpeta de salida. "
                "Reintente Unir PDFs.",
                "consolidated_upload_failed",
            )
        if "upload" in mlow and "consolidated" not in mlow:
            return (
                "No se pudo subir uno de los PDF consolidados a SharePoint.",
                "Revise permisos en 06 ASIENTO CONTABLES GENERADOS y que ningún PDF consolidado esté abierto en el navegador.",
                "consolidated_upload_failed",
            )

    if job_type == "amortization_dry_run":
        mstripped = msg.strip()
        if mstripped == "report_date_iso_required" or mstripped.startswith("report_date_iso_required"):
            return (
                "No se indicó la fecha del reporte ni una ruta de manifest de Merge.",
                "Vuelva a ejecutar el dry-run enviando report_date_iso (YYYY-MM-DD) o merge_manifest_path.",
                "report_date_iso_required",
            )
        if mstripped.startswith("merge_manifest_not_found"):
            return (
                "No se encontró el archivo merge_manifest en SharePoint para la fecha indicada.",
                "Confirme que Merge se ejecutó para ese día y que el JSON está en la carpeta de logs "
                "(GRAPH_PAYMENT_VALIDATION_LOGS_PATH). Corrija report_date_iso o pase merge_manifest_path explícito.",
                "merge_manifest_not_found",
            )
        if mstripped.startswith("invalid_merge_manifest_json"):
            return (
                "El manifest de Merge existe pero no es un JSON válido.",
                "Revise el archivo en SharePoint o vuelva a ejecutar Merge para regenerar el manifest.",
                "invalid_merge_manifest_json",
            )
        if "graph_config" in mlow or "missing environment variable" in mlow:
            return (
                "Faltan variables de configuración de SharePoint para leer el manifest o los archivos.",
                "Contacte a soporte técnico con el detalle del error (GRAPH_SHAREPOINT_SITE_SEARCH, rutas de logs, etc.).",
                "graph_config_error",
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
                "Hizo bien el envío del correo: los destinatarios deberían haberlo recibido.",
                "Para el siguiente paso (unir PDFs): en 00 CONTROL / control_merge_pdfs.xlsx ya hay un proceso "
                "pendiente. Termine o cancele ese proceso antes de volver a registrar uno nuevo.",
                "warning",
            )
        if ec == "missing_email_pdf_path_for_merge_control":
            return (
                "Hizo bien el envío del correo: los destinatarios deberían haberlo recibido.",
                "No se guardó el PDF copia del correo en 05 EMAIL ni se registró el paso para unir PDFs después. "
                "Revise en SharePoint la carpeta 05 EMAIL y que la exportación del PDF esté activa; luego reintente "
                "solo el registro en control o contacte soporte antes de ejecutar Unir PDFs.",
                "warning",
            )
        if ec == "merge_control_workbook_not_found":
            return (
                "Hizo bien el envío del correo: los destinatarios deberían haberlo recibido.",
                "No se actualizó control_merge_pdfs.xlsx porque el archivo no existe en 00 CONTROL. "
                "Pida a soporte ejecutar una sola vez el setup del archivo de control; después el flujo de correo "
                "quedará listo para el paso de unir PDFs.",
                "warning",
            )
        if ec == "merge_control_invalid_structure":
            return (
                "Hizo bien el envío del correo: los destinatarios deberían haberlo recibido.",
                "No se pudo actualizar control_merge_pdfs.xlsx porque el formato del archivo no es el esperado "
                "(hoja Procesos, encabezados o fila 2). Pida a soporte restaurar o recrear el archivo de control.",
                "warning",
            )
        if ec in ("merge_control_read_failed", "merge_control_write_failed"):
            return (
                "Hizo bien el envío del correo: los destinatarios deberían haberlo recibido.",
                wtxt
                or "No se pudo leer o guardar control_merge_pdfs.xlsx (permisos o archivo abierto). "
                "Cierre ese Excel en SharePoint y reintente; si persiste, contacte soporte.",
                "warning",
            )
        if not result.get("merge_control_updated") and wtxt and not ec:
            return (
                "Hizo bien el envío del correo: los destinatarios deberían haberlo recibido.",
                wtxt,
                "warning",
            )
        return (
            "El correo de abonos del banco se envió correctamente con la tabla del día y los extractos configurados.",
            "Revise la bandeja de los destinatarios (y correo no deseado). Siguiente paso operativo: la secretaría "
            "carga los asientos en las carpetas del soporte; cuando termine, ejecute la unión de PDFs.",
            "success",
        )
    if job_type == "merge_composite_validado_pdfs":
        outputs = result.get("outputs") or []
        skipped = result.get("skipped") or []
        out_count = result.get("outputs_count")
        if out_count is None and isinstance(outputs, list):
            out_count = len(outputs)
        skip_count = result.get("skipped_count")
        if skip_count is None and isinstance(skipped, list):
            skip_count = len(skipped)
        if isinstance(outputs, list) and len(outputs) == 0 and isinstance(skipped, list) and len(skipped) > 0:
            return (
                "La unión de PDFs terminó sin generar ningún archivo: a todos los pagos les faltó el asiento contable, "
                "el extracto u otro requisito.",
                "Abra result.skipped en Power Automate (cada línea indica id_pago y motivo). "
                "La secretaría debe subir los PDF en ASIENTOS CONTABLES de cada crédito según el soporte de asientos. "
                "Corrija y ejecute Unir PDFs de nuevo.",
                "warning",
            )
        if isinstance(outputs, list) and len(outputs) > 0:
            if isinstance(skipped, list) and len(skipped) > 0:
                return (
                    f"Unió correctamente {out_count or len(outputs)} pago(s) en PDF consolidados; "
                    f"{skip_count or len(skipped)} pago(s) se omitieron por archivos faltantes.",
                    "Revise en SharePoint la carpeta 06 ASIENTO CONTABLES GENERADOS los PDF generados. "
                    "En result.skipped vea qué pagos faltan (asiento o extracto). Suba lo pendiente y reintente solo "
                    "para esos pagos si el flujo lo permite.",
                    "warning",
                )
            return (
                "La unión de PDFs terminó correctamente: cada pago validado quedó en un solo PDF "
                "(correo del día + asientos + extractos). Los asientos contables permanecen en sus carpetas.",
                "Revise en SharePoint la carpeta 06 ASIENTO CONTABLES GENERADOS. Siguiente paso: ejecutar la "
                "actualización de tablas de amortización cuando esa API exista (los asientos deben seguir disponibles).",
                "success",
            )
        return (
            "La unión de PDFs terminó; revise en result cuántos archivos se generaron y si hay pagos omitidos.",
            "Abra la carpeta 06 ASIENTO CONTABLES GENERADOS en SharePoint y, si aparece skipped en la respuesta, "
            "atienda los pagos listados antes de considerar el cierre completo.",
            "success",
        )
    if job_type == "amortization_dry_run":
        summary = result.get("summary") or {}
        total_events = summary.get("total_events") or summary.get("total") or 0
        errors = summary.get("errors") or 0
        if errors and total_events:
            return (
                f"El análisis preliminar terminó con {total_events} evento(s) revisados; "
                f"{errors} con error (revise items y error_code).",
                "Corrija asientos, tablas o manifest según cada ítem en result.items. "
                "No se modificó ninguna tabla de amortización.",
                "warning",
            )
        return (
            "El análisis preliminar de amortización terminó correctamente. No se modificó ninguna tabla.",
            "Revise el detalle del dry-run en result (items por asiento, application_status, filas destino). "
            "Si los valores, asientos y filas son correctos, el siguiente paso será ejecutar la "
            "actualización real cuando esté disponible.",
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
