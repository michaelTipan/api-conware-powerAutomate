"""
Enriquecimiento de documentos de job para respuestas HTTP homogéneas (Power Automate).

Solo afecta la salida de GET job; no modifica lógica de negocio ni almacenamiento interno
más allá de devolver una copia enriquecida desde los routers.
"""

from __future__ import annotations

import re
from typing import Any

from app.application.operational_message_policy import apply_audience_policy

ENRICHABLE_JOB_TYPES = frozenset(
    {
        "generate",
        "finalize",
        "notify_validar_extractos",
        "merge_composite_validado_pdfs",
        "amortization_dry_run",
        "amortization_apply",
    }
)

_UNKNOWN_USER = (
    "No fue posible completar el proceso debido a un inconveniente técnico."
)
_UNKNOWN_NEXT = (
    "No continúe con el siguiente paso. Contacte a soporte e indique el banco, "
    "la fecha y la etapa en la que se produjo el error."
)


def _mapped(code: str, user: str, next_a: str) -> tuple[str, str, str]:
    u, n = apply_audience_policy(code, user, next_a)
    return u, n, code

_PROCESS_CONTROL_FILES_HINT = (
    "el control de proceso del banco en 00 CONTROL "
    "(control_proceso_validacion_pagos_banco_bogota.xlsx o "
    "control_proceso_validacion_pagos_banco_bancolombia.xlsx)"
)

_GENERATE_MESSAGES: dict[str, tuple[str, str]] = {
    "review_folder_not_empty": (
        "No se pudo generar el archivo nuevo porque en 01 REVISION todavía hay un Excel de un día anterior.",
        "Mueva o archive los archivos validacion_pagos_*.xlsx viejos en 01 REVISION y vuelva a ejecutar la generación. "
        "Deje solo el reporte del banco actualizado en su carpeta.",
    ),
    "invalid_bank_code": (
        "No fue posible iniciar el proceso porque el banco indicado no es válido.",
        "No continúe con el siguiente paso. Contacte a soporte e indique el banco, la fecha y la etapa del proceso.",
    ),
    "active_process_exists": (
        "Ya existe un proceso activo en el control del banco y no se puede iniciar otro Generate.",
        "Finalice o cancele el proceso actual y vuelva a intentar.",
    ),
    "missing_sharepoint_folder": (
        "No fue posible completar el proceso debido a un inconveniente de configuración en SharePoint.",
        "No continúe con el siguiente paso. Contacte a soporte e indique el banco, la fecha y la etapa del proceso.",
    ),
    "bank_headers_not_found": (
        "El archivo BANCO_BOGOTA.xlsx no tiene las columnas que el sistema espera (fecha, monto/crédito, concepto).",
        "Revise el formato del reporte del banco. Debe coincidir con la plantilla habitual. "
        "Corrija el Excel, vuelva a subirlo y ejecute de nuevo la generación.",
    ),
    "tipo_aplicacion_column_missing": (
        "El archivo bancario no contiene la columna obligatoria Tipo Aplicación.",
        "Use la plantilla oficial, restaure la columna Tipo Aplicación y vuelva a ejecutar Generate.",
    ),
    "tipo_aplicacion_column_duplicate": (
        "El archivo bancario tiene más de una columna Tipo Aplicación; el sistema no puede determinar cuál usar.",
        "Deje una sola columna Tipo Aplicación en la plantilla oficial y vuelva a ejecutar Generate.",
    ),
    "tipo_aplicacion_required": (
        "Hay una transacción bancaria sin Tipo Aplicación (PAGO o ABONO).",
        "Complete Tipo Aplicación en todas las filas del reporte del banco y vuelva a ejecutar Generate.",
    ),
    "tipo_aplicacion_invalid": (
        "Hay una transacción bancaria con un Tipo Aplicación no válido.",
        "Use solo los valores visibles del banco: PAGO, PAGO Y ABONO CAPITAL, ABONO CAPITAL o ABONO MORA. "
        "Vuelva a ejecutar Generate.",
    ),
    "generic_abono_not_supported": (
        "El reporte del banco tiene un Tipo Aplicación ABONO genérico, que ya no se acepta.",
        "Clasifique cada movimiento como ABONO CAPITAL o ABONO MORA en el Excel del banco y vuelva a ejecutar Generate.",
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
        "Aún no se marcó el archivo como listo para procesar.",
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
        "En cada fila con pago, elija un valor de la lista: ADELANTADO, ATRASADO, NORMAL o REVISIÓN MANUAL. "
        "Guarde y vuelva a finalizar.",
    ),
    "invalid_estado_pago": (
        "Hay un Estado Pago escrito a mano o un valor que no está en la lista permitida.",
        "Use solo la lista desplegable: ADELANTADO, ATRASADO, NORMAL, REVISIÓN MANUAL. "
        "No escriba texto libre. Guarde y vuelva a finalizar.",
    ),
    "INCOMPLETO_NOT_SUPPORTED": (
        "El estado INCOMPLETO ya no forma parte del flujo operativo.",
        "Clasifique el movimiento como PAGO si cierra el extracto o como ABONO si corresponde a una aplicación parcial.",
    ),
    "estado_pago_no_finalizable": (
        "Quedan filas en REVISIÓN MANUAL sin resolver; no se puede cerrar el día.",
        "Revise esas filas en Distribución: cambie el estado cuando el caso esté resuelto o ajuste Validar Pago. "
        "Guarde y vuelva a finalizar.",
    ),
    "no_validar_requires_observation": (
        "Marcó Validar Pago = NO en una fila NORMAL pero no puso observación.",
        "En esa fila, escriba en Observación el motivo (por qué no se valida). Guarde y vuelva a finalizar.",
    ),
    "missing_valor_intereses": (
        "Falta Aplicar a extracto en una fila que debe validarse (Validar Pago = SI).",
        "Complete Aplicar a extracto (o 0 si no aplica). Revise que Saldo por asignar vaya quedando en cero. "
        "Guarde y vuelva a finalizar.",
    ),
    "missing_abono_k": (
        "Falta Abono a capital en una fila con Validar Pago = SI.",
        "Complete Abono a capital (o 0). Guarde y vuelva a finalizar.",
    ),
    "missing_abono_capital": (
        "Falta Abono a capital en una fila con Validar Pago = SI.",
        "Complete Abono a capital (o 0). Guarde y vuelva a finalizar.",
    ),
    "missing_mora_a_aplicar": (
        "Falta Mora a aplicar en una fila con Validar Pago = SI.",
        "Complete Mora a aplicar (o 0). Guarde y vuelva a finalizar.",
    ),
    "missing_otros_valores": (
        "Faltan Otros valores en una fila con Validar Pago = SI.",
        "Complete Otros valores (o 0). Guarde y vuelva a finalizar.",
    ),
    "review_schema_version_1_requires_regenerate": (
        "El archivo de revisión usa un esquema antiguo (ReviewSchemaVersion 1).",
        "Ejecute Generate de nuevo para obtener un workbook con columnas Mora a aplicar y "
        "Abono a capital separadas (versión 2).",
    ),
    "missing_mora": (
        "Faltan Otros valores en una fila con Validar Pago = SI.",
        "Complete Otros valores (o 0). Guarde y vuelva a finalizar.",
    ),
    "duplicate_bank_amount_in_payment_group": (
        "Monto banco está repetido en más de una fila del mismo ID Pago en Distribución.",
        "Deje Monto banco solo en la primera fila del pago (como en workbooks generados). "
        "Borre el valor duplicado en filas secundarias y vuelva a finalizar.",
    ),
    "amount_mismatch": (
        "Los valores distribuidos no coinciden con el monto registrado por el banco.",
        "Revise las columnas Aplicar a extracto, Mora a aplicar, Abono a capital y Otros valores en Distribucion_Pagos. "
        "Corrija los valores hasta que Saldo por asignar sea $0, guarde el archivo y vuelva a ejecutar la finalización.",
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
        "Ejecute primero la generación del archivo de revisión del día. Cuando exista el archivo y esté completo, ejecute la finalización.",
    ),
    "upload_failed": (
        "El proceso validó bien pero no pudo guardar el histórico o el soporte en SharePoint (red, permisos o archivo bloqueado).",
        "Cierre Excel en escritorio y en el navegador. Verifique espacio y permisos. Reintente Finalize; "
        "si persiste, contacte soporte con la hora del error.",
    ),
    "invalid_validar_abono": (
        "Hay un valor no permitido en Validar Abono (solo se acepta SI o NO).",
        "Abra Distribucion_Abonos y use la lista desplegable. Guarde y vuelva a finalizar.",
    ),
    "abono_without_selected_credit": (
        "Un abono no tiene ningún crédito seleccionado.",
        "Abra Distribucion_Abonos y marque SI en al menos un crédito para ese abono.",
    ),
    "abono_duplicate_selected_credit": (
        "Un abono tiene el mismo crédito seleccionado más de una vez.",
        "En Distribucion_Abonos deje solo una fila en SI por crédito para ese ID Pago.",
    ),
    "abono_group_inconsistent": (
        "Las filas seleccionadas de un abono no son consistentes (cliente, monto, fecha o tipo).",
        "Revise Distribucion_Abonos: todas las filas SI del mismo ID Pago deben coincidir en cliente, monto y fecha.",
    ),
    "abono_credit_without_unit_path": (
        "Un crédito de abono seleccionado no tiene ruta de unidad de crédito.",
        "Ejecute Generate de nuevo para regenerar rutas técnicas y vuelva a finalizar.",
    ),
    "abono_credit_without_amortization_path": (
        "Un crédito de abono seleccionado no tiene ruta de tabla de amortización.",
        "Ejecute Generate de nuevo y verifique la tabla en SharePoint antes de finalizar.",
    ),
    "abono_missing_bank_amount": (
        "Un abono seleccionado no tiene monto bancario.",
        "Revise Distribucion_Abonos o ejecute Generate de nuevo.",
    ),
    "abono_missing_bank_date": (
        "Un abono seleccionado no tiene fecha bancaria.",
        "Revise Distribucion_Abonos o ejecute Generate de nuevo.",
    ),
    "abono_invalid_application_type": (
        "Una fila de Distribucion_Abonos no está marcada correctamente como ABONO.",
        "Ejecute Generate de nuevo; no edite a mano las columnas técnicas TipoAplicacion o RequiereExtracto.",
    ),
    "abono_mora_missing_reference_extract": (
        "Un ABONO MORA seleccionado no tiene extracto de referencia.",
        "Revise Distribucion_Abonos: cada crédito ABONO MORA con Validar Abono = SI debe tener link o ruta de extracto. "
        "Guarde y vuelva a finalizar.",
    ),
    "abono_mora_missing_reference_date": (
        "Un ABONO MORA seleccionado no tiene fecha límite de referencia.",
        "Complete Fecha límite en Distribucion_Abonos para cada crédito ABONO MORA validado y vuelva a finalizar.",
    ),
    "missing_reference_extract_route": (
        "Un ABONO MORA validado no tiene ruta resoluble al extracto de referencia en SharePoint.",
        "Verifique el link de extracto en Distribucion_Abonos y que el PDF exista en la carpeta del crédito. "
        "Guarde y vuelva a finalizar.",
    ),
    "pago_y_abono_capital_missing_parte_cuota": (
        "Un PAGO Y ABONO CAPITAL validado no tiene parte de cuota (Aplicar a extracto) mayor que cero.",
        "Complete Aplicar a extracto en Distribución y vuelva a finalizar.",
    ),
    "pago_y_abono_capital_missing_capital": (
        "Un PAGO Y ABONO CAPITAL validado no tiene abono a capital mayor que cero "
        "(columna «Abono a capital» en Distribucion_Pagos).",
        "Complete Abono a capital en Distribución y vuelva a finalizar.",
    ),
    "pago_y_abono_capital_saldo_must_be_zero": (
        "Un PAGO Y ABONO CAPITAL validado no tiene Saldo por asignar en cero.",
        "Revise Aplicar a extracto, Abono a capital y Otros valores hasta que Saldo por asignar sea 0. "
        "Guarde y vuelva a finalizar.",
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
        "Marcó Validar Pago = SI pero el total aplicado es cero o negativo (NORMAL o ATRASADO).",
        "Ingrese los montos a aplicar o cambie Validar Pago a NO con observación. Guarde y vuelva a finalizar.",
    ),
    "invalid_bank_code": (
        "No fue posible finalizar el proceso porque el banco indicado no es válido.",
        "No continúe con el siguiente paso. Contacte a soporte e indique el banco, la fecha y la etapa del proceso.",
    ),
    "NO_READY_PROCESS": (
        "No hay ninguna revisión lista para finalizar.",
        "Verifique que en Control figure Procesar=SI y Estado=EN_REVISION en el Excel de revisión del banco. "
        "Si ya está marcado, ejecute la generación del archivo de revisión para ese banco.",
    ),
    "MULTIPLE_READY_PROCESSES": (
        "Hay más de un banco listo para finalizar.",
        "No continúe con el siguiente paso. Contacte a soporte e indique el banco, la fecha y la etapa del proceso.",
    ),
    "control_not_ready_for_finalize": (
        "El control del banco no está en estado REVISION_CREADA activo; no se puede finalizar.",
        "Ejecute Generate para iniciar un proceso nuevo o corrija el estado en el control por banco si aplica.",
    ),
    "missing_validation_file_path": (
        "El control del banco no tiene registrada la ruta del Excel de revisión.",
        "Ejecute la generación del archivo de revisión de nuevo para ese banco y vuelva a ejecutar la finalización.",
    ),
    "active_process_exists": (
        "Ya existe un proceso activo en el control del banco y no se puede finalizar otro proceso distinto.",
        "Finalice o cancele el proceso actual y vuelva a intentar.",
    ),
}

_GLOBAL_ERROR_MESSAGES: dict[str, tuple[str, str]] = {
    "amortization_table_ambiguous": (
        "Hay más de una tabla de amortización posible para un crédito y no fue posible determinar cuál usar.",
        "Deje una sola tabla de amortización por crédito en SharePoint y vuelva a ejecutar la generación del archivo de revisión. "
        "Si el error continúa, contacte a soporte.",
    ),
    "bank_amount_parse_error": (
        "No fue posible leer el monto bancario de un grupo; el valor no tiene formato numérico válido.",
        "Revise el monto en Distribucion_Pagos o Distribucion_Abonos y vuelva a ejecutar la finalización.",
    ),
    "bank_code_and_process_date_required": (
        "No fue posible continuar porque faltan datos internos del proceso.",
        "No continúe con la validación previa ni con la aplicación. Contacte a soporte para revisar el estado del proceso.",
    ),
    "bank_date_parse_error": (
        "No fue posible leer la fecha bancaria de un abono.",
        "Corrija la fecha en Distribucion_Abonos y vuelva a ejecutar la finalización.",
    ),
    "control_not_ready_for_dry_run": (
        "El control del banco aún no está listo para la validación previa de amortización.",
        "Complete la unión de PDF y verifique que el estado del proceso sea CONSOLIDADO o MERGE_PARCIAL antes de la validación previa.",
    ),
    "control_not_ready_for_merge": (
        "El control del banco no está listo para consolidar los soportes.",
        "Ejecute Finalizar validación de pagos (Flujo 2) y confirme que el correo de extractos se envió antes de "
        "Consolidación de soportes (Flujo 3).",
    ),
    "destination_name_exhausted": (
        "No se encontró un nombre disponible para guardar un PDF consolidado sin sobrescribir otro archivo.",
        "Revise la carpeta 06 ASIENTO CONTABLES GENERADOS, archive PDF antiguos si hace falta y vuelva a ejecutar la unión de PDF.",
    ),
    "missing_distribucion_abonos_headers": (
        "El histórico no tiene los encabezados esperados en Distribucion_Abonos.",
        "Vuelva a ejecutar la finalización con la plantilla actual; no edite los encabezados a mano.",
    ),
    "missing_distribucion_pagos_sheet": (
        "El histórico no tiene la hoja Distribucion_Pagos necesaria para este paso.",
        "Vuelva a ejecutar la finalización con el archivo generado por el sistema (no utilice una copia manual).",
    ),
    "missing_merge_manifest_path": (
        "No fue posible localizar el registro interno de la unión de documentos.",
        "No continúe con la validación previa ni con la aplicación. Contacte a soporte para revisar el estado del proceso.",
    ),
    "pdf_no_text": (
        "El PDF cargado no contiene texto legible para la automatización.",
        "Verifique que el documento sea correcto y que su contenido pueda seleccionarse. "
        "Cargue una copia legible y vuelva a ejecutar la validación previa. Si el error continúa, contacte a soporte.",
    ),
    "preflight_errors": (
        "La validación previa detectó errores que impiden continuar.",
        "Revise los asientos, extractos y tablas indicados en el resumen de la validación previa. "
        "Corrija los documentos y vuelva a ejecutar la validación previa.",
    ),
    "preflight_revision_manual": (
        "La validación previa requiere revisión manual de uno o más movimientos.",
        "Revise el Excel de revisión o el histórico según el tipo de problema y vuelva a ejecutar la validación previa.",
    ),
    "preflight_warnings_not_allowed": (
        "La validación previa encontró advertencias que no están permitidas para continuar.",
        "Corrija las advertencias señaladas en el resumen y vuelva a ejecutar la validación previa.",
    ),
    "process_control_invalid_structure": (
        "El control de proceso del banco no tiene la estructura esperada.",
        "No continúe con el siguiente paso. Contacte a soporte para revisar el estado del proceso.",
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
        return _mapped("unknown_error", _UNKNOWN_USER, _UNKNOWN_NEXT)

    if code in _GLOBAL_ERROR_MESSAGES:
        u, n = _GLOBAL_ERROR_MESSAGES[code]
        return _mapped(code, u, n)
    if code in table:
        u, n = table[code]
        return _mapped(code, u, n)
    hit = _pick_table_code(table, _strip_exception_prefix(full_message))
    if hit and hit in table:
        u, n = table[hit]
        return _mapped(hit, u, n)
    resolved = code if code and code != "unknown_error" else "unknown_error"
    return _mapped(resolved, _UNKNOWN_USER, _UNKNOWN_NEXT)


def _notify_merge_string_mapping(job_type: str, msg: str) -> tuple[str, str, str]:
    mlow = msg.lower()
    base_code = msg.strip().split("|", 1)[0].strip()
    if base_code in _GLOBAL_ERROR_MESSAGES:
        u, n = _GLOBAL_ERROR_MESSAGES[base_code]
        return _mapped(base_code, u, n)

    if job_type == "notify_validar_extractos":
        mstripped = msg.strip()
        if mstripped == "NO_READY_PROCESS":
            return (
                "No hay ningún histórico listo para notificar.",
                "Ejecute Finalize del banco correspondiente hasta que el control por banco quede en FINALIZADO. "
                "Luego ejecute el envío del correo sin parámetros adicionales o contacte a soporte si hay varios bancos activos.",
                "NO_READY_PROCESS",
            )
        if mstripped.startswith("MULTIPLE_READY_PROCESSES"):
            return _mapped(
                "MULTIPLE_READY_PROCESSES",
                "Hay más de un banco listo para notificar.",
                "No continúe con el siguiente paso. Contacte a soporte e indique el banco, la fecha y la etapa del proceso.",
            )
        if mstripped == "control_not_ready_for_notify":
            return (
                "El control del banco no está listo para enviar el correo de extractos.",
                "Si la revisión aún no se cerró, ejecute Finalizar validación de pagos (Flujo 2). "
                "Si el proceso ya avanzó, contacte a soporte.",
                "control_not_ready_for_notify",
            )
        if mstripped == "already_notified":
            return (
                "El correo de extractos ya fue enviado y registrado para este proceso.",
                "Revise la bandeja de los destinatarios. El siguiente paso es ejecutar "
                "Consolidación de soportes (Flujo 3).",
                "already_notified",
            )
        if mstripped == "missing_historical_file_path" or mstripped.startswith(
            "missing_historical_file_path|"
        ):
            return _mapped(
                "missing_historical_file_path",
                "No se pudo resolver el histórico para el correo (falta la ruta del Excel cartera_validada).",
                "Ejecute Finalizar validación de pagos (Flujo 2) para ese banco y, si el error continúa, contacte a soporte.",
            )
        if mstripped.startswith("historical_file_not_found"):
            return (
                "No se pudo abrir el archivo histórico de validación en SharePoint (no existe, fue movido o sin permiso).",
                "Confirme que el Flujo 2 terminó correctamente y que el histórico del día sigue en 02 HISTORICO. "
                "Si fue movido o borrado, contacte a soporte antes de volver a ejecutar el Flujo 2.",
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
                "Abra el histórico en Distribucion_Pagos y confirme que haya pagos validados para el día. "
                "Si no se marcaron filas, corrija el Excel de revisión y vuelva a ejecutar la finalización.",
                "no_validated_rows",
            )
        if "no hay destinatarios" in mlow or (
            "receptores" in mlow and "vacía" in mlow
        ):
            return (
                "El correo no se envió porque no hay destinatarios válidos (columna RECEPTORES vacía o correos mal escritos).",
                "Abra 00 CONTROL / CORREOS.xlsx: columna RECEPTORES debe tener al menos un correo por fila. "
                "Guarde el archivo y vuelva a ejecutar Finalizar validación de pagos (Flujo 2).",
                "recipients_not_configured",
            )
        if "emisor" in mlow or "receptores" in mlow or "correos" in mlow or "remitente" in mlow:
            return (
                "Falta configurar quién envía o quién recibe el correo en CORREOS.xlsx (EMISOR y RECEPTORES).",
                "En 00 CONTROL / CORREOS.xlsx complete EMISOR (un correo) y RECEPTORES (uno o más correos). "
                "Guarde el archivo en SharePoint y vuelva a ejecutar Finalizar validación de pagos (Flujo 2).",
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
                "No continúe con el siguiente paso. Contacte a soporte e indique el banco, la fecha y la etapa del proceso.",
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
                "No fue posible enviar el correo de extractos (permisos del buzón, remitente o destinatarios).",
                "Verifique en CORREOS.xlsx que EMISOR sea un buzón autorizado y que RECEPTORES tenga correos válidos. "
                "Luego vuelva a ejecutar Finalizar validación de pagos (Flujo 2). Si persiste, contacte a soporte.",
                "graph_sendmail_failed",
            )

    if job_type == "merge_composite_validado_pdfs":
        mstripped = msg.strip()
        if mstripped == "merge_control_no_pending_process":
            return (
                "No se puede unir PDFs porque el sistema no tiene un proceso activo registrado tras el correo del día.",
                "Ejecute en este orden: Finalize → envío de correo de extractos (debe quedar "
                f"PENDIENTE_ASIENTOS en {_PROCESS_CONTROL_FILES_HINT}). Luego ejecute Unir PDFs. "
                "Si el correo de hoy no actualizó el control del banco, reenvíe el correo o pida a soporte.",
                "merge_control_no_pending_process",
            )
        if mstripped == "missing_historical_file_path":
            return (
                "El archivo de control no tiene la ruta del histórico de validación necesaria para unir PDFs.",
                f"Vuelva a ejecutar el envío de correo del día (debe registrar HistoricalFilePath en "
                f"{_PROCESS_CONTROL_FILES_HINT}). "
                "Si el control del banco está vacío o dañado, ejecute setup de controles por banco o pida a soporte.",
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
                f"No existe el control de proceso del banco en 00 CONTROL; sin {_PROCESS_CONTROL_FILES_HINT} "
                "no puede iniciar la unión de PDFs.",
                "No continúe con el siguiente paso. Contacte a soporte para revisar la configuración del control de proceso.",
                "merge_control_workbook_not_found",
            )
        if mstripped == "merge_control_invalid_structure":
            return (
                f"El control de proceso del banco está dañado o fue editado (falta hoja Procesos, encabezados o fila 2).",
                f"No modifique {_PROCESS_CONTROL_FILES_HINT} a mano. Ejecute setup o pida a soporte y repita correo + Unir PDFs.",
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
                "En el histórico, hoja Distribucion_Pagos, confirme que haya filas validadas para el día. "
                "Si la revisión no quedó cerrada correctamente, corrija el Excel y vuelva a ejecutar la finalización antes del correo y este paso.",
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
                "Ejecute la finalización de la revisión de nuevo para regenerar esa columna. Debe haberse completado la revisión antes.",
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

    if job_type in ("amortization_dry_run", "amortization_apply"):
        mstripped = msg.strip()
        if mstripped == "merge_control_manifest_path_missing":
            return _mapped(
                "merge_control_manifest_path_missing",
                "No fue posible localizar el registro interno de la unión de documentos en el control del banco.",
                "No continúe con la validación previa ni con la aplicación. Contacte a soporte para revisar el estado del proceso.",
            )
        if mstripped == "merge_control_amortization_not_ready":
            return (
                "El proceso de Merge aún no está listo para amortización (estado distinto de CONSOLIDADO/MERGE_PARCIAL).",
                "Espere a que Merge finalice o corrija el estado en el control del banco (CONSOLIDADO/MERGE_PARCIAL) "
                "antes de aplicar amortización.",
                "merge_control_amortization_not_ready",
            )
        if mstripped == "merge_control_workbook_not_found":
            return (
                f"No existe el control de proceso del banco en SharePoint.",
                "Ejecute setup de controles por banco y el flujo Notify/Merge antes de amortización.",
                "merge_control_workbook_not_found",
            )
        if mstripped in (
            "merge_control_invalid_structure",
            "missing_historical_file_path",
        ):
            return (
                f"El control de proceso del banco no tiene la estructura esperada o faltan rutas.",
                "Ejecute setup de controles por banco o contacte soporte; no edite encabezados a mano.",
                mstripped,
            )

    if job_type == "amortization_dry_run":
        mstripped = msg.strip()
        if mstripped == "report_date_iso_required" or mstripped.startswith("report_date_iso_required"):
            return _mapped(
                "report_date_iso_required",
                "No fue posible continuar porque faltan datos internos del proceso de unión de documentos.",
                "No continúe con la validación previa ni con la aplicación. Contacte a soporte para revisar el estado del proceso.",
            )
        if mstripped.startswith("merge_manifest_not_found"):
            return _mapped(
                "merge_manifest_not_found",
                "No se encontró el registro de la unión de documentos para la fecha indicada.",
                "Ejecute la unión de PDF para ese día y vuelva a ejecutar la validación previa. "
                "Si el error continúa, contacte a soporte.",
            )
        if mstripped.startswith("invalid_merge_manifest_json"):
            return _mapped(
                "invalid_merge_manifest_json",
                "El registro de la unión de documentos existe pero no tiene un formato válido.",
                "No continúe con el siguiente paso. Contacte a soporte e indique el banco, la fecha y la etapa del proceso.",
            )
        if "graph_config" in mlow or "missing environment variable" in mlow:
            return _mapped(
                "graph_config_error",
                "No fue posible completar el proceso debido a un inconveniente de configuración.",
                "No continúe con el siguiente paso. Contacte a soporte e indique el banco, la fecha y la etapa del proceso.",
            )

    return _mapped("unknown_error", _UNKNOWN_USER, _UNKNOWN_NEXT)


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

    user, next_a = apply_audience_policy(code, user, next_a)
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
        custom_um = str(result.get("user_message") or "").strip()
        custom_na = str(result.get("next_action") or "").strip()
        estado_control = str(result.get("process_control_estado") or "")
        sev = "warning" if estado_control == "REVISION_REQUIERE_CORRECCION" else "success"
        if custom_um:
            return (
                custom_um,
                custom_na
                or "Abra el Excel en 01 REVISION, complete las hojas de distribución y en Control ponga Procesar = SI.",
                sev,
            )
        if estado_control == "REVISION_REQUIERE_CORRECCION":
            return (
                "Se generó el archivo de revisión, pero la hoja Errores contiene registros que requieren atención.",
                "Revise la hoja Errores, corrija los datos bancarios o documentos, y vuelva a ejecutar la generación.",
                sev,
            )
        return (
            "Se generó el archivo de revisión del día. Ya puede abrirlo en la carpeta 01 REVISION de SharePoint.",
            "Abra ese Excel, complete Distribucion_Pagos (Estado Pago y Validar Pago en cada fila) y en la hoja Control "
            "marque Procesar = SI cuando termine. Luego ejecute la finalización de la revisión.",
            sev,
        )
    if job_type == "finalize":
        custom_um = str(result.get("user_message") or "").strip()
        custom_na = str(result.get("next_action") or "").strip()
        if custom_um:
            return (
                custom_um,
                custom_na
                or "Abra el soporte de asientos y cargue los PDF en cada carpeta ASIENTOS CONTABLES.",
                "success",
            )
        return (
            "Se finalizó la revisión correctamente. Se guardó el histórico del día y el soporte "
            "para cargar los asientos contables.",
            "Abra Asientos_Pendientes, cargue cada PDF en la carpeta ASIENTOS CONTABLES del crédito correspondiente "
            "y continúe con el envío del correo de extractos cuando corresponda.",
            "success",
        )
    if job_type == "notify_validar_extractos":
        ec = str(result.get("merge_control_error_code") or "").strip()
        wtxt = str(result.get("merge_control_warning") or "").strip()
        if ec == "merge_control_active_process_exists":
            return (
                "Se envió el correo correctamente; los destinatarios deberían haberlo recibido.",
                f"Para el siguiente paso (unir PDFs): en 00 CONTROL el control del banco ya tiene un proceso "
                "pendiente. Termine o cancele ese proceso antes de volver a registrar uno nuevo.",
                "warning",
            )
        if ec == "missing_email_pdf_path_for_merge_control":
            return (
                "Se envió el correo correctamente; los destinatarios deberían haberlo recibido.",
                "No se guardó el PDF copia del correo en 05 EMAIL ni se registró el paso para unir PDFs después. "
                "Revise en SharePoint la carpeta 05 EMAIL y que la exportación del PDF esté activa; luego reintente "
                "solo el registro en control o contacte soporte antes de ejecutar Unir PDFs.",
                "warning",
            )
        if ec == "merge_control_workbook_not_found":
            return (
                "Se envió el correo correctamente; los destinatarios deberían haberlo recibido.",
                f"No se actualizó el control de proceso del banco porque no existe en 00 CONTROL. "
                "Ejecute setup de controles por banco; después el flujo de correo quedará listo para unir PDFs.",
                "warning",
            )
        if ec == "merge_control_invalid_structure":
            return (
                "Se envió el correo correctamente; los destinatarios deberían haberlo recibido.",
                f"No se pudo actualizar el control de proceso del banco: formato no esperado "
                "(hoja Procesos, encabezados o fila 2). Ejecute setup o pida a soporte restaurar el control del banco.",
                "warning",
            )
        if ec in ("merge_control_read_failed", "merge_control_write_failed"):
            return (
                "Se envió el correo correctamente; los destinatarios deberían haberlo recibido.",
                wtxt
                or f"No se pudo leer o guardar el control de proceso del banco (permisos o archivo abierto). "
                "Cierre el Excel del banco en SharePoint y reintente; si persiste, contacte soporte.",
                "warning",
            )
        if not result.get("merge_control_updated") and wtxt and not ec:
            return (
                "Se envió el correo correctamente; los destinatarios deberían haberlo recibido.",
                wtxt,
                "warning",
            )
        abono_groups = int(result.get("abono_groups_included") or 0)
        if abono_groups > 0:
            return (
                "El correo de movimientos bancarios se envió correctamente. Los pagos incluyeron sus extractos "
                "y los abonos se reportaron sin extracto, según corresponde.",
                "Revise la bandeja de los destinatarios (y correo no deseado). Cargue los asientos en las carpetas del soporte; "
                "cuando termine, ejecute Consolidación de soportes (Flujo 3).",
                "success",
            )
        return (
            "El correo de abonos del banco se envió correctamente con la tabla del día y los extractos configurados.",
            "Revise la bandeja de los destinatarios (y correo no deseado). Cargue los asientos en las carpetas del soporte; "
            "cuando termine, ejecute Consolidación de soportes (Flujo 3).",
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
                f"La unión de PDFs terminó sin generar ningún archivo: los {skip_count or len(skipped)} pago(s) "
                "requieren documentos faltantes (asiento contable, extracto u otro requisito).",
                "Revise Asientos_Pendientes y cargue los PDF en ASIENTOS CONTABLES de cada crédito según el soporte de asientos. "
                "Corrija y vuelva a ejecutar Consolidación de soportes (Flujo 3).",
                "warning",
            )
        incomplete_count = int(result.get("incomplete_groups_count") or 0)
        if incomplete_count > 0:
            return (
                f"La consolidación quedó incompleta: {incomplete_count} grupo(s) con documentos obligatorios faltantes.",
                "Revise Asientos_Pendientes y la carpeta de extractos del crédito. Cargue los asientos o extractos faltantes "
                "y vuelva a ejecutar Consolidación de soportes (Flujo 3) antes del Flujo 4.",
                "warning",
            )
        if isinstance(outputs, list) and len(outputs) > 0:
            if isinstance(skipped, list) and len(skipped) > 0:
                return (
                    f"Se unieron {out_count or len(outputs)} grupo(s) completo(s); "
                    f"quedan {skip_count or len(skipped)} omisión(es) documentales.",
                    "Revise Asientos_Pendientes y cargue los documentos pendientes. "
                    "Vuelva a ejecutar Consolidación de soportes (Flujo 3).",
                    "warning",
                )
            abono_out = int(result.get("abono_outputs_count") or 0)
            if abono_out > 0:
                return (
                    "La consolidación generó los soportes de pagos y abonos. Los abonos se consolidaron con sus "
                    "asientos contables, sin exigir extractos.",
                    "Revise en SharePoint la carpeta 06 ASIENTO CONTABLES GENERADOS y, cuando confirme que los PDF "
                    "consolidados están completos, ejecute Llenar tabla de amortización (Flujo 4).",
                    "success",
                )
            return (
                "La consolidación de soportes terminó correctamente: cada pago validado quedó en un solo PDF "
                "(correo del día + asientos + extractos).",
                "Revise en SharePoint la carpeta 06 ASIENTO CONTABLES GENERADOS y, cuando confirme que los PDF "
                "consolidados están completos, ejecute Llenar tabla de amortización (Flujo 4).",
                "success",
            )
        return (
            "La unión de PDFs terminó. Revise en SharePoint la carpeta 06 ASIENTO CONTABLES GENERADOS.",
            "Si quedaron pagos omitidos, revise Asientos_Pendientes y cargue los documentos faltantes antes de continuar.",
            "success",
        )
    if job_type == "amortization_dry_run":
        custom_um = str(result.get("user_message") or "").strip()
        custom_na = str(result.get("next_action") or "").strip()
        if custom_um:
            sev = "warning" if result.get("requires_business_rule") else "success"
            if not result.get("can_apply", True):
                sev = "warning"
            return (
                custom_um,
                custom_na
                or "Revise el resumen de abonos y los detalles de la validación previa antes de aplicar pagos y abonos.",
                sev,
            )
        summary = result.get("summary") or {}
        total_events = summary.get("total_events") or summary.get("total") or 0
        errors = summary.get("errors") or 0
        abono_not_reconciled = int(result.get("abono_groups_not_reconciled") or 0)
        if errors and total_events:
            return (
                f"La validación previa terminó con {total_events} evento(s) revisados; "
                f"{errors} con errores que impiden continuar.",
                "Corrija asientos, extractos o tablas según el resumen de la validación previa. "
                "No se modificó ninguna tabla de amortización.",
                "warning",
            )
        if abono_not_reconciled > 0:
            return (
                f"La validación previa detectó {abono_not_reconciled} grupo(s) de abono con errores de documentos o cuadre.",
                "Revise el resumen de abonos (montos y documentos faltantes). "
                "Cargue los asientos faltantes en la carpeta del crédito; no es necesario extracto para abonos.",
                "warning",
            )
        if result.get("can_apply") and int(result.get("abono_groups_ready") or 0) > 0:
            return (
                "La validación previa de abonos cuadró correctamente. Cada abono utilizará la siguiente fila "
                "libre de Aplicación de Pagos; el IBR no se modificará.",
                "Revise los detalles de cada asiento y fila planificada; si todo cuadra, ejecute la aplicación de pagos y abonos.",
                "success",
            )
        return (
            "La validación previa terminó correctamente y no se modificó ninguna tabla.",
            "Revise el resumen de la validación previa. Si todo cuadra, ejecute la aplicación de pagos y abonos.",
            "success",
        )
    if job_type == "amortization_apply":
        if str(result.get("status") or "") == "blocked":
            custom_um = str(result.get("user_message") or "").strip()
            custom_na = str(result.get("next_action") or "").strip()
            if custom_um:
                return (
                    custom_um,
                    custom_na
                    or "Revise los grupos de abono bloqueados antes de volver a aplicar pagos y abonos.",
                    "warning",
                )
            return (
                "La amortización no puede aplicarse: hay grupos de abono bloqueados.",
                "Revise los grupos de abono bloqueados y corrija asientos o documentos antes de volver a aplicar pagos y abonos.",
                "warning",
            )
        if result.get("already_applied"):
            custom_um = str(result.get("user_message") or "").strip()
            custom_na = str(result.get("next_action") or "").strip()
            if custom_um:
                return (custom_um, custom_na, "success")
        if str(result.get("status") or "") == "preflight_failed":
            return (
                "No fue posible actualizar las tablas de amortización porque se detectaron errores en la revisión interna.",
                "Revise los documentos y tablas indicados, corrija el inconveniente y vuelva a ejecutar "
                "Llenar tabla de amortización (Flujo 4).",
                "warning",
            )
        if str(result.get("status") or "") in ("ok", "partial"):
            custom_um = str(result.get("user_message") or "").strip()
            custom_na = str(result.get("next_action") or "").strip()
            sev = "success" if result.get("status") == "ok" else "warning"
            if custom_um:
                return (
                    custom_um,
                    custom_na
                    or (
                        "Abra cada tabla actualizada y confirme que los pagos aplicados y el cronograma "
                        "quedaron correctos. Este es el último paso automático del proceso."
                        if result.get("status") == "ok"
                        else "Revise las tablas pendientes, corrija el inconveniente y vuelva a ejecutar "
                        "Llenar tabla de amortización (Flujo 4)."
                    ),
                    sev,
                )
            tables_n = int(result.get("tables_uploaded_count") or 0)
            if result.get("status") == "partial":
                return (
                    f"Se actualizaron {tables_n} tabla(s), pero quedaron tablas pendientes de revisión.",
                    "Revise las tablas pendientes, corrija el inconveniente y vuelva a ejecutar "
                    "Llenar tabla de amortización (Flujo 4).",
                    "warning",
                )
            return (
                "El proceso de validación de pagos finalizó correctamente. "
                f"Se actualizaron {tables_n} tabla(s) de amortización en SharePoint.",
                "Abra cada tabla actualizada y confirme que los pagos aplicados y el cronograma "
                "quedaron correctos. Este es el último paso automático del proceso.",
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
