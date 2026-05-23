"""
Fuente única de verdad para los nombres de hojas y columnas del Excel de revisión.

Generate escribe usando estas constantes.
Finalize lee usando estas mismas constantes.
Si necesitas cambiar un nombre de columna, cámbialo AQUÍ, no en cada archivo.
"""
from __future__ import annotations

import re
from typing import Any

# Nombres de hojas
class ReviewSheets:
    CONTROL = "Control"
    RESUMEN = "Resumen"
    CASOS_PAGO = "Casos_Pago"
    DISTRIBUCION = "Distribucion"
    LISTAS = "_Listas"
    ERRORES = "Errores"


# Columnas de hoja Control
class ControlCols:
    CAMPO = "Campo"
    VALOR = "Valor"
    # Valores de filas conocidos
    ROW_PROCESAR = "Procesar"
    ROW_ESTADO_PROCESO = "Estado proceso"
    ROW_FECHA_PROCESAMIENTO = "Fecha procesamiento"
    ROW_RESULTADO = "Resultado"
    ROW_ID_PROCESO = "ID Proceso"
    ROW_ESTADO = "Estado"
    ROW_FECHA = "Fecha"
    # Valores esperados
    VAL_PROCESAR_SI = "SI"
    VAL_PROCESAR_NO = "NO"
    VAL_PROCESADO = "PROCESADO"
    VAL_FINALIZADO = "FINALIZADO"
    ESTADO_OPTIONS = [
        "EN_REVISION",
        "PROCESANDO",
        "PROCESADO",
        "ERROR",
        "CANCELADO",
    ]
    PROCESAR_OPTIONS = [VAL_PROCESAR_NO, VAL_PROCESAR_SI]


# Columnas de hoja Casos_Pago
class CasosPagoCols:
    ID_PAGO = "ID Pago"
    FECHA_BANCO = "Fecha banco"
    CLIENTE = "Cliente"
    CONCEPTO_BANCO = "Concepto banco"
    MONTO_BANCO = "Monto banco"
    OBSERVACION = "Observación"

    HEADERS = [
        "ID Pago",
        "Fecha banco",
        "Cliente",
        "Concepto banco",
        "Monto banco",
        "Observación",
    ]


class ValidarPago:
    SI = "SI"
    NO = "NO"


class EstadoPago:
    """Valores permitidos en columna «Estado Pago» (Distribución)."""

    ADELANTADO = "ADELANTADO"
    ATRASADO = "ATRASADO"
    INCOMPLETO = "INCOMPLETO"
    NORMAL = "NORMAL"
    REVISION_MANUAL = "REVISION_MANUAL"

    ALLOWED = frozenset(
        {
            ADELANTADO,
            ATRASADO,
            INCOMPLETO,
            NORMAL,
            REVISION_MANUAL,
        }
    )

    OPTIONS_ORDERED = [
        ADELANTADO,
        ATRASADO,
        INCOMPLETO,
        NORMAL,
        REVISION_MANUAL,
    ]

    # Finalizar solo bloquea revisión manual pendiente (caso no listo para cierre operativo)
    FINALIZE_FORBIDDEN = frozenset({REVISION_MANUAL})

    # Suma aplicada / filas que consolidan en histórico y tabla de amortización
    # (incluye ATRASADO y ADELANTADO: adelanto sigue requiriendo seguimiento operativo)
    COUNTERS_POSITIVE_TOTAL = frozenset({NORMAL, INCOMPLETO, ATRASADO, ADELANTADO})

    # Rutas extracto / soporte secretaría (antes «VALIDAR»)
    SECRETARY_AND_RUTA = frozenset({NORMAL, INCOMPLETO, ATRASADO, ADELANTADO})

    # Estados con seguimiento operativo positivo (pagos_adelantados / pagos_incompletos vía followup)
    CLEARS_PENDING = COUNTERS_POSITIVE_TOTAL


# Columnas de hoja Distribucion
class DistribucionCols:
    ID_PAGO = "ID Pago"
    CLIENTE = "Cliente"
    CREDITO = "Crédito"
    MONTO_BANCO = "Monto banco"
    FECHA_BANCO = "Fecha banco"
    FECHA_LIMITE = "Fecha límite"
    DIAS_MORA = "Días mora"
    VALOR_EXTRACTO = "Valor extracto"
    APLICAR_A_EXTRACTO = "Aplicar a extracto"
    MORA_A_APLICAR = "Mora a aplicar"
    OTROS_VALORES = "Otros valores"
    TOTAL_APLICADO = "Total aplicado"
    SALDO_POR_ASIGNAR = "Saldo por asignar"
    ESTADO_PAGO = "Estado Pago"
    VALIDAR_PAGO = "Validar Pago"
    OBSERVACION = "Observación"
    LINK_EXTRACTO = "Link extracto"
    LINK_TABLA = "Link tabla amortización"
    LINK_CARPETA_CREDITO = "Link carpeta crédito"
    RUTA = "Ruta"
    RUTA_UNIDAD_CREDITO = "RutaUnidadCredito"
    # Solo en cartera_validada (Finalize); no en validacion_pagos de Generate
    RUTA_ASIENTOS_CONTABLES = "RutaAsientosContables"
    # Columnas técnicas (ocultas en Excel); ruta estable para amortización / dry-run
    RUTA_TABLA_AMORTIZACION = "RutaTablaAmortizacion"
    CREDITO_NORMALIZADO = "CreditoNormalizado"
    # Alias de código (mismo texto visible que las columnas renombradas)
    VALOR_INTERESES = APLICAR_A_EXTRACTO
    ABONO_K = MORA_A_APLICAR
    INTERESES_MORA = OTROS_VALORES

    # Alias retrocompat lecturas (misma cadena que ESTADO_PAGO)
    ESTADO_LINEA = ESTADO_PAGO

    HEADERS = [
        ID_PAGO,
        CLIENTE,
        CREDITO,
        MONTO_BANCO,
        FECHA_BANCO,
        FECHA_LIMITE,
        DIAS_MORA,
        VALOR_EXTRACTO,
        APLICAR_A_EXTRACTO,
        MORA_A_APLICAR,
        OTROS_VALORES,
        TOTAL_APLICADO,
        SALDO_POR_ASIGNAR,
        ESTADO_PAGO,
        VALIDAR_PAGO,
        OBSERVACION,
        LINK_EXTRACTO,
        LINK_TABLA,
        LINK_CARPETA_CREDITO,
        RUTA,
        RUTA_UNIDAD_CREDITO,
        RUTA_TABLA_AMORTIZACION,
        CREDITO_NORMALIZADO,
    ]


# Columnas técnicas de Distribución que deben quedar ocultas en el Excel
DISTRIBUCION_TECHNICAL_HIDDEN_COLUMNS = frozenset(
    {
        DistribucionCols.RUTA_TABLA_AMORTIZACION,
        DistribucionCols.CREDITO_NORMALIZADO,
    }
)


def normalize_credito_digits(raw: Any) -> str:
    """Número de crédito limpio (ej. 258) desde etiqueta visible o celda."""
    s = str(raw or "").strip()
    if not s:
        return ""
    m = re.search(r"(?i)credito#?\s*(\d+)", s)
    if m:
        return m.group(1)
    m2 = re.search(r"\d{1,12}", s)
    return m2.group(0) if m2 else ""


# Encabezados legacy en libros generados antes del renombre UX (Finalize los normaliza)
DISTRIB_LEGACY_HEADER_ALIASES: dict[str, str] = {
    "Valor intereses": DistribucionCols.APLICAR_A_EXTRACTO,
    "Abono a K": DistribucionCols.MORA_A_APLICAR,
    "Intereses de mora": DistribucionCols.OTROS_VALORES,
    "Estado línea": DistribucionCols.ESTADO_PAGO,
    "Estado": DistribucionCols.ESTADO_PAGO,
}


def normalize_distrib_row_keys(row: dict[str, Any]) -> dict[str, Any]:
    """Rekey filas leídas por encabezado visible hacia constantes canónicas del schema."""
    out: dict[str, Any] = dict(row)
    for legacy, canonical in DISTRIB_LEGACY_HEADER_ALIASES.items():
        if legacy in out and canonical not in out:
            out[canonical] = out[legacy]
    return out


class EstadoLinea:
    """Tokens legacy en columna «Estado» antes de Estado Pago + Validar Pago (solo migración)."""

    VALIDAR = "VALIDAR"
    VALIDAR_PARCIAL = "VALIDAR_PARCIAL"
    REPROGRAMAR = "REPROGRAMAR"
    NO_VALIDAR = "NO_VALIDAR"
    PENDIENTE_MORA = "PENDIENTE_MORA"
    REVISION_MANUAL = "REVISION_MANUAL"
    OPTIONS = [
        VALIDAR,
        VALIDAR_PARCIAL,
        PENDIENTE_MORA,
        REPROGRAMAR,
        NO_VALIDAR,
        REVISION_MANUAL,
    ]


_LEGACY_ESTADO_TO_PAIR: dict[str, tuple[str, str]] = {
    EstadoLinea.VALIDAR: (EstadoPago.NORMAL, ValidarPago.SI),
    EstadoLinea.VALIDAR_PARCIAL: (EstadoPago.INCOMPLETO, ValidarPago.SI),
    EstadoLinea.REPROGRAMAR: (EstadoPago.ADELANTADO, ValidarPago.SI),
    EstadoLinea.NO_VALIDAR: (EstadoPago.NORMAL, ValidarPago.NO),
    EstadoLinea.PENDIENTE_MORA: (EstadoPago.ATRASADO, ValidarPago.SI),
    EstadoLinea.REVISION_MANUAL: (EstadoPago.REVISION_MANUAL, ValidarPago.NO),
}


def _normalize_validar_pago_raw(raw: Any) -> str:
    s = str(raw or "").strip().upper()
    if s in ("SI", "SÍ", "YES", "TRUE", "1", "Y"):
        return ValidarPago.SI
    if s in ("NO", "N", "FALSE", "0"):
        return ValidarPago.NO
    return ""


def normalize_validar_pago_value(raw: Any) -> str:
    """Devuelve ``SI``, ``NO`` o cadena vacía si no se reconoce."""
    return _normalize_validar_pago_raw(raw)


def is_validar_pago_si(row: dict[str, Any]) -> bool:
    return normalize_validar_pago_value(row.get(DistribucionCols.VALIDAR_PAGO)) == ValidarPago.SI


def apply_legacy_estado_migration(row: dict[str, Any]) -> None:
    """
    Normaliza en sitio: token legacy en «Estado Pago» → par (Estado Pago, Validar Pago).
    Si ya es un EstadoPago permitido, solo rellena Validar Pago faltante con NO.
    """
    raw_ep = row.get(DistribucionCols.ESTADO_PAGO)
    token = str(raw_ep or "").strip().upper()
    vp = _normalize_validar_pago_raw(row.get(DistribucionCols.VALIDAR_PAGO))

    if token in EstadoPago.ALLOWED:
        row[DistribucionCols.ESTADO_PAGO] = token
        row[DistribucionCols.VALIDAR_PAGO] = vp if vp else ValidarPago.NO
        return

    pair = _LEGACY_ESTADO_TO_PAIR.get(token)
    if pair:
        row[DistribucionCols.ESTADO_PAGO], row[DistribucionCols.VALIDAR_PAGO] = pair
        return

    # Token desconocido: dejar tal cual para que Finalize rechace con invalid_estado_pago
    row[DistribucionCols.ESTADO_PAGO] = str(raw_ep or "").strip()
    if not vp:
        row[DistribucionCols.VALIDAR_PAGO] = ValidarPago.NO


# Columnas de hoja Errores (bandeja operativa secretaría + código técnico al final para soporte)
class ErroresCols:
    ID_PAGO = "ID Pago"
    CLIENTE = "Cliente"
    CREDITO = "Crédito"
    TIPO_CASO = "Tipo de caso"
    DESCRIPCION = "Descripción para revisión"
    QUE_DEBE_HACER = "Qué debe hacer"
    REQUIERE_SOPORTE = "Requiere soporte"
    LINK_EXTRACTO = "Link extracto"
    LINK_CARPETA_CREDITO = "Link carpeta crédito"
    CODIGO_TECNICO = "Código técnico"

    HEADERS = [
        ID_PAGO,
        CLIENTE,
        CREDITO,
        TIPO_CASO,
        DESCRIPCION,
        QUE_DEBE_HACER,
        REQUIERE_SOPORTE,
        LINK_EXTRACTO,
        LINK_CARPETA_CREDITO,
        CODIGO_TECNICO,
    ]
