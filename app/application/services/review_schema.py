"""
Fuente única de verdad para los nombres de hojas y columnas del Excel de revisión.

Generate escribe usando estas constantes.
Finalize lee usando estas mismas constantes.
Si necesitas cambiar un nombre de columna, cámbialo AQUÍ, no en cada archivo.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from enum import Enum
from typing import Any

class TipoAplicacion(str, Enum):
    """Tipo canónico del motor (PAGO / ABONO)."""

    PAGO = "PAGO"
    ABONO = "ABONO"


class TipoAplicacionVisible:
    PAGO = "PAGO"
    PAGO_Y_ABONO_CAPITAL = "PAGO Y ABONO CAPITAL"
    ABONO_CAPITAL = "ABONO CAPITAL"
    ABONO_MORA = "ABONO MORA"
    ABONO_LEGACY = "ABONO"


BANK_VISIBLE_APPLICATION_TYPES: tuple[str, ...] = (
    TipoAplicacionVisible.PAGO,
    TipoAplicacionVisible.PAGO_Y_ABONO_CAPITAL,
    TipoAplicacionVisible.ABONO_CAPITAL,
    TipoAplicacionVisible.ABONO_MORA,
)

APPLICATION_TYPES_SUPPORTED: tuple[str, ...] = BANK_VISIBLE_APPLICATION_TYPES


class CanonicalApplicationType:
    PAGO = "PAGO"
    ABONO = "ABONO"


class ApplicationSubtype:
    CUOTA = "CUOTA"
    CUOTA_MAS_CAPITAL = "CUOTA_MAS_CAPITAL"
    CAPITAL = "CAPITAL"
    MORA = "MORA"
    GENERAL = "GENERAL"


class ExtractRole:
    CIERRE_CUOTA = "CIERRE_CUOTA"
    REFERENCIA_MORA = "REFERENCIA_MORA"
    NO_APLICA = "NO_APLICA"


@dataclass(frozen=True)
class ApplicationPolicy:
    tipo_aplicacion_original: str
    tipo_aplicacion_canonica: str
    subtipo_aplicacion: str
    requiere_extracto: bool
    rol_extracto: str
    cierra_cuota: bool
    actualiza_ibr: bool
    genera_siguiente_extracto: bool
    legacy: bool = False

    @property
    def canonical_enum(self) -> TipoAplicacion:
        if self.tipo_aplicacion_canonica == TipoAplicacion.ABONO.value:
            return TipoAplicacion.ABONO
        return TipoAplicacion.PAGO

    def policy_dict(self) -> dict[str, Any]:
        return {
            "tipo_aplicacion_original": self.tipo_aplicacion_original,
            "tipo_aplicacion_canonica": self.tipo_aplicacion_canonica,
            "subtipo_aplicacion": self.subtipo_aplicacion,
            "requiere_extracto": self.requiere_extracto,
            "rol_extracto": self.rol_extracto,
            "cierra_cuota": self.cierra_cuota,
            "actualiza_ibr": self.actualiza_ibr,
            "genera_siguiente_extracto": self.genera_siguiente_extracto,
            "legacy": self.legacy,
        }


def _accent_fold_upper(text: str) -> str:
    folded = unicodedata.normalize("NFD", text)
    stripped = "".join(c for c in folded if unicodedata.category(c) != "Mn")
    return stripped.upper()


def _normalize_visible_tipo_text(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    collapsed = re.sub(r"\s+", " ", raw)
    return _accent_fold_upper(collapsed)


def resolve_application_policy(value: Any, *, from_bank: bool = False) -> ApplicationPolicy:
    """Resuelve política inmutable desde Tipo Aplicación visible u original."""
    text = _normalize_visible_tipo_text(value)
    if not text:
        raise ValueError("tipo_aplicacion_required")

    if text == _normalize_visible_tipo_text(TipoAplicacionVisible.PAGO):
        return ApplicationPolicy(
            tipo_aplicacion_original=TipoAplicacionVisible.PAGO,
            tipo_aplicacion_canonica=CanonicalApplicationType.PAGO,
            subtipo_aplicacion=ApplicationSubtype.CUOTA,
            requiere_extracto=True,
            rol_extracto=ExtractRole.CIERRE_CUOTA,
            cierra_cuota=True,
            actualiza_ibr=True,
            genera_siguiente_extracto=True,
        )
    if text == _normalize_visible_tipo_text(TipoAplicacionVisible.PAGO_Y_ABONO_CAPITAL):
        return ApplicationPolicy(
            tipo_aplicacion_original=TipoAplicacionVisible.PAGO_Y_ABONO_CAPITAL,
            tipo_aplicacion_canonica=CanonicalApplicationType.PAGO,
            subtipo_aplicacion=ApplicationSubtype.CUOTA_MAS_CAPITAL,
            requiere_extracto=True,
            rol_extracto=ExtractRole.CIERRE_CUOTA,
            cierra_cuota=True,
            actualiza_ibr=True,
            genera_siguiente_extracto=True,
        )
    if text == _normalize_visible_tipo_text(TipoAplicacionVisible.ABONO_CAPITAL):
        return ApplicationPolicy(
            tipo_aplicacion_original=TipoAplicacionVisible.ABONO_CAPITAL,
            tipo_aplicacion_canonica=CanonicalApplicationType.ABONO,
            subtipo_aplicacion=ApplicationSubtype.CAPITAL,
            requiere_extracto=False,
            rol_extracto=ExtractRole.NO_APLICA,
            cierra_cuota=False,
            actualiza_ibr=False,
            genera_siguiente_extracto=False,
        )
    if text == _normalize_visible_tipo_text(TipoAplicacionVisible.ABONO_MORA):
        return ApplicationPolicy(
            tipo_aplicacion_original=TipoAplicacionVisible.ABONO_MORA,
            tipo_aplicacion_canonica=CanonicalApplicationType.ABONO,
            subtipo_aplicacion=ApplicationSubtype.MORA,
            requiere_extracto=True,
            rol_extracto=ExtractRole.REFERENCIA_MORA,
            cierra_cuota=False,
            actualiza_ibr=False,
            genera_siguiente_extracto=False,
        )
    if text == _normalize_visible_tipo_text(TipoAplicacionVisible.ABONO_LEGACY):
        if from_bank:
            raise ValueError("generic_abono_not_supported")
        return ApplicationPolicy(
            tipo_aplicacion_original=TipoAplicacionVisible.ABONO_LEGACY,
            tipo_aplicacion_canonica=CanonicalApplicationType.ABONO,
            subtipo_aplicacion=ApplicationSubtype.GENERAL,
            requiere_extracto=False,
            rol_extracto=ExtractRole.NO_APLICA,
            cierra_cuota=False,
            actualiza_ibr=False,
            genera_siguiente_extracto=False,
            legacy=True,
        )

    if from_bank:
        raise ValueError("tipo_aplicacion_invalid")
    raise ValueError("tipo_aplicacion_invalid")


def parse_bank_tipo_aplicacion(value: Any) -> ApplicationPolicy:
    """Valida Tipo Aplicación del Excel bancario (solo valores visibles nuevos)."""
    return resolve_application_policy(value, from_bank=True)


def normalize_tipo_aplicacion(value: Any) -> TipoAplicacion:
    """Devuelve tipo canónico; acepta valores visibles nuevos y ABONO legacy en lecturas."""
    return resolve_application_policy(value, from_bank=False).canonical_enum


def requiere_extracto(tipo: TipoAplicacion | ApplicationPolicy | str | Any) -> bool:
    if isinstance(tipo, ApplicationPolicy):
        return tipo.requiere_extracto
    if isinstance(tipo, TipoAplicacion):
        return resolve_application_policy(tipo.value, from_bank=False).requiere_extracto
    try:
        return resolve_application_policy(tipo, from_bank=False).requiere_extracto
    except ValueError:
        return tipo == TipoAplicacion.PAGO or str(tipo or "").strip().upper() == TipoAplicacion.PAGO.value


def policy_requires_closing_extract(policy: ApplicationPolicy) -> bool:
    return policy.rol_extracto == ExtractRole.CIERRE_CUOTA and policy.requiere_extracto


def policy_requires_reference_extract(policy: ApplicationPolicy) -> bool:
    return policy.rol_extracto == ExtractRole.REFERENCIA_MORA and policy.requiere_extracto


def resolve_manifest_policy(
    source: dict[str, Any],
    *,
    default_canonical: str | None = None,
) -> ApplicationPolicy:
    """Resuelve política desde manifest/output/credit_item; legacy → PAGO o ABONO general."""
    original = (
        source.get("tipo_aplicacion_original")
        or source.get(DistribucionCols.TIPO_APLICACION_ORIGINAL)
        or source.get(DistribucionAbonosCols.TIPO_APLICACION_ORIGINAL)
        or source.get("tipo_aplicacion")
        or source.get(DistribucionAbonosCols.TIPO_APLICACION)
    )
    if original:
        try:
            return resolve_application_policy(original, from_bank=False)
        except ValueError:
            pass

    canon = str(source.get("tipo_aplicacion_canonica") or "").strip().upper()
    subtipo = str(source.get("subtipo_aplicacion") or "").strip().upper()
    if canon and subtipo:
        return ApplicationPolicy(
            tipo_aplicacion_original=str(
                source.get("tipo_aplicacion_original") or original or canon
            ),
            tipo_aplicacion_canonica=canon,
            subtipo_aplicacion=subtipo,
            requiere_extracto=bool(source.get("requiere_extracto")),
            rol_extracto=str(source.get("rol_extracto") or ""),
            cierra_cuota=bool(source.get("cierra_cuota")),
            actualiza_ibr=bool(source.get("actualiza_ibr")),
            genera_siguiente_extracto=bool(source.get("genera_siguiente_extracto")),
            legacy=bool(source.get("legacy") or source.get("policy_legacy")),
        )

    if canon == CanonicalApplicationType.ABONO:
        return resolve_application_policy(TipoAplicacionVisible.ABONO_LEGACY, from_bank=False)
    if canon == CanonicalApplicationType.PAGO:
        return resolve_application_policy(TipoAplicacionVisible.PAGO, from_bank=False)
    if default_canonical == TipoAplicacion.ABONO.value:
        return resolve_application_policy(TipoAplicacionVisible.ABONO_LEGACY, from_bank=False)
    return resolve_application_policy(TipoAplicacionVisible.PAGO, from_bank=False)


def policy_observability_dict(policy: ApplicationPolicy) -> dict[str, Any]:
    """Campos de política para items de dry-run/apply."""
    pd = policy.policy_dict()
    return {
        "tipo_aplicacion": policy.tipo_aplicacion_canonica,
        "tipo_aplicacion_original": pd["tipo_aplicacion_original"],
        "tipo_aplicacion_canonica": pd["tipo_aplicacion_canonica"],
        "subtipo_aplicacion": pd["subtipo_aplicacion"],
        "requiere_extracto": pd["requiere_extracto"],
        "requires_extract": pd["requiere_extracto"],
        "rol_extracto": pd["rol_extracto"],
        "cierra_cuota": pd["cierra_cuota"],
        "actualiza_ibr": pd["actualiza_ibr"],
        "genera_siguiente_extracto": pd["genera_siguiente_extracto"],
        "updates_ibr": pd["actualiza_ibr"],
        "policy_legacy": pd["legacy"],
    }


def policy_fields_for_manifest(policy: ApplicationPolicy) -> dict[str, Any]:
    """Subconjunto serializable para manifest (sin aliases de observabilidad)."""
    pd = policy.policy_dict()
    return {k: pd[k] for k in pd if k != "legacy"}


# Nombres de hojas
class ReviewSheets:
    CONTROL = "Control"
    RESUMEN = "Resumen"
    CASOS_PAGO = "Casos_Pago"
    DISTRIBUCION_PAGOS = "Distribucion_Pagos"
    DISTRIBUCION_LEGACY = "Distribucion"
    DISTRIBUCION_ABONOS = "Distribucion_Abonos"
    LISTAS = "_Listas"
    ERRORES = "Errores"
    # Alias legacy: lectores antiguos; Generate usa DISTRIBUCION_PAGOS.
    DISTRIBUCION = DISTRIBUCION_LEGACY
    DISTRIBUCION_PAGOS_FUTURE = DISTRIBUCION_PAGOS


def _norm_sheet_name(title: str) -> str:
    return _accent_fold_upper(str(title or "").strip().replace("_", " "))


def find_distribucion_pagos_sheet(wb: Any) -> Any:
    """Busca Distribucion_Pagos; si no existe, alias legacy Distribucion."""
    targets = (
        _norm_sheet_name(ReviewSheets.DISTRIBUCION_PAGOS),
        _norm_sheet_name(ReviewSheets.DISTRIBUCION_LEGACY),
    )
    for ws in wb.worksheets:
        if _norm_sheet_name(ws.title) in targets:
            return ws
    raise ValueError("missing_distribucion_pagos_sheet")


def workbook_has_distribucion_pagos_sheet(wb: Any) -> bool:
    try:
        find_distribucion_pagos_sheet(wb)
        return True
    except ValueError:
        return False


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


class ValidarAbono:
    SI = "SI"
    NO = "NO"


class EstadoPago:
    """Valores permitidos en columna «Estado Pago» (Distribución)."""

    ADELANTADO = "ADELANTADO"
    ATRASADO = "ATRASADO"
    NORMAL = "NORMAL"
    REVISION_MANUAL = "REVISION_MANUAL"

    ALLOWED = frozenset(
        {
            ADELANTADO,
            ATRASADO,
            NORMAL,
            REVISION_MANUAL,
        }
    )

    OPTIONS_ORDERED = [
        ADELANTADO,
        ATRASADO,
        NORMAL,
        REVISION_MANUAL,
    ]

    # Finalizar solo bloquea revisión manual pendiente (caso no listo para cierre operativo)
    FINALIZE_FORBIDDEN = frozenset({REVISION_MANUAL})

    # Suma aplicada / filas que consolidan en histórico y tabla de amortización
    # (incluye ATRASADO y ADELANTADO: adelanto sigue requiriendo seguimiento operativo)
    COUNTERS_POSITIVE_TOTAL = frozenset({NORMAL, ATRASADO, ADELANTADO})

    # Rutas extracto / soporte secretaría (antes «VALIDAR»)
    SECRETARY_AND_RUTA = frozenset({NORMAL, ATRASADO, ADELANTADO})

    # Estados con seguimiento operativo positivo (pagos_adelantados vía followup)
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
    # ABONO_A_CAPITAL: monto adicional aplicado directamente a capital.
    ABONO_A_CAPITAL = "Abono a capital"
    # MORA_A_APLICAR: alias legacy para libros antiguos cuyo encabezado visible era «Mora a aplicar».
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
    TIPO_APLICACION_ORIGINAL = "TipoAplicacionOriginal"
    TIPO_APLICACION_CANONICA = "TipoAplicacionCanonica"
    SUBTIPO_APLICACION = "SubtipoAplicacion"
    REQUIERE_EXTRACTO = "RequiereExtracto"
    ROL_EXTRACTO = "RolExtracto"
    CIERRA_CUOTA = "CierraCuota"
    ACTUALIZA_IBR = "ActualizaIBR"
    # Alias de código (mismo texto visible que las columnas renombradas)
    VALOR_INTERESES = APLICAR_A_EXTRACTO
    ABONO_K = ABONO_A_CAPITAL
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
        ABONO_A_CAPITAL,
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
        TIPO_APLICACION_ORIGINAL,
        TIPO_APLICACION_CANONICA,
        SUBTIPO_APLICACION,
        REQUIERE_EXTRACTO,
        ROL_EXTRACTO,
        CIERRA_CUOTA,
        ACTUALIZA_IBR,
    ]


# Columnas técnicas de Distribución que deben quedar ocultas en el Excel
DISTRIBUCION_TECHNICAL_HIDDEN_COLUMNS = frozenset(
    {
        DistribucionCols.RUTA_TABLA_AMORTIZACION,
        DistribucionCols.CREDITO_NORMALIZADO,
        DistribucionCols.TIPO_APLICACION_ORIGINAL,
        DistribucionCols.TIPO_APLICACION_CANONICA,
        DistribucionCols.SUBTIPO_APLICACION,
        DistribucionCols.REQUIERE_EXTRACTO,
        DistribucionCols.ROL_EXTRACTO,
        DistribucionCols.CIERRA_CUOTA,
        DistribucionCols.ACTUALIZA_IBR,
    }
)


class DistribucionAbonosCols:
    ID_PAGO = "ID Pago"
    CLIENTE = "Cliente"
    CREDITO = "Crédito"
    MONTO_BANCO = "Monto banco"
    FECHA_BANCO = "Fecha banco"
    VALIDAR_ABONO = "Validar Abono"
    OBSERVACION = "Observación"
    LINK_TABLA = "Link tabla amortización"
    LINK_CARPETA_CREDITO = "Link carpeta crédito"
    ORIGEN_CREDITO = "Origen crédito"
    RUTA_UNIDAD_CREDITO = "RutaUnidadCredito"
    RUTA_TABLA_AMORTIZACION = "RutaTablaAmortizacion"
    CREDITO_NORMALIZADO = "CreditoNormalizado"
    TIPO_APLICACION = "TipoAplicacion"
    TIPO_APLICACION_ORIGINAL = "TipoAplicacionOriginal"
    TIPO_APLICACION_CANONICA = "TipoAplicacionCanonica"
    SUBTIPO_APLICACION = "SubtipoAplicacion"
    REQUIERE_EXTRACTO = "RequiereExtracto"
    ROL_EXTRACTO = "RolExtracto"
    CIERRA_CUOTA = "CierraCuota"
    ACTUALIZA_IBR = "ActualizaIBR"
    LINK_EXTRACTO = "Link extracto"
    RUTA_EXTRACTO = "Ruta"
    FECHA_LIMITE = "Fecha límite"
    # Solo en cartera_validada (Finalize); no en validacion_pagos de Generate
    RUTA_ASIENTOS_CONTABLES = "RutaAsientosContables"

    HEADERS = [
        ID_PAGO,
        CLIENTE,
        CREDITO,
        MONTO_BANCO,
        FECHA_BANCO,
        VALIDAR_ABONO,
        OBSERVACION,
        LINK_TABLA,
        LINK_CARPETA_CREDITO,
        ORIGEN_CREDITO,
        LINK_EXTRACTO,
        FECHA_LIMITE,
        RUTA_UNIDAD_CREDITO,
        RUTA_TABLA_AMORTIZACION,
        CREDITO_NORMALIZADO,
        TIPO_APLICACION,
        TIPO_APLICACION_ORIGINAL,
        TIPO_APLICACION_CANONICA,
        SUBTIPO_APLICACION,
        REQUIERE_EXTRACTO,
        ROL_EXTRACTO,
        CIERRA_CUOTA,
        ACTUALIZA_IBR,
        RUTA_EXTRACTO,
    ]


SUPPORT_NOT_APPLICABLE = "NO APLICA"


class AsientosPendientesCols:
    """Hoja única de soporte secretaría (pagos y abonos)."""

    TIPO_APLICACION = "Tipo Aplicación"
    ID_PAGO = "ID Pago"
    BANCO = "Banco"
    CLIENTE = "Cliente"
    CREDITO = "Crédito"
    MONTO_BANCO = "Monto banco"
    FECHA_BANCO = "Fecha banco"
    FECHA_LIMITE = "Fecha límite"
    TOTAL_VALIDADO = "Total validado"
    LINK_CARPETA_ASIENTOS = "Link carpeta asientos contables"
    LINK_EXTRACTO = "Link extracto"
    LINK_TABLA = "Link tabla amortización"
    OBSERVACION = "Observación"

    HEADERS = [
        TIPO_APLICACION,
        ID_PAGO,
        BANCO,
        CLIENTE,
        CREDITO,
        MONTO_BANCO,
        FECHA_BANCO,
        FECHA_LIMITE,
        TOTAL_VALIDADO,
        LINK_CARPETA_ASIENTOS,
        LINK_EXTRACTO,
        LINK_TABLA,
        OBSERVACION,
    ]


DISTRIBUCION_ABONOS_TECHNICAL_HIDDEN_COLUMNS = frozenset(
    {
        DistribucionAbonosCols.RUTA_UNIDAD_CREDITO,
        DistribucionAbonosCols.RUTA_TABLA_AMORTIZACION,
        DistribucionAbonosCols.CREDITO_NORMALIZADO,
        DistribucionAbonosCols.TIPO_APLICACION,
        DistribucionAbonosCols.TIPO_APLICACION_ORIGINAL,
        DistribucionAbonosCols.TIPO_APLICACION_CANONICA,
        DistribucionAbonosCols.SUBTIPO_APLICACION,
        DistribucionAbonosCols.REQUIERE_EXTRACTO,
        DistribucionAbonosCols.ROL_EXTRACTO,
        DistribucionAbonosCols.CIERRA_CUOTA,
        DistribucionAbonosCols.ACTUALIZA_IBR,
        DistribucionAbonosCols.RUTA_EXTRACTO,
        DistribucionAbonosCols.RUTA_ASIENTOS_CONTABLES,
    }
)


def apply_policy_to_pagos_row(row: dict[str, Any], policy: ApplicationPolicy) -> None:
    """Escribe columnas técnicas de política en fila Distribucion_Pagos."""
    pd = policy.policy_dict()
    row[DistribucionCols.TIPO_APLICACION_ORIGINAL] = pd["tipo_aplicacion_original"]
    row[DistribucionCols.TIPO_APLICACION_CANONICA] = pd["tipo_aplicacion_canonica"]
    row[DistribucionCols.SUBTIPO_APLICACION] = pd["subtipo_aplicacion"]
    row[DistribucionCols.REQUIERE_EXTRACTO] = pd["requiere_extracto"]
    row[DistribucionCols.ROL_EXTRACTO] = pd["rol_extracto"]
    row[DistribucionCols.CIERRA_CUOTA] = pd["cierra_cuota"]
    row[DistribucionCols.ACTUALIZA_IBR] = pd["actualiza_ibr"]


def apply_policy_to_abonos_row(row: dict[str, Any], policy: ApplicationPolicy) -> None:
    """Escribe columnas técnicas de política en fila Distribucion_Abonos."""
    pd = policy.policy_dict()
    row[DistribucionAbonosCols.TIPO_APLICACION] = policy.tipo_aplicacion_canonica
    row[DistribucionAbonosCols.TIPO_APLICACION_ORIGINAL] = pd["tipo_aplicacion_original"]
    row[DistribucionAbonosCols.TIPO_APLICACION_CANONICA] = pd["tipo_aplicacion_canonica"]
    row[DistribucionAbonosCols.SUBTIPO_APLICACION] = pd["subtipo_aplicacion"]
    row[DistribucionAbonosCols.REQUIERE_EXTRACTO] = pd["requiere_extracto"]
    row[DistribucionAbonosCols.ROL_EXTRACTO] = pd["rol_extracto"]
    row[DistribucionAbonosCols.CIERRA_CUOTA] = pd["cierra_cuota"]
    row[DistribucionAbonosCols.ACTUALIZA_IBR] = pd["actualiza_ibr"]


def policy_from_row(row: dict[str, Any]) -> ApplicationPolicy:
    """Reconstruye política desde columnas técnicas o TipoAplicacion legacy en fila."""
    original = (
        row.get(DistribucionCols.TIPO_APLICACION_ORIGINAL)
        or row.get(DistribucionAbonosCols.TIPO_APLICACION_ORIGINAL)
        or row.get(DistribucionAbonosCols.TIPO_APLICACION)
        or row.get(AsientosPendientesCols.TIPO_APLICACION)
    )
    if original:
        return resolve_application_policy(original, from_bank=False)
    canon = row.get(DistribucionAbonosCols.TIPO_APLICACION_CANONICA) or row.get(
        DistribucionCols.TIPO_APLICACION_CANONICA
    )
    if canon == CanonicalApplicationType.PAGO:
        return resolve_application_policy(TipoAplicacionVisible.PAGO, from_bank=False)
    if canon == CanonicalApplicationType.ABONO:
        return resolve_application_policy(TipoAplicacionVisible.ABONO_LEGACY, from_bank=False)
    raise ValueError("tipo_aplicacion_required")


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
    "Abono a K": DistribucionCols.ABONO_A_CAPITAL,
    DistribucionCols.MORA_A_APLICAR: DistribucionCols.ABONO_A_CAPITAL,
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
    REPROGRAMAR = "REPROGRAMAR"
    NO_VALIDAR = "NO_VALIDAR"
    PENDIENTE_MORA = "PENDIENTE_MORA"
    REVISION_MANUAL = "REVISION_MANUAL"
    OPTIONS = [
        VALIDAR,
        PENDIENTE_MORA,
        REPROGRAMAR,
        NO_VALIDAR,
        REVISION_MANUAL,
    ]


_LEGACY_ESTADO_TO_PAIR: dict[str, tuple[str, str]] = {
    EstadoLinea.VALIDAR: (EstadoPago.NORMAL, ValidarPago.SI),
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


def normalize_validar_abono_value(raw: Any) -> str:
    """Devuelve ``SI``, ``NO`` o cadena vacía si no se reconoce."""
    return _normalize_validar_pago_raw(raw)


def is_validar_abono_si(row: dict[str, Any]) -> bool:
    return (
        normalize_validar_abono_value(row.get(DistribucionAbonosCols.VALIDAR_ABONO))
        == ValidarAbono.SI
    )


def require_validar_abono_value(raw: Any) -> str:
    """Normaliza Validar Abono; vacío → NO; valor no reconocido → error."""
    if raw is None or str(raw).strip() == "":
        return ValidarAbono.NO
    norm = normalize_validar_abono_value(raw)
    if not norm:
        raise ValueError("invalid_validar_abono")
    return norm


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
