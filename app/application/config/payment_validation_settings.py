"""
Configuración central del flujo de validación de pagos (rutas SharePoint y bancos).

Sin lógica financiera. Prioridad: variables PAYMENT_* canónicas → alias GRAPH_* → defaults HBI.
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from enum import Enum

logger = logging.getLogger(__name__)

# --- Códigos y nombres por defecto (única fuente en código) ---
BANK_CODE_BOGOTA = "banco_bogota"
BANK_CODE_BANCOLOMBIA = "banco_bancolombia"

DEFAULT_BANK_NAME_BOGOTA = "Banco de Bogotá"
DEFAULT_BANK_NAME_BANCOLOMBIA = "Bancolombia"

# Etiqueta visible en correo (subject, cuerpo {banco}, PDF exportado del mail).
DEFAULT_BANK_EMAIL_LABEL_BOGOTA = "BANCO BOGOTA"
DEFAULT_BANK_EMAIL_LABEL_BANCOLOMBIA = "BANCO BANCOLOMBIA"

DEFAULT_INPUT_FILENAME_BOGOTA = "BANCO_BOGOTA.xlsx"
DEFAULT_INPUT_FILENAME_BANCOLOMBIA = "BANCO_BANCOLOMBIA.xlsx"

DEFAULT_CONTROL_FILENAME_BOGOTA = "control_proceso_validacion_pagos_banco_bogota.xlsx"
DEFAULT_CONTROL_FILENAME_BANCOLOMBIA = "control_proceso_validacion_pagos_banco_bancolombia.xlsx"

DEFAULT_EMAIL_SUBJECT_BOGOTA = "ABONOS BANCO BOGOTA"
DEFAULT_EMAIL_SUBJECT_BANCOLOMBIA = "ABONOS BANCO BANCOLOMBIA"

DEFAULT_EMAIL_PDF_TEMPLATE_BOGOTA = "ABONOS BANCO BOGOTA {fecha}.pdf"
DEFAULT_EMAIL_PDF_TEMPLATE_BANCOLOMBIA = "ABONOS BANCO BANCOLOMBIA {fecha}.pdf"

DEFAULT_CORREOS_FILENAME = "CORREOS.xlsx"
DEFAULT_IBR_FILENAME = "IBR_DIARIO.xlsx"
DEFAULT_FOLLOWUP_ADELANTADOS = "pagos_adelantados.xlsx"
DEFAULT_FOLLOWUP_INCOMPLETOS = "pagos_incompletos.xlsx"

_DEFAULT_BASE_FOLDER = "INFORMACION CREDITOS-CLIENTES/02 COMWARE - VALIDACION PAGOS"
_DEFAULT_SUBFOLDERS: dict[str, str] = {
    "control": "00 CONTROL",
    "review": "01 REVISION",
    "historical": "02 HISTORICO",
    "errors": "03 ERRORES",
    "logs": "04 LOGS",
    "email": "05 EMAIL",
    "asientos": "06 ASIENTO CONTABLES GENERADOS",
}

# Alias legacy GRAPH_* → carpeta lógica
_GRAPH_LEGACY_FOLDER_KEYS: dict[str, str] = {
    "control": "GRAPH_PAYMENT_VALIDATION_CONTROL_PATH",
    "review": "GRAPH_PAYMENT_VALIDATION_REVIEW_PATH",
    "historical": "GRAPH_PAYMENT_VALIDATION_HISTORY_PATH",
    "errors": "GRAPH_PAYMENT_VALIDATION_ERRORS_PATH",
    "logs": "GRAPH_PAYMENT_VALIDATION_LOGS_PATH",
}

_PAYMENT_CANONICAL_FOLDER_KEYS: dict[str, str] = {
    "control": "PAYMENT_VALIDATION_CONTROL_FOLDER",
    "review": "PAYMENT_VALIDATION_REVIEW_FOLDER",
    "historical": "PAYMENT_VALIDATION_HISTORICAL_FOLDER",
    "errors": "PAYMENT_VALIDATION_ERRORS_FOLDER",
    "logs": "PAYMENT_VALIDATION_LOGS_FOLDER",
    "email": "PAYMENT_VALIDATION_EMAIL_FOLDER",
    "asientos": "PAYMENT_VALIDATION_ASIENTOS_FOLDER",
}


class PaymentValidationFolderName(str, Enum):
    CONTROL = "control"
    REVIEW = "review"
    HISTORICAL = "historical"
    ERRORS = "errors"
    LOGS = "logs"
    EMAIL = "email"
    ASIENTOS = "asientos"


@dataclass(frozen=True)
class PaymentValidationPaths:
    base_folder: str
    control: str
    review: str
    historical: str
    errors: str
    logs: str
    email: str
    asientos: str


@dataclass(frozen=True)
class PaymentBankConfig:
    bank_code: str
    bank_name: str
    bank_email_label: str
    input_file_path: str
    control_file_path: str
    email_subject_default: str
    email_pdf_name_template_default: str


def _strip_env(key: str) -> str:
    return os.getenv(key, "").strip()


def _join_relative(*parts: str) -> str:
    cleaned: list[str] = []
    for part in parts:
        for segment in part.replace("\\", "/").split("/"):
            seg = segment.strip()
            if seg:
                cleaned.append(seg)
    return "/".join(cleaned)


def _resolve_base_folder() -> str:
    base = _strip_env("PAYMENT_VALIDATION_BASE_FOLDER")
    if base:
        return base.rstrip("/")
    legacy_base = _strip_env("GRAPH_PAYMENT_VALIDATION_BASE_PATH")
    if legacy_base:
        return legacy_base.rstrip("/")
    return _DEFAULT_BASE_FOLDER


def _resolve_folder_path(folder_key: str) -> str:
    legacy_key = _GRAPH_LEGACY_FOLDER_KEYS.get(folder_key)
    if legacy_key:
        legacy = _strip_env(legacy_key)
        if legacy:
            return legacy.rstrip("/")

    canonical_key = _PAYMENT_CANONICAL_FOLDER_KEYS.get(folder_key)
    if canonical_key:
        sub = _strip_env(canonical_key)
        if sub:
            return _join_relative(_resolve_base_folder(), sub)

    default_sub = _DEFAULT_SUBFOLDERS.get(folder_key, "")
    if default_sub:
        return _join_relative(_resolve_base_folder(), default_sub)
    return _resolve_base_folder()


def resolve_payment_validation_folder(name: PaymentValidationFolderName | str) -> str:
    key = name.value if isinstance(name, PaymentValidationFolderName) else str(name).strip().lower()
    if key not in _DEFAULT_SUBFOLDERS:
        raise ValueError(f"invalid_payment_validation_folder:{key}")
    return _resolve_folder_path(key)


def get_payment_validation_paths() -> PaymentValidationPaths:
    base = _resolve_base_folder()
    return PaymentValidationPaths(
        base_folder=base,
        control=resolve_payment_validation_folder(PaymentValidationFolderName.CONTROL),
        review=resolve_payment_validation_folder(PaymentValidationFolderName.REVIEW),
        historical=resolve_payment_validation_folder(PaymentValidationFolderName.HISTORICAL),
        errors=resolve_payment_validation_folder(PaymentValidationFolderName.ERRORS),
        logs=resolve_payment_validation_folder(PaymentValidationFolderName.LOGS),
        email=resolve_payment_validation_folder(PaymentValidationFolderName.EMAIL),
        asientos=resolve_payment_validation_folder(PaymentValidationFolderName.ASIENTOS),
    )


def normalize_bank_code(bank_code: str | None) -> str:
    bc = (bank_code or "").strip()
    return bc or BANK_CODE_BOGOTA


def validate_bank_code(bank_code: str) -> None:
    if bank_code not in (BANK_CODE_BOGOTA, BANK_CODE_BANCOLOMBIA):
        raise ValueError("invalid_bank_code")


def _heuristic_bancolombia_from_bogota_path(bogota_path: str) -> str:
    if bogota_path.endswith("/BANCO_BOGOTA.xlsx"):
        return bogota_path.rsplit("/", 1)[0] + "/" + DEFAULT_INPUT_FILENAME_BANCOLOMBIA
    if bogota_path.endswith("BANCO_BOGOTA.xlsx"):
        return bogota_path.replace(DEFAULT_INPUT_FILENAME_BOGOTA, DEFAULT_INPUT_FILENAME_BANCOLOMBIA)
    return ""


def _resolve_input_path_for_bogota() -> str:
    canonical = _strip_env("PAYMENT_BANK_BOGOTA_INPUT_FILE_PATH")
    if canonical:
        return canonical
    for key in ("GRAPH_BANK_PAYMENTS_FILE_PATH", "GRAPH_SHAREPOINT_FILE_PATH"):
        val = _strip_env(key)
        if val:
            return val
    control = resolve_payment_validation_folder(PaymentValidationFolderName.CONTROL)
    return _join_relative(control, DEFAULT_INPUT_FILENAME_BOGOTA)


def _resolve_input_path_for_bancolombia() -> str:
    canonical = _strip_env("PAYMENT_BANK_BANCOLOMBIA_INPUT_FILE_PATH")
    if canonical:
        return canonical
    for key in ("GRAPH_BANK_PAYMENTS_FILE_PATH_BANCOLOMBIA", "GRAPH_SHAREPOINT_FILE_PATH_BANCOLOMBIA"):
        val = _strip_env(key)
        if val:
            return val
    bogota = _resolve_input_path_for_bogota()
    guessed = _heuristic_bancolombia_from_bogota_path(bogota)
    if guessed:
        logger.warning(
            "payment_validation_settings: Bancolombia input path inferred from Bogotá path "
            "(deprecated). Set PAYMENT_BANK_BANCOLOMBIA_INPUT_FILE_PATH or "
            "GRAPH_BANK_PAYMENTS_FILE_PATH_BANCOLOMBIA."
        )
        return guessed
    control = resolve_payment_validation_folder(PaymentValidationFolderName.CONTROL)
    return _join_relative(control, DEFAULT_INPUT_FILENAME_BANCOLOMBIA)


def _resolve_control_path(bank_code: str) -> str:
    control_dir = resolve_payment_validation_folder(PaymentValidationFolderName.CONTROL)
    if bank_code == BANK_CODE_BOGOTA:
        filename = (
            _strip_env("PAYMENT_BANK_BOGOTA_CONTROL_FILE")
            or DEFAULT_CONTROL_FILENAME_BOGOTA
        )
    elif bank_code == BANK_CODE_BANCOLOMBIA:
        filename = (
            _strip_env("PAYMENT_BANK_BANCOLOMBIA_CONTROL_FILE")
            or DEFAULT_CONTROL_FILENAME_BANCOLOMBIA
        )
    else:
        raise ValueError("invalid_bank_code")
    if "/" in filename:
        return filename.strip().strip("/")
    return _join_relative(control_dir, filename)


def _build_bank_config(bank_code: str) -> PaymentBankConfig:
    if bank_code == BANK_CODE_BOGOTA:
        code = _strip_env("PAYMENT_BANK_BOGOTA_CODE") or BANK_CODE_BOGOTA
        name = _strip_env("PAYMENT_BANK_BOGOTA_NAME") or DEFAULT_BANK_NAME_BOGOTA
        email_label = _strip_env("PAYMENT_BANK_BOGOTA_EMAIL_LABEL") or DEFAULT_BANK_EMAIL_LABEL_BOGOTA
        input_path = _resolve_input_path_for_bogota()
        subject = _strip_env("PAYMENT_BANK_BOGOTA_EMAIL_SUBJECT") or DEFAULT_EMAIL_SUBJECT_BOGOTA
        pdf_tpl = (
            _strip_env("PAYMENT_BANK_BOGOTA_EMAIL_PDF_NAME_TEMPLATE")
            or DEFAULT_EMAIL_PDF_TEMPLATE_BOGOTA
        )
    elif bank_code == BANK_CODE_BANCOLOMBIA:
        code = _strip_env("PAYMENT_BANK_BANCOLOMBIA_CODE") or BANK_CODE_BANCOLOMBIA
        name = _strip_env("PAYMENT_BANK_BANCOLOMBIA_NAME") or DEFAULT_BANK_NAME_BANCOLOMBIA
        email_label = _strip_env("PAYMENT_BANK_BANCOLOMBIA_EMAIL_LABEL") or DEFAULT_BANK_EMAIL_LABEL_BANCOLOMBIA
        input_path = _resolve_input_path_for_bancolombia()
        subject = _strip_env("PAYMENT_BANK_BANCOLOMBIA_EMAIL_SUBJECT") or DEFAULT_EMAIL_SUBJECT_BANCOLOMBIA
        pdf_tpl = (
            _strip_env("PAYMENT_BANK_BANCOLOMBIA_EMAIL_PDF_NAME_TEMPLATE")
            or DEFAULT_EMAIL_PDF_TEMPLATE_BANCOLOMBIA
        )
    else:
        raise ValueError("invalid_bank_code")

    return PaymentBankConfig(
        bank_code=code,
        bank_name=name,
        bank_email_label=email_label,
        input_file_path=input_path,
        control_file_path=_resolve_control_path(bank_code),
        email_subject_default=subject,
        email_pdf_name_template_default=pdf_tpl,
    )


def get_payment_bank_config(bank_code: str) -> PaymentBankConfig:
    validate_bank_code(normalize_bank_code(bank_code))
    bc = normalize_bank_code(bank_code)
    return _build_bank_config(bc)


def list_payment_banks() -> tuple[PaymentBankConfig, ...]:
    return (_build_bank_config(BANK_CODE_BOGOTA), _build_bank_config(BANK_CODE_BANCOLOMBIA))


def resolve_bank_input_file_path(bank_code: str) -> str:
    return get_payment_bank_config(bank_code).input_file_path


def resolve_bank_control_file_path(bank_code: str) -> str:
    return get_payment_bank_config(bank_code).control_file_path


def resolve_bank_display_name(bank_code: str) -> str:
    """Nombre legible en control Excel (p. ej. Banco de Bogotá), no etiqueta de correo."""
    return get_payment_bank_config(bank_code).bank_name


def resolve_bank_email_label(bank_code: str) -> str:
    """Etiqueta visible en correo: subject, cuerpo {banco}, PDF del mail."""
    return get_payment_bank_config(bank_code).bank_email_label


def _warn_deprecated_global_notify_env(var_name: str, bank_code: str) -> None:
    if _strip_env(var_name):
        logger.warning(
            "payment_validation_settings: %s is set but ignored for multi-bank notify; "
            "use PAYMENT_BANK_%s_EMAIL_* per bank instead (requested bank_code=%s).",
            var_name,
            "BOGOTA" if normalize_bank_code(bank_code) == BANK_CODE_BOGOTA else "BANCOLOMBIA",
            normalize_bank_code(bank_code),
        )


def resolve_email_subject(bank_code: str) -> str:
    """Asunto del correo por banco (prioridad PAYMENT_BANK_*_EMAIL_SUBJECT)."""
    _warn_deprecated_global_notify_env("GRAPH_VALIDAR_NOTIFY_EMAIL_SUBJECT", bank_code)
    return get_payment_bank_config(bank_code).email_subject_default


def resolve_email_pdf_name_template(bank_code: str) -> str:
    """Plantilla de nombre del PDF exportado del correo por banco."""
    _warn_deprecated_global_notify_env(
        "GRAPH_VALIDAR_NOTIFY_EXPORT_EMAIL_PDF_NAME_TEMPLATE",
        bank_code,
    )
    return get_payment_bank_config(bank_code).email_pdf_name_template_default


def resolve_correos_xlsx_path() -> str:
    rel = _strip_env("GRAPH_VALIDAR_NOTIFY_CORREOS_XLSX_PATH")
    if rel:
        return rel.strip("/")
    control = resolve_payment_validation_folder(PaymentValidationFolderName.CONTROL)
    return _join_relative(control, DEFAULT_CORREOS_FILENAME)


def resolve_email_export_folder_path() -> str:
    rel = _strip_env("GRAPH_VALIDAR_NOTIFY_EXPORT_EMAIL_PDF_FOLDER_PATH")
    if rel:
        return rel.strip("/")
    return resolve_payment_validation_folder(PaymentValidationFolderName.EMAIL)


def resolve_merge_output_folder_path() -> str:
    rel = _strip_env("GRAPH_MERGE_COMPOSITE_OUTPUT_FOLDER_PATH")
    if rel:
        return rel.strip("/")
    return resolve_payment_validation_folder(PaymentValidationFolderName.ASIENTOS)


def resolve_logs_folder_path() -> str:
    return resolve_payment_validation_folder(PaymentValidationFolderName.LOGS)


def resolve_followup_workbook_path(filename: str) -> str:
    name = filename.strip()
    if name == DEFAULT_FOLLOWUP_ADELANTADOS:
        for key in ("GRAPH_FOLLOWUP_PAGOS_ADELANTADOS_PATH", "GRAPH_AUDIT_PAGOS_ADELANTADOS_PATH"):
            p = _strip_env(key)
            if p:
                return p
    elif name == DEFAULT_FOLLOWUP_INCOMPLETOS:
        for key in ("GRAPH_FOLLOWUP_PAGOS_INCOMPLETOS_PATH", "GRAPH_AUDIT_PAGOS_INCOMPLETOS_PATH"):
            p = _strip_env(key)
            if p:
                return p
    control = resolve_payment_validation_folder(PaymentValidationFolderName.CONTROL)
    return _join_relative(control, name)


def resolve_ibr_workbook_path() -> str:
    p = _strip_env("GRAPH_IBR_DIARIO_PATH")
    if p:
        return p
    control = resolve_payment_validation_folder(PaymentValidationFolderName.CONTROL)
    return _join_relative(control, DEFAULT_IBR_FILENAME)


def resolve_bank_report_path(bank_code: str) -> str:
    """Ruta del Excel de reporte del banco (notify / merge / generate)."""
    return resolve_bank_input_file_path(bank_code)
