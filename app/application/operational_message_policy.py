"""
Clasificación interna de mensajes operativos (no expuesta en contrato HTTP).

SECRETARY_CAN_CORRECT — la secretaría puede corregir archivos/documentos y reintentar.
RETRY_ONLY — reintentar el paso sin cambios en archivos.
SUPPORT_REQUIRED — inconveniente técnico; no pedir parámetros ni detalles técnicos.
CORRECT_THEN_SUPPORT — corrección operativa posible; escalar si persiste.
"""

from __future__ import annotations

from typing import Literal

MessageAudience = Literal[
    "SECRETARY_CAN_CORRECT",
    "RETRY_ONLY",
    "SUPPORT_REQUIRED",
    "CORRECT_THEN_SUPPORT",
]

SUPPORT_USER_DEFAULT = (
    "No fue posible completar el proceso debido a un inconveniente técnico."
)
SUPPORT_NEXT_DEFAULT = (
    "No continúe con el siguiente paso. Contacte a soporte e indique el banco, "
    "la fecha y la etapa en la que se produjo el error."
)

CORRECT_THEN_SUPPORT_SUFFIX = " Si el error continúa, contacte a soporte."

# Clasificación por error_code estable (raises y enrichment).
AUDIENCE_BY_CODE: dict[str, MessageAudience] = {
    # Generate — operativos
    "review_folder_not_empty": "SECRETARY_CAN_CORRECT",
    "bank_headers_not_found": "SECRETARY_CAN_CORRECT",
    "tipo_aplicacion_column_missing": "SECRETARY_CAN_CORRECT",
    "tipo_aplicacion_column_duplicate": "SECRETARY_CAN_CORRECT",
    "tipo_aplicacion_required": "SECRETARY_CAN_CORRECT",
    "tipo_aplicacion_invalid": "SECRETARY_CAN_CORRECT",
    "generic_abono_not_supported": "SECRETARY_CAN_CORRECT",
    "customer_not_found": "SECRETARY_CAN_CORRECT",
    "customer_ambiguous": "SECRETARY_CAN_CORRECT",
    "credit_folder_not_found": "SECRETARY_CAN_CORRECT",
    "extract_not_found": "SECRETARY_CAN_CORRECT",
    "extract_amount_not_found": "SECRETARY_CAN_CORRECT",
    "pending_installment_not_found": "SECRETARY_CAN_CORRECT",
    "amortization_table_not_found": "SECRETARY_CAN_CORRECT",
    "amortization_sheet_not_found": "SECRETARY_CAN_CORRECT",
    "amortization_table_ambiguous": "CORRECT_THEN_SUPPORT",
    # Generate — soporte / reintento
    "invalid_bank_code": "SUPPORT_REQUIRED",
    "missing_sharepoint_folder": "SUPPORT_REQUIRED",
    "active_process_exists": "RETRY_ONLY",
    # Finalize — operativos
    "process_not_approved": "SECRETARY_CAN_CORRECT",
    "missing_control_state": "SECRETARY_CAN_CORRECT",
    "invalid_control_state": "RETRY_ONLY",
    "empty_estado_pago": "SECRETARY_CAN_CORRECT",
    "invalid_estado_pago": "SECRETARY_CAN_CORRECT",
    "INCOMPLETO_NOT_SUPPORTED": "SECRETARY_CAN_CORRECT",
    "estado_pago_no_finalizable": "SECRETARY_CAN_CORRECT",
    "no_validar_requires_observation": "SECRETARY_CAN_CORRECT",
    "missing_valor_intereses": "SECRETARY_CAN_CORRECT",
    "missing_abono_k": "SECRETARY_CAN_CORRECT",
    "missing_abono_capital": "SECRETARY_CAN_CORRECT",
    "missing_mora_a_aplicar": "SECRETARY_CAN_CORRECT",
    "missing_otros_valores": "SECRETARY_CAN_CORRECT",
    "review_schema_version_1_requires_regenerate": "SECRETARY_CAN_CORRECT",
    "missing_mora": "SECRETARY_CAN_CORRECT",
    "duplicate_bank_amount_in_payment_group": "SECRETARY_CAN_CORRECT",
    "amount_mismatch": "SECRETARY_CAN_CORRECT",
    "missing_extract_route": "SECRETARY_CAN_CORRECT",
    "missing_ruta_unidad_credito": "SECRETARY_CAN_CORRECT",
    "credit_number_not_resolved": "CORRECT_THEN_SUPPORT",
    "asientos_folder_create_failed": "CORRECT_THEN_SUPPORT",
    "missing_ruta_asientos_contables": "SECRETARY_CAN_CORRECT",
    "no_validation_file_found": "SECRETARY_CAN_CORRECT",
    "upload_failed": "CORRECT_THEN_SUPPORT",
    "invalid_validar_abono": "SECRETARY_CAN_CORRECT",
    "abono_without_selected_credit": "SECRETARY_CAN_CORRECT",
    "abono_duplicate_selected_credit": "SECRETARY_CAN_CORRECT",
    "abono_group_inconsistent": "SECRETARY_CAN_CORRECT",
    "abono_credit_without_unit_path": "SECRETARY_CAN_CORRECT",
    "abono_credit_without_amortization_path": "SECRETARY_CAN_CORRECT",
    "abono_missing_bank_amount": "SECRETARY_CAN_CORRECT",
    "abono_missing_bank_date": "SECRETARY_CAN_CORRECT",
    "abono_invalid_application_type": "SECRETARY_CAN_CORRECT",
    "abono_mora_missing_reference_extract": "SECRETARY_CAN_CORRECT",
    "abono_mora_missing_reference_date": "SECRETARY_CAN_CORRECT",
    "missing_reference_extract_route": "SECRETARY_CAN_CORRECT",
    "pago_y_abono_capital_missing_parte_cuota": "SECRETARY_CAN_CORRECT",
    "pago_y_abono_capital_missing_capital": "SECRETARY_CAN_CORRECT",
    "pago_y_abono_capital_saldo_must_be_zero": "SECRETARY_CAN_CORRECT",
    "missing_control_sheet": "SECRETARY_CAN_CORRECT",
    "missing_distribucion_sheet": "SECRETARY_CAN_CORRECT",
    "missing_sheet_headers": "SECRETARY_CAN_CORRECT",
    "credit_folder_ambiguous": "SECRETARY_CAN_CORRECT",
    "validar_requires_positive_total": "SECRETARY_CAN_CORRECT",
    "NO_READY_PROCESS": "SECRETARY_CAN_CORRECT",
    "MULTIPLE_READY_PROCESSES": "SUPPORT_REQUIRED",
    "control_not_ready_for_finalize": "SECRETARY_CAN_CORRECT",
    "missing_validation_file_path": "SECRETARY_CAN_CORRECT",
    # Unmapped / dry-run / merge gates
    "bank_amount_parse_error": "SECRETARY_CAN_CORRECT",
    "bank_date_parse_error": "SECRETARY_CAN_CORRECT",
    "control_not_ready_for_dry_run": "SECRETARY_CAN_CORRECT",
    "control_not_ready_for_merge": "SECRETARY_CAN_CORRECT",
    "missing_distribucion_abonos_headers": "SECRETARY_CAN_CORRECT",
    "missing_distribucion_pagos_sheet": "SECRETARY_CAN_CORRECT",
    "pdf_no_text": "CORRECT_THEN_SUPPORT",
    "preflight_errors": "SECRETARY_CAN_CORRECT",
    "preflight_revision_manual": "SECRETARY_CAN_CORRECT",
    "preflight_warnings_not_allowed": "SECRETARY_CAN_CORRECT",
    "bank_code_and_process_date_required": "SUPPORT_REQUIRED",
    "destination_name_exhausted": "SECRETARY_CAN_CORRECT",
    "missing_merge_manifest_path": "SUPPORT_REQUIRED",
    "process_control_invalid_structure": "SUPPORT_REQUIRED",
    # Notify / merge heurísticos
    "control_not_ready_for_notify": "SECRETARY_CAN_CORRECT",
    "already_notified": "RETRY_ONLY",
    "missing_historical_file_path": "CORRECT_THEN_SUPPORT",
    "historical_file_not_found": "SECRETARY_CAN_CORRECT",
    "missing_distribucion_headers": "SECRETARY_CAN_CORRECT",
    "missing_distribucion_status_column": "SECRETARY_CAN_CORRECT",
    "missing_distribucion_route_column": "SECRETARY_CAN_CORRECT",
    "missing_ruta_column": "SECRETARY_CAN_CORRECT",
    "no_validated_rows": "SECRETARY_CAN_CORRECT",
    "recipients_not_configured": "SECRETARY_CAN_CORRECT",
    "bank_report_missing_date_column": "SECRETARY_CAN_CORRECT",
    "bank_report_no_valid_dates": "SECRETARY_CAN_CORRECT",
    "bank_report_invalid_data": "SECRETARY_CAN_CORRECT",
    "notify_config_missing": "SUPPORT_REQUIRED",
    "extract_pdf_not_found": "SECRETARY_CAN_CORRECT",
    "graph_sendmail_failed": "CORRECT_THEN_SUPPORT",
    "merge_control_no_pending_process": "SECRETARY_CAN_CORRECT",
    "missing_email_pdf_path": "SECRETARY_CAN_CORRECT",
    "merge_control_workbook_not_found": "SUPPORT_REQUIRED",
    "merge_control_invalid_structure": "SUPPORT_REQUIRED",
    "email_pdf_not_found": "SECRETARY_CAN_CORRECT",
    "missing_asiento_contable_pdf": "SECRETARY_CAN_CORRECT",
    "extract_routes_missing": "SECRETARY_CAN_CORRECT",
    "no_validated_rows_merge": "SECRETARY_CAN_CORRECT",
    "asiento_contable_not_found": "SECRETARY_CAN_CORRECT",
    "asiento_contable_credit_mismatch": "SECRETARY_CAN_CORRECT",
    "asiento_contable_ambiguous": "SECRETARY_CAN_CORRECT",
    "extract_pdf_download_failed": "CORRECT_THEN_SUPPORT",
    "asiento_pdf_download_failed": "CORRECT_THEN_SUPPORT",
    "consolidated_upload_failed": "CORRECT_THEN_SUPPORT",
    "merge_control_manifest_path_missing": "SUPPORT_REQUIRED",
    "merge_control_amortization_not_ready": "SECRETARY_CAN_CORRECT",
    "report_date_iso_required": "SUPPORT_REQUIRED",
    "merge_manifest_not_found": "CORRECT_THEN_SUPPORT",
    "invalid_merge_manifest_json": "SUPPORT_REQUIRED",
    "graph_config_error": "SUPPORT_REQUIRED",
    "MERGE_INCOMPLETE_NOT_APPLICABLE": "SECRETARY_CAN_CORRECT",
    "APPLICATION_PAYMENT_SECTION_FULL": "CORRECT_THEN_SUPPORT",
    "POST_UPLOAD_VERIFICATION_FAILED": "SUPPORT_REQUIRED",
    "APPLY_SAFETY_ABORT": "SUPPORT_REQUIRED",
    "APPLY_EXPECTED_EVENTS_INCOMPLETE": "SUPPORT_REQUIRED",
    "unknown_error": "SUPPORT_REQUIRED",
}


def get_message_audience(code: str) -> MessageAudience:
    base = (code or "").strip().split("|", 1)[0]
    return AUDIENCE_BY_CODE.get(base, "CORRECT_THEN_SUPPORT")


def apply_audience_policy(
    code: str,
    user_message: str,
    next_action: str,
) -> tuple[str, str]:
    """Ajusta next_action según audiencia sin alterar error_code ni technical_message."""
    audience = get_message_audience(code)
    um = (user_message or "").strip()
    na = (next_action or "").strip()

    if audience == "SUPPORT_REQUIRED":
        if not um or um == SUPPORT_USER_DEFAULT:
            um = SUPPORT_USER_DEFAULT
        na = SUPPORT_NEXT_DEFAULT
    elif audience == "CORRECT_THEN_SUPPORT" and na and CORRECT_THEN_SUPPORT_SUFFIX not in na:
        if "contacte a soporte" not in na.lower():
            na = na.rstrip(".") + "." + CORRECT_THEN_SUPPORT_SUFFIX

    return um, na
