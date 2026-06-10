"""
Fail-closed gate: Apply no debe escribir si hay grupos ABONO no habilitados.

Cardinalidad del cuadre (documentado):
- ``parse_accounting_text`` devuelve exactamente un ``PaymentApplicationEvent`` por PDF.
- El cuadre ABONO suma ``valor_pagado_cliente`` una sola vez por ``asiento_pdf_path``
  normalizado dentro del grupo (ver ``reconcile_abono_group``).
"""

from __future__ import annotations

from typing import Any

from app.application.services.abono_dry_run import ABONO_SCHEDULE_RULE_NOT_CONFIGURED

SCHEDULE_RESOLVED = "RESOLVED"
SCHEDULE_NOT_REQUIRED = "NOT_REQUIRED"
_ALLOWED_SCHEDULE_STATUSES = frozenset({SCHEDULE_RESOLVED, SCHEDULE_NOT_REQUIRED, ""})


def _blocking_group_entry(gr: dict[str, Any]) -> dict[str, Any]:
    return {
        "id_pago": gr.get("id_pago"),
        "creditos_seleccionados": list(gr.get("creditos_seleccionados") or []),
        "reconciliation_status": gr.get("reconciliation_status"),
        "schedule_resolution_status": gr.get("schedule_resolution_status"),
        "blocking_errors": list(gr.get("blocking_errors") or []),
    }


def _group_blocks_apply(gr: dict[str, Any]) -> bool:
    if not gr.get("group_ready_for_apply", False):
        return True
    if gr.get("reconciliation_status") != "PASSED":
        return True
    schedule = str(gr.get("schedule_resolution_status") or "").strip()
    if schedule and schedule not in _ALLOWED_SCHEDULE_STATUSES:
        return True
    if gr.get("blocking_errors"):
        return True
    return False


def _resolve_block_error_code(
    dry_run: dict[str, Any],
    blocking_groups: list[dict[str, Any]],
) -> str:
    if dry_run.get("requires_business_rule"):
        for gr in blocking_groups:
            if (
                gr.get("reconciliation_status") == "PASSED"
                and gr.get("schedule_resolution_status") == "NOT_CONFIGURED"
            ):
                return ABONO_SCHEDULE_RULE_NOT_CONFIGURED
    for gr in blocking_groups:
        for err in gr.get("blocking_errors") or []:
            if isinstance(err, dict):
                code = str(err.get("error_code") or "").strip()
                if code:
                    return code
    return "abono_apply_blocked"


def evaluate_abono_apply_block(dry_run: dict[str, Any]) -> dict[str, Any] | None:
    """
    Evalúa si Apply debe abortar por grupos ABONO.

    Devuelve metadatos de bloqueo o None si no hay restricción ABONO.
    """
    abono_results = [
        gr for gr in (dry_run.get("abono_group_results") or []) if isinstance(gr, dict)
    ]
    abono_total = int(dry_run.get("abono_groups_total") or 0)
    if abono_total == 0 and not abono_results:
        return None

    blocking_groups = [_blocking_group_entry(gr) for gr in abono_results if _group_blocks_apply(gr)]

    global_block = (
        dry_run.get("can_apply") is False
        or dry_run.get("requires_business_rule")
        or bool(blocking_groups)
    )
    if not global_block:
        return None

    if not blocking_groups and abono_results:
        blocking_groups = [_blocking_group_entry(gr) for gr in abono_results]

    error_code = _resolve_block_error_code(dry_run, blocking_groups)
    user_message = ""
    next_action = ""
    if error_code == ABONO_SCHEDULE_RULE_NOT_CONFIGURED:
        user_message = (
            "Los asientos del abono cuadran, pero la amortización no puede aplicarse porque "
            "todavía no está definida la regla de fila contractual e IBR."
        )
        next_action = (
            "No vuelva a ejecutar Apply hasta que la regla contable del abono haya sido configurada."
        )
    elif any(
        gr.get("reconciliation_status") != "PASSED" for gr in blocking_groups
    ):
        user_message = (
            "Los asientos del abono no cuadran con el monto bancario; no se modificó ninguna tabla."
        )
        next_action = (
            "Revise montos de asientos, monto_banco del manifest y result.blocking_abono_groups."
        )
    else:
        user_message = (
            "La amortización no puede aplicarse: hay grupos ABONO con errores documentales o de cuadre."
        )
        next_action = (
            "Revise result.blocking_abono_groups y corrija asientos o manifest antes de reintentar Apply."
        )

    return {
        "error_code": error_code,
        "blocking_abono_groups": blocking_groups,
        "abono_groups_total": abono_total or len(abono_results),
        "abono_groups_blocked": len(blocking_groups),
        "requires_business_rule": bool(dry_run.get("requires_business_rule")),
        "can_apply": False,
        "user_message": user_message,
        "next_action": next_action,
    }
