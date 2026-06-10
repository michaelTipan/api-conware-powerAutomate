"""
Gate fail-closed: Dry-run y Apply bloqueados si el manifest Merge está incompleto.
"""

from __future__ import annotations

from typing import Any

from app.application.services.merge_group_validation import (
    MANIFEST_STATUS_COMPLETE,
    MANIFEST_STATUS_PARTIAL,
    MERGE_GROUP_COMPLETE,
    output_creditos_match_expected,
)

MERGE_INCOMPLETE_NOT_APPLICABLE = "MERGE_INCOMPLETE_NOT_APPLICABLE"
APPLY_EXPECTED_EVENTS_INCOMPLETE = "APPLY_EXPECTED_EVENTS_INCOMPLETE"
MERGE_GROUP_PENDING_INPUTS = "MERGE_GROUP_PENDING_INPUTS"
MERGE_EXPECTED_CREDITS_MISMATCH = "MERGE_EXPECTED_CREDITS_MISMATCH"


def _norm_manifest_status(manifest: dict[str, Any]) -> str:
    raw = str(manifest.get("manifest_status") or "").strip().upper()
    if raw in (MANIFEST_STATUS_COMPLETE, MANIFEST_STATUS_PARTIAL):
        return raw
    incomplete = manifest.get("incomplete_groups") or []
    skipped = manifest.get("skipped") or []
    outputs = manifest.get("outputs") or []
    if incomplete or skipped:
        return MANIFEST_STATUS_PARTIAL
    for out in outputs:
        if not isinstance(out, dict):
            continue
        if not output_creditos_match_expected(out):
            return MANIFEST_STATUS_PARTIAL
        status = str(out.get("status") or MERGE_GROUP_COMPLETE).strip()
        if status and status != MERGE_GROUP_COMPLETE:
            return MANIFEST_STATUS_PARTIAL
    if outputs:
        return MANIFEST_STATUS_COMPLETE
    return MANIFEST_STATUS_PARTIAL


def assess_manifest_completeness(manifest: dict[str, Any]) -> dict[str, Any]:
    """Evalúa si el manifest habilita Dry-run/Apply."""
    status = _norm_manifest_status(manifest)
    incomplete_groups = [
        g for g in (manifest.get("incomplete_groups") or []) if isinstance(g, dict)
    ]
    incomplete_count = int(
        manifest.get("incomplete_groups_count")
        if manifest.get("incomplete_groups_count") is not None
        else len(incomplete_groups)
    )
    eligible = bool(manifest.get("eligible_for_dry_run"))
    if manifest.get("eligible_for_dry_run") is None:
        eligible = status == MANIFEST_STATUS_COMPLETE and incomplete_count == 0

    output_issues: list[dict[str, Any]] = []
    for out in manifest.get("outputs") or []:
        if not isinstance(out, dict):
            continue
        out_status = str(out.get("status") or MERGE_GROUP_COMPLETE).strip()
        if out_status != MERGE_GROUP_COMPLETE:
            output_issues.append(
                {
                    "id_pago": out.get("id_pago"),
                    "error_code": MERGE_GROUP_PENDING_INPUTS,
                    "status": out_status,
                }
            )
        if not output_creditos_match_expected(out):
            output_issues.append(
                {
                    "id_pago": out.get("id_pago"),
                    "error_code": MERGE_EXPECTED_CREDITS_MISMATCH,
                }
            )
        if out.get("eligible_for_dry_run") is False:
            output_issues.append(
                {
                    "id_pago": out.get("id_pago"),
                    "error_code": MERGE_INCOMPLETE_NOT_APPLICABLE,
                }
            )

    return {
        "manifest_status": status,
        "eligible_for_dry_run": eligible and not output_issues,
        "incomplete_groups_count": incomplete_count,
        "output_issues": output_issues,
        "incomplete_groups": incomplete_groups,
    }


def evaluate_merge_incomplete_block(
    manifest: dict[str, Any],
    *,
    estado_proceso: str = "",
    merge_skipped_count: int | None = None,
) -> dict[str, Any] | None:
    """
    Devuelve metadatos de bloqueo o None si el manifest permite continuar.
    """
    assessment = assess_manifest_completeness(manifest)
    estado = str(estado_proceso or "").strip()
    skipped = merge_skipped_count
    if skipped is None:
        skipped = len(manifest.get("skipped") or [])

    reasons: list[str] = []
    if estado and estado not in ("CONSOLIDADO",):
        reasons.append(f"estado_proceso={estado}")
    if assessment["manifest_status"] != MANIFEST_STATUS_COMPLETE:
        reasons.append(f"manifest_status={assessment['manifest_status']}")
    if not assessment["eligible_for_dry_run"]:
        reasons.append("eligible_for_dry_run=false")
    if assessment["incomplete_groups_count"] > 0:
        reasons.append(f"incomplete_groups_count={assessment['incomplete_groups_count']}")
    if skipped and int(skipped) > 0:
        reasons.append(f"merge_skipped_count={skipped}")
    if assessment["output_issues"]:
        reasons.append("output_validation_failed")

    if not reasons:
        return None

    error_code = MERGE_INCOMPLETE_NOT_APPLICABLE
    if any(i.get("error_code") == MERGE_EXPECTED_CREDITS_MISMATCH for i in assessment["output_issues"]):
        error_code = MERGE_EXPECTED_CREDITS_MISMATCH
    elif assessment["incomplete_groups_count"] > 0:
        error_code = MERGE_GROUP_PENDING_INPUTS

    return {
        "error_code": error_code,
        "blocked": True,
        "can_apply": False,
        "manifest_status": assessment["manifest_status"],
        "eligible_for_dry_run": assessment["eligible_for_dry_run"],
        "incomplete_groups_count": assessment["incomplete_groups_count"],
        "incomplete_groups": assessment["incomplete_groups"],
        "output_issues": assessment["output_issues"],
        "blocking_reasons": reasons,
        "user_message": (
            "El proceso todavía tiene grupos incompletos en Merge y no puede continuar "
            "al llenado de amortización."
        ),
        "next_action": (
            "Cargue los documentos faltantes y vuelva a ejecutar Merge."
        ),
    }


def count_expected_credit_events(manifest: dict[str, Any]) -> int:
    """Eventos esperados: un bloque por crédito en cada output COMPLETE."""
    total = 0
    for out in manifest.get("outputs") or []:
        if not isinstance(out, dict):
            continue
        if str(out.get("status") or MERGE_GROUP_COMPLETE) != MERGE_GROUP_COMPLETE:
            continue
        items = out.get("credit_items") or []
        if items:
            total += len(items)
        else:
            paths = out.get("asiento_pdf_paths") or []
            total += max(1, len(paths))
    return total


def assess_apply_event_completeness(
    manifest: dict[str, Any],
    apply_items: list[dict[str, Any]],
) -> dict[str, Any]:
    """Compara eventos esperados del manifest vs aplicados/idempotentes."""
    expected = count_expected_credit_events(manifest)
    terminal_statuses = frozenset(
        {"APPLIED", "ADOPTED", "SKIPPED_IDEMPOTENT", "applied", "adopted", "skipped_idempotent"}
    )
    completed = sum(
        1
        for it in apply_items
        if isinstance(it, dict) and str(it.get("apply_status") or "") in terminal_statuses
    )
    missing: list[dict[str, Any]] = []
    seen_keys: set[tuple[str, str]] = set()
    for it in apply_items:
        if not isinstance(it, dict):
            continue
        if str(it.get("apply_status") or "") in terminal_statuses:
            key = (str(it.get("id_pago") or ""), str(it.get("credito") or ""))
            seen_keys.add(key)

    for out in manifest.get("outputs") or []:
        if not isinstance(out, dict):
            continue
        id_pago = str(out.get("id_pago") or "")
        for ci in out.get("credit_items") or []:
            if not isinstance(ci, dict):
                continue
            cred = str(ci.get("credito") or "")
            if (id_pago, cred) not in seen_keys:
                missing.append({"id_pago": id_pago, "credito": cred})

    if expected == 0:
        all_done = not missing
    else:
        all_done = completed >= expected and not missing
    return {
        "expected_credit_events_count": expected,
        "completed_credit_events_count": completed,
        "missing_credit_events": missing,
        "all_expected_events_completed": all_done,
        "error_code": APPLY_EXPECTED_EVENTS_INCOMPLETE if not all_done else "",
    }
