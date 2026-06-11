"""Contrato editorial de mensajes operativos visibles para la secretaría."""

from __future__ import annotations

import re

import pytest

from app.application.job_status_enrichment import (
    _FINALIZE_MESSAGES,
    _GENERATE_MESSAGES,
    _GLOBAL_ERROR_MESSAGES,
    enrich_job_for_http_response,
)
from app.application.operational_message_policy import (
    AUDIENCE_BY_CODE,
    get_message_audience,
)

UNMAPPED_CODES = [
    "amortization_table_ambiguous",
    "bank_amount_parse_error",
    "bank_code_and_process_date_required",
    "bank_date_parse_error",
    "control_not_ready_for_dry_run",
    "control_not_ready_for_merge",
    "destination_name_exhausted",
    "missing_distribucion_abonos_headers",
    "missing_distribucion_pagos_sheet",
    "missing_merge_manifest_path",
    "pdf_no_text",
    "preflight_errors",
    "preflight_revision_manual",
    "preflight_warnings_not_allowed",
    "process_control_invalid_structure",
]

FIRST_PERSON = (
    "No encontré",
    "Encontré",
    "No pude",
    "Generé",
    "Guardé",
    "Apliqué",
    "Envié",
    "Hizo bien",
)

INFORMAL_IMPERATIVE = re.compile(
    r"\b(Revisa|Completa|Corrige|Carga|Guarda|Ejecuta|Vuelve|Indica)\b"
)

FORMAL_IMPERATIVE = re.compile(
    r"\b(Revise|Complete|Corrija|Cargue|Guarde|Ejecute|Vuelva|Verifique|Contacte|No continúe|No vuelva)\b"
)

FORBIDDEN_IN_VISIBLE = (
    "La secretaría",
    "El usuario",
    "result.items",
    "result.skipped",
    "result.preflight",
    "result.incomplete_groups",
    "result.blocking",
    "ProcessKey",
    "IdempotencyKey",
    "technical_message",
    "copie el detalle técnico",
    "Copie el detalle técnico",
    "indique bank_code",
    "validation_file_path",
    "merge_manifest_path",
    "historical_file_path",
)

def _failed_job(job_type: str, code: str) -> dict:
    return {
        "job_id": "j",
        "type": job_type,
        "status": "failed",
        "error": {"type": "ValueError", "message": code},
    }


def _collect_catalog_messages() -> list[tuple[str, str, str]]:
    rows: list[tuple[str, str, str]] = []
    for code, (um, na) in _GENERATE_MESSAGES.items():
        rows.append((code, um, na))
    for code, (um, na) in _FINALIZE_MESSAGES.items():
        rows.append((code, um, na))
    for code, (um, na) in _GLOBAL_ERROR_MESSAGES.items():
        rows.append((code, um, na))
    return rows


@pytest.mark.parametrize("code", UNMAPPED_CODES)
def test_unmapped_codes_have_friendly_messages_not_unknown(code: str):
    jt = (
        "amortization_dry_run"
        if code.startswith(
            ("preflight", "control_not_ready_for_dry", "pdf_", "bank_", "amortization", "missing_merge", "process_control")
        )
        else (
            "merge_composite_validado_pdfs"
            if code
            in (
                "control_not_ready_for_merge",
                "destination_name_exhausted",
                "missing_distribucion_pagos_sheet",
                "missing_distribucion_abonos_headers",
            )
            else "amortization_apply"
        )
    )
    out = enrich_job_for_http_response(_failed_job(jt, code))
    e = out["error"]
    assert e["error_code"] == code
    assert e["technical_message"] == code
    assert e["error_code"] != "unknown_error"
    assert "inconveniente técnico" in e["user_message"].lower() or "no fue posible" in e["user_message"].lower() or e["user_message"].strip()
    assert e["user_message"].strip()
    assert e["next_action"].strip()


def test_amount_mismatch_required_wording():
    out = enrich_job_for_http_response(_failed_job("finalize", "amount_mismatch"))
    e = out["error"]
    assert e["error_code"] == "amount_mismatch"
    assert e["technical_message"] == "amount_mismatch"
    assert "Los valores distribuidos no coinciden" in e["user_message"]
    assert "Distribucion_Pagos" in e["next_action"]
    assert "Mora a aplicar" in e["next_action"]
    assert "vuelva a ejecutar la finalización" in e["next_action"].lower()


def test_support_required_codes_contact_soporte_no_technical_tasks():
    support_samples = [
        "bank_code_and_process_date_required",
        "missing_merge_manifest_path",
        "process_control_invalid_structure",
        "graph_config_error",
        "unknown_error",
        "invalid_bank_code",
        "MULTIPLE_READY_PROCESSES",
    ]
    for code in support_samples:
        jt = "amortization_dry_run" if code in ("missing_merge_manifest_path", "graph_config_error") else "generate"
        if code == "MULTIPLE_READY_PROCESSES":
            jt = "finalize"
        out = enrich_job_for_http_response(_failed_job(jt, code))
        na = out["error"]["next_action"].lower()
        assert "contacte a soporte" in na, code
        assert "excel" not in na or "no continúe" in na
        assert "manifest" not in na
        assert "bank_code" not in na


def test_unknown_error_is_support_not_retry_loop():
    out = enrich_job_for_http_response(_failed_job("finalize", "totally_unknown_code_xyz"))
    e = out["error"]
    assert "contacte a soporte" in e["next_action"].lower()
    assert "copie" not in e["next_action"].lower()
    assert "detalle técnico" not in e["next_action"].lower()


@pytest.mark.parametrize("phrase", FIRST_PERSON)
def test_catalog_avoids_first_person(phrase: str):
    for code, um, na in _collect_catalog_messages():
        joined = f"{um} {na}"
        assert phrase not in joined, f"{phrase!r} in {code}"


def test_enriched_errors_avoid_first_and_informal_imperative():
    samples = [
        enrich_job_for_http_response(_failed_job("finalize", "amount_mismatch")),
        enrich_job_for_http_response(_failed_job("generate", "tipo_aplicacion_invalid")),
        enrich_job_for_http_response(_failed_job("finalize", "missing_mora_a_aplicar")),
        enrich_job_for_http_response(_failed_job("amortization_dry_run", "missing_merge_manifest_path")),
    ]
    for out in samples:
        e = out["error"]
        text = f"{e['user_message']} {e['next_action']}"
        for phrase in FIRST_PERSON:
            assert phrase not in text
        assert not INFORMAL_IMPERATIVE.search(e["next_action"]), e["next_action"]
        assert FORMAL_IMPERATIVE.search(e["next_action"]), e["next_action"]


def test_enriched_errors_avoid_forbidden_phrases():
    samples = [
        enrich_job_for_http_response(_failed_job("finalize", "amount_mismatch")),
        enrich_job_for_http_response(
            {
                "job_id": "j",
                "type": "merge_composite_validado_pdfs",
                "status": "completed",
                "result": {"outputs": [], "skipped": ["x"], "skipped_count": 1},
            }
        ),
        enrich_job_for_http_response(
            {
                "job_id": "j",
                "type": "merge_composite_validado_pdfs",
                "status": "completed",
                "result": {"outputs": [{"id_pago": "1"}], "incomplete_groups_count": 2},
            }
        ),
    ]
    for out in samples:
        texts = [out.get("user_message", ""), out.get("next_action", "")]
        err = out.get("error") or {}
        if isinstance(err, dict):
            texts.extend([err.get("user_message", ""), err.get("next_action", "")])
        joined = " ".join(t for t in texts if t)
        for bad in FORBIDDEN_IN_VISIBLE:
            assert bad not in joined, f"found {bad!r} in {joined!r}"


def test_secretary_can_correct_indicates_operational_fix():
    out = enrich_job_for_http_response(_failed_job("finalize", "missing_mora_a_aplicar"))
    na = out["error"]["next_action"]
    assert get_message_audience("missing_mora_a_aplicar") == "SECRETARY_CAN_CORRECT"
    assert "Distribucion" in na or "distribución" in na.lower() or "Mora" in na or "vuelva" in na.lower()


def test_error_code_and_technical_message_preserved():
    raw = _failed_job("finalize", "amount_mismatch")
    out = enrich_job_for_http_response(raw)
    assert out["error"]["error_code"] == "amount_mismatch"
    assert out["error"]["technical_message"] == "amount_mismatch"
    assert out["error"]["message"] == "amount_mismatch"


def test_merge_partial_uses_count_not_result_skipped():
    out = enrich_job_for_http_response(
        {
            "job_id": "j",
            "type": "merge_composite_validado_pdfs",
            "status": "completed",
            "result": {"outputs": [], "skipped": ["a", "b", "c"], "skipped_count": 3},
        }
    )
    um = out["user_message"]
    assert "3" in um
    assert "result.skipped" not in um
    assert "result.skipped" not in out.get("next_action", "")


def test_pdf_no_text_correct_then_support():
    out = enrich_job_for_http_response(_failed_job("amortization_dry_run", "pdf_no_text"))
    assert "contacte a soporte" in out["error"]["next_action"].lower()
    assert get_message_audience("pdf_no_text") == "CORRECT_THEN_SUPPORT"
