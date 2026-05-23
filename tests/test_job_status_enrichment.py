"""Tests de enriquecimiento de jobs (user_message, next_action, severity, error estándar)."""

import json

from app.application.job_status_enrichment import (
    enrich_job_for_http_response,
    examples_for_parse_json_tests,
)


def _completed(job_type: str, result: dict) -> dict:
    return {
        "job_id": "jid",
        "type": job_type,
        "status": "completed",
        "result": result,
    }


def test_generate_completed_job_includes_user_message_next_action_severity():
    raw = _completed("generate", {"validation_file": "f.xlsx", "summary": {}})
    out = enrich_job_for_http_response(raw)
    assert out["severity"] == "success"
    assert "archivo de revisión de pagos" in out["user_message"].lower()
    assert "Procesar" in out["next_action"] or "SI" in out["next_action"]


def test_finalize_completed_job_includes_user_message_next_action_severity():
    raw = _completed("finalize", {"validated_rows": 1, "historical_file_path": "h.xlsx"})
    out = enrich_job_for_http_response(raw)
    assert out["severity"] == "success"
    assert "cerrada correctamente" in out["user_message"].lower()
    assert "soporte" in out["user_message"].lower() or "asientos" in out["user_message"].lower()


def test_notify_completed_job_includes_user_message_next_action_severity():
    raw = _completed(
        "notify_validar_extractos",
        {"status": "ok", "historico_excel_path": "p.xlsx", "attachments_count": 2},
    )
    out = enrich_job_for_http_response(raw)
    assert out["severity"] == "success"
    assert "correo" in out["user_message"].lower()
    assert "pdfs" in out["next_action"].lower() or "bandeja" in out["next_action"].lower()


def test_notify_completed_merge_control_active_process_has_warning_severity():
    raw = _completed(
        "notify_validar_extractos",
        {
            "status": "ok",
            "merge_control_updated": False,
            "merge_control_error_code": "merge_control_active_process_exists",
            "merge_control_warning": "Ya existe un proceso pendiente",
        },
    )
    out = enrich_job_for_http_response(raw)
    assert out["severity"] == "warning"
    assert "pendiente" in out["next_action"].lower()


def test_notify_completed_missing_email_pdf_for_merge_control_has_warning_severity():
    raw = _completed(
        "notify_validar_extractos",
        {
            "status": "ok",
            "merge_control_updated": False,
            "merge_control_error_code": "missing_email_pdf_path_for_merge_control",
        },
    )
    out = enrich_job_for_http_response(raw)
    assert out["severity"] == "warning"
    assert "pdf" in out["next_action"].lower()


def test_merge_completed_job_includes_user_message_next_action_severity():
    raw = _completed(
        "merge_composite_validado_pdfs",
        {"outputs": [{"id_pago": "1", "output_relative_path": "out/1.pdf"}], "skipped": []},
    )
    out = enrich_job_for_http_response(raw)
    assert out["severity"] == "success"
    assert "pdf" in out["user_message"].lower()


def test_merge_completed_with_outputs_empty_and_skipped_has_warning_severity():
    raw = _completed(
        "merge_composite_validado_pdfs",
        {"outputs": [], "skipped": ["id1: sin rutas"]},
    )
    out = enrich_job_for_http_response(raw)
    assert out["severity"] == "warning"
    assert "sin generar" in out["user_message"].lower() or "asiento" in out["user_message"].lower()


def test_merge_completed_with_outputs_and_skipped_warning():
    raw = _completed(
        "merge_composite_validado_pdfs",
        {
            "outputs": [{"id_pago": "1", "output_relative_path": "a.pdf"}],
            "skipped": ["x: skip"],
        },
    )
    out = enrich_job_for_http_response(raw)
    assert out["severity"] == "warning"


def test_generate_failed_known_error_returns_standard_error_object():
    raw = {
        "job_id": "j",
        "type": "generate",
        "status": "failed",
        "error": {"type": "ValueError", "message": "review_folder_not_empty"},
    }
    out = enrich_job_for_http_response(raw)
    assert out["severity"] == "error"
    e = out["error"]
    assert e["message"] == "review_folder_not_empty"
    assert e["error_code"] == "review_folder_not_empty"
    assert e["technical_message"] == "review_folder_not_empty"
    assert "01 revision" in e["user_message"].lower()
    assert e["next_action"]


def test_finalize_failed_amount_mismatch_returns_standard_error_object():
    raw = {
        "job_id": "j",
        "type": "finalize",
        "status": "failed",
        "error": {"type": "ValueError", "message": "amount_mismatch"},
    }
    out = enrich_job_for_http_response(raw)
    e = out["error"]
    assert e["error_code"] == "amount_mismatch"
    assert "no suman" in e["user_message"].lower()


def test_notify_failed_string_error_is_converted_to_standard_error_object():
    raw = {
        "job_id": "j",
        "type": "notify_validar_extractos",
        "status": "failed",
        "error": "No se encontró en CORREOS.xlsx una hoja con columnas",
    }
    out = enrich_job_for_http_response(raw)
    e = out["error"]
    assert isinstance(e, dict)
    assert e["message"] == raw["error"]
    assert e["technical_message"] == raw["error"]
    assert "receptores" in e["user_message"].lower() or "emisor" in e["user_message"].lower()


def test_merge_failed_string_error_is_converted_to_standard_error_object():
    msg = "No se encontró PDF de correo en '05' con la fecha del reporte (2026-01-01). Archivos .pdf vistos: (ninguno)"
    raw = {
        "job_id": "j",
        "type": "merge_composite_validado_pdfs",
        "status": "failed",
        "error": msg,
    }
    out = enrich_job_for_http_response(raw)
    e = out["error"]
    assert e["message"] == msg
    assert "correo" in e["user_message"].lower()


def test_unknown_error_returns_generic_user_message():
    raw = {
        "job_id": "j",
        "type": "generate",
        "status": "failed",
        "error": {"type": "RuntimeError", "message": "totally_unknown_code_xyz_123"},
    }
    out = enrich_job_for_http_response(raw)
    assert "no pudo clasificar" in out["error"]["user_message"]


def test_existing_job_fields_remain_backward_compatible():
    raw = {
        "job_id": "j1",
        "type": "finalize",
        "status": "completed",
        "queued_at": "t",
        "finished_at": "t2",
        "result": {"historical_file_url": "https://x", "secretary_file_url": "https://y", "validated_rows": 3},
    }
    out = enrich_job_for_http_response(raw)
    assert out["result"]["historical_file_url"] == "https://x"
    assert out["result"]["validated_rows"] == 3
    assert out["job_id"] == "j1"


def test_enrich_does_not_mutate_input_document():
    raw = {"job_id": "j", "type": "generate", "status": "completed", "result": {"a": 1}}
    enrich_job_for_http_response(raw)
    assert "user_message" not in raw


def test_parse_json_flexible_schema_can_parse_completed_and_failed_examples():
    completed, failed = examples_for_parse_json_tests()
    assert json.loads(json.dumps(completed))["status"] == "completed"
    assert json.loads(json.dumps(failed))["error"]["error_code"] == "amount_mismatch"


def test_non_enrichable_job_type_unchanged():
    raw = {"job_id": "x", "type": "ensure_asientos_contables", "status": "completed", "result": {}}
    out = enrich_job_for_http_response(raw)
    assert "user_message" not in out
