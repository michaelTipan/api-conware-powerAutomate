"""Unit tests for payment_validation_settings (no secrets, env-isolated)."""

from __future__ import annotations

import os

import pytest

from app.application.config import payment_validation_settings as pvs


@pytest.fixture(autouse=True)
def _clear_payment_env(monkeypatch: pytest.MonkeyPatch) -> None:
    keys = [
        k
        for k in os.environ
        if k.startswith("PAYMENT_")
        or k.startswith("GRAPH_PAYMENT_VALIDATION")
        or k.startswith("GRAPH_BANK_PAYMENTS")
        or k.startswith("GRAPH_SHAREPOINT_FILE_PATH")
        or k.startswith("PAYMENT_BANK_")
        or k == "GRAPH_VALIDAR_NOTIFY_EMAIL_SUBJECT"
        or k == "GRAPH_VALIDAR_NOTIFY_EXPORT_EMAIL_PDF_NAME_TEMPLATE"
    ]
    for key in keys:
        monkeypatch.delenv(key, raising=False)


def test_paths_from_graph_legacy_vars(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(
        "GRAPH_PAYMENT_VALIDATION_CONTROL_PATH",
        "CUSTOM/00 CONTROL",
    )
    monkeypatch.setenv(
        "GRAPH_PAYMENT_VALIDATION_REVIEW_PATH",
        "CUSTOM/01 REVISION",
    )
    paths = pvs.get_payment_validation_paths()
    assert paths.control == "CUSTOM/00 CONTROL"
    assert paths.review == "CUSTOM/01 REVISION"


def test_paths_from_base_and_subfolders(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PAYMENT_VALIDATION_BASE_FOLDER", "BASE/FLUJO")
    monkeypatch.setenv("PAYMENT_VALIDATION_CONTROL_FOLDER", "CTRL")
    monkeypatch.setenv("PAYMENT_VALIDATION_LOGS_FOLDER", "LOGS")
    paths = pvs.get_payment_validation_paths()
    assert paths.base_folder == "BASE/FLUJO"
    assert paths.control == "BASE/FLUJO/CTRL"
    assert paths.logs == "BASE/FLUJO/LOGS"


def test_bank_bogota_from_graph_bank_payments(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("GRAPH_BANK_PAYMENTS_FILE_PATH", "bancos/BANCO_BOGOTA.xlsx")
    cfg = pvs.get_payment_bank_config(pvs.BANK_CODE_BOGOTA)
    assert cfg.input_file_path == "bancos/BANCO_BOGOTA.xlsx"
    assert cfg.bank_code == pvs.BANK_CODE_BOGOTA


def test_bank_bancolombia_explicit_path(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(
        "GRAPH_BANK_PAYMENTS_FILE_PATH_BANCOLOMBIA",
        "bancos/BANCO_BANCOLOMBIA.xlsx",
    )
    cfg = pvs.get_payment_bank_config(pvs.BANK_CODE_BANCOLOMBIA)
    assert cfg.input_file_path == "bancos/BANCO_BANCOLOMBIA.xlsx"


def test_bank_bancolombia_heuristic_fallback_logs_warning(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    monkeypatch.setenv("GRAPH_BANK_PAYMENTS_FILE_PATH", "x/BANCO_BOGOTA.xlsx")
    cfg = pvs.get_payment_bank_config(pvs.BANK_CODE_BANCOLOMBIA)
    assert cfg.input_file_path == "x/BANCO_BANCOLOMBIA.xlsx"
    assert any("deprecated" in r.message.lower() for r in caplog.records)


def test_canonical_bank_vars_override_defaults(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PAYMENT_BANK_BOGOTA_CODE", "bogota_custom")
    monkeypatch.setenv("PAYMENT_BANK_BOGOTA_NAME", "Mi Banco")
    monkeypatch.setenv("PAYMENT_BANK_BOGOTA_INPUT_FILE_PATH", "in/bogota.xlsx")
    monkeypatch.setenv("PAYMENT_BANK_BOGOTA_CONTROL_FILE", "ctrl/bogota.xlsx")
    monkeypatch.setenv("PAYMENT_BANK_BOGOTA_EMAIL_SUBJECT", "ASUNTO TEST")
    monkeypatch.setenv(
        "PAYMENT_BANK_BOGOTA_EMAIL_PDF_NAME_TEMPLATE",
        "pdf_{fecha}.pdf",
    )
    cfg = pvs.get_payment_bank_config(pvs.BANK_CODE_BOGOTA)
    assert cfg.bank_code == "bogota_custom"
    assert cfg.bank_name == "Mi Banco"
    assert cfg.input_file_path == "in/bogota.xlsx"
    assert cfg.control_file_path == "ctrl/bogota.xlsx"
    assert pvs.resolve_email_subject("banco_bogota") == "ASUNTO TEST"


def test_normalize_bank_code_default() -> None:
    assert pvs.normalize_bank_code(None) == pvs.BANK_CODE_BOGOTA
    assert pvs.normalize_bank_code("") == pvs.BANK_CODE_BOGOTA


def test_validate_bank_code_invalid() -> None:
    with pytest.raises(ValueError, match="invalid_bank_code"):
        pvs.validate_bank_code("banco_xyz")


def test_list_payment_banks_count() -> None:
    banks = pvs.list_payment_banks()
    assert len(banks) == 2
    codes = {b.bank_code for b in banks}
    assert pvs.BANK_CODE_BOGOTA in codes
    assert pvs.BANK_CODE_BANCOLOMBIA in codes


def test_defaults_when_env_empty() -> None:
    control = pvs.resolve_payment_validation_folder(pvs.PaymentValidationFolderName.CONTROL)
    assert "00 CONTROL" in control
    bogota = pvs.resolve_bank_input_file_path(pvs.BANK_CODE_BOGOTA)
    assert bogota.endswith(pvs.DEFAULT_INPUT_FILENAME_BOGOTA)
