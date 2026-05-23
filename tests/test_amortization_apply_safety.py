"""Tests de seguridad para apply de amortización."""

import pytest

from app.application.services.accounting_pdf_parser import (
    ACCOUNT_CAPITAL,
    ACCOUNT_SALDOS_MENORES,
    WARNING_BANK_INFERRED,
)
from app.application.services.amortization_apply_safety import (
    check_idempotency_against_log,
    collect_disallowed_warning_items,
    item_warnings_allowed,
)
from app.application.services.amortization_workbook import AutomationLogRecord
from app.application.use_cases.amortization_fill_apply import (
    AmortizationPreflightError,
    validate_amortization_preflight,
)


def test_item_warnings_allowed_saldos_menores():
    item = {
        "warnings": [WARNING_BANK_INFERRED],
        "detected_codes": [ACCOUNT_SALDOS_MENORES, ACCOUNT_CAPITAL],
        "payment_application": {
            "valor_pagado_cliente": 100.0,
            "saldos_menores": 100.0,
            "capital": 100.0,
        },
    }
    assert item_warnings_allowed(item) is True


def test_item_warnings_allowed_rejects_unknown_warning():
    item = {
        "warnings": ["OTRO_WARNING"],
        "detected_codes": [ACCOUNT_SALDOS_MENORES, ACCOUNT_CAPITAL],
        "payment_application": {"valor_pagado_cliente": 100.0, "saldos_menores": 100.0, "capital": 100.0},
    }
    assert item_warnings_allowed(item) is False


def test_preflight_blocks_disallowed_warnings():
    dry = {
        "summary": {"errors": 0, "revision_manual": 0},
        "items": [
            {
                "warnings": ["valor_pagado_cliente inferido desde líneas contables"],
                "detected_codes": ["544113410519", "544113430501"],
                "payment_application": {"valor_pagado_cliente": 1.0},
            }
        ],
    }
    with pytest.raises(AmortizationPreflightError) as exc:
        validate_amortization_preflight(dry)
    assert exc.value.error_code == "preflight_warnings_not_allowed"


def test_idempotency_same_key_different_hash_blocks():
    item = {"idempotency_key": "k1", "asiento_pdf_hash": "hash-b"}
    log_index = {
        "k1": AutomationLogRecord(
            idempotency_key="k1",
            pdf_hash="hash-a",
            pdf_etag="",
        )
    }
    result = check_idempotency_against_log(item, log_index)
    assert result.pdf_changed is True
    assert result.skip_idempotent is False


def test_idempotency_same_hash_skips():
    item = {"idempotency_key": "k1", "asiento_pdf_hash": "hash-a"}
    log_index = {
        "k1": AutomationLogRecord(idempotency_key="k1", pdf_hash="hash-a", pdf_etag="")
    }
    result = check_idempotency_against_log(item, log_index)
    assert result.skip_idempotent is True


def test_collect_disallowed_warning_items_skips_errors():
    dry = {
        "items": [
            {"error_code": "X", "warnings": ["cualquiera"]},
            {"warnings": [], "detected_codes": []},
        ]
    }
    assert collect_disallowed_warning_items(dry) == []
