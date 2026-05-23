"""Tests parser texto asiento contable."""

from datetime import date

import pytest

from app.application.services.accounting_pdf_parser import (
    ACCOUNT_CAPITAL,
    ACCOUNT_INTERESES,
    ACCOUNT_MORA,
    ACCOUNT_SALDOS_MENORES,
    ACCOUNT_VALOR_PAGADO_CLIENTE,
    WARNING_BANK_INFERRED,
    AccountingParseError,
    parse_accounting_text,
)

CTX = {
    "id_pago": "7785e37e",
    "cliente": "EQUINORTE",
    "credito": "CREDITO # 258",
    "asiento_pdf_path": "clientes/EQUINORTE/ASIENTOS/asiento.pdf",
}


def _text_credito_258() -> str:
    return """
    Comprobante 12345 Fecha 22/05/2026
    48,497,027.00PAGO: No.Rad. 258 Linea 544111100505
    21,562,855.00PAGO: No.Rad. 258 Linea 544113410519
    26,920,909.00PAGO: No.Rad. 258 Linea 544113430501
    13,263.00PAGO: No.Rad. 258 Linea 544141502030
    """


def _text_credito_265() -> str:
    return """
    1,041,446.00PAGO: No.Rad. 265 Linea 544111100505
    463,050.00PAGO: No.Rad. 265 Linea 544113410519
    578,111.00PAGO: No.Rad. 265 Linea 544113430501
    285.00PAGO: No.Rad. 265 Linea 544141502030
    """


def _text_abono_saldos_menores() -> str:
    return """
    285.00PAGO: No.Rad. 265 Linea 544153159505
    285.00PAGO: No.Rad. 265 Linea 544113410519
    """


def _text_credito_258_pypdf() -> str:
    return """
    Comprobante 12345 Fecha 22/05/2026
    48,497,027.00 PAGO: No.Rad. 258 Linea 544 1 11100505
    21,562,855.00 PAGO: No.Rad. 258 Linea 544 1 13410519
    26,920,909.00 PAGO: No.Rad. 258 Linea 544 1 13430501
    13,263.00 PAGO: No.Rad. 258 Linea 544 1 41502030
    48,497,027.00 48,497,027.00
    """


def _text_credito_265_pypdf() -> str:
    return """
    1,041,446.00 PAGO: No.Rad. 265 Linea 544 1 11100505
    463,050.00 PAGO: No.Rad. 265 Linea 544 1 13410519
    578,111.00 PAGO: No.Rad. 265 Linea 544 1 13430501
    285.00 PAGO: No.Rad. 265 Linea 544 1 41502030
    """


def _text_abono_saldos_menores_pypdf() -> str:
    return """
    285.00 PAGO: No.Rad. 265 Linea 544 1 53159505
    285.00 PAGO: No.Rad. 265 Linea 544 1 13410519
    """


def test_parser_full_payment_breakdown_code_before_amount():
    text = """
    Comprobante 12345 Fecha 22/05/2026
    544111100505 50,000,000.00
    544113410519 49,118,143.00
    544113430501 0.00
    544141502030 881,857.00
    Retenciones intereses facturas 1,234.56
    544153159505 0.00
    """
    ev = parse_accounting_text(text, CTX)
    assert ev.valor_pagado_cliente == 50_000_000.0
    assert ev.capital == 49_118_143.0
    assert ev.intereses == 0.0
    assert ev.mora == 881_857.0
    assert ev.retenciones == 1234.56
    assert ev.saldos_menores == 0.0
    assert ev.fecha_asiento == date(2026, 5, 22)


def test_parser_capital_only_payment_code_before_amount():
    text = """
    544111100505 10,000,000.00
    544113410519 10,000,000.00
    """
    ev = parse_accounting_text(text, CTX)
    assert ev.capital == 10_000_000.0
    assert ev.valor_pagado_cliente == 10_000_000.0


def test_parser_amount_before_code_single_line():
    text = "48,497,027.00PAGO: No.Rad. 258 Linea 544111100505"
    ev = parse_accounting_text(text, CTX)
    assert ev.valor_pagado_cliente == 48_497_027.0


def test_parser_amount_before_code_with_spaces():
    text = "48,497,027.00 PAGO: No.Rad. 258 Línea 544111100505"
    ev = parse_accounting_text(text, CTX)
    assert ev.valor_pagado_cliente == 48_497_027.0


def test_parser_credito_258_real_format():
    ev = parse_accounting_text(_text_credito_258(), CTX)
    assert ev.valor_pagado_cliente == 48_497_027.0
    assert ev.capital == 21_562_855.0
    assert ev.intereses == 26_920_909.0
    assert ev.mora == 13_263.0


def test_parser_credito_265_real_format():
    ev = parse_accounting_text(_text_credito_265(), CTX)
    assert ev.valor_pagado_cliente == 1_041_446.0
    assert ev.capital == 463_050.0
    assert ev.intereses == 578_111.0
    assert ev.mora == 285.0


def test_parser_abono_saldos_menores_without_bank_line():
    ev = parse_accounting_text(_text_abono_saldos_menores(), CTX)
    assert ev.saldos_menores == 285.0
    assert ev.capital == 285.0
    assert ev.valor_pagado_cliente == 285.0
    assert WARNING_BANK_INFERRED in ev.parse_warnings


def test_parser_credito_258_pypdf_split_codes():
    ev = parse_accounting_text(_text_credito_258_pypdf(), CTX)
    assert ev.parser_mode == "split_code_pypdf"
    assert ev.valor_pagado_cliente == 48_497_027.0
    assert ev.capital == 21_562_855.0
    assert ev.intereses == 26_920_909.0
    assert ev.mora == 13_263.0
    assert ACCOUNT_VALOR_PAGADO_CLIENTE in ev.detected_codes


def test_parser_credito_265_pypdf_split_codes():
    ev = parse_accounting_text(_text_credito_265_pypdf(), CTX)
    assert ev.parser_mode == "split_code_pypdf"
    assert ev.valor_pagado_cliente == 1_041_446.0
    assert ev.capital == 463_050.0
    assert ev.intereses == 578_111.0
    assert ev.mora == 285.0


def test_parser_abono_pypdf_split_codes_not_570():
    ev = parse_accounting_text(_text_abono_saldos_menores_pypdf(), CTX)
    assert ev.parser_mode == "split_code_pypdf"
    assert ev.saldos_menores == 285.0
    assert ev.capital == 285.0
    assert ev.valor_pagado_cliente == 285.0
    assert WARNING_BANK_INFERRED in ev.parse_warnings


def test_parser_pypdf_ignores_duplicate_summary_total_line():
    text = _text_credito_258_pypdf()
    ev = parse_accounting_text(text, CTX)
    assert ev.valor_pagado_cliente == 48_497_027.0
    assert ev.valor_pagado_cliente != 48_497_027.0 * 2


def test_parser_sums_duplicate_same_code():
    text = """
    100.00PAGO: No.Rad. 258 Linea 544113410519
    200.00PAGO: No.Rad. 258 Linea 544113410519
    544111100505 50.00
    """
    ev = parse_accounting_text(text, CTX)
    assert ev.capital == 300.0
    assert ev.valor_pagado_cliente == 50.0


def test_parser_infers_banco_when_missing_bank_but_has_component_lines():
    ev = parse_accounting_text(
        "544113410519 100.00\n544113430501 50.00",
        CTX,
    )
    assert ev.valor_pagado_cliente == 150.0
    assert ev.capital == 100.0
    assert ev.intereses == 50.0
    assert ev.parse_warnings
    assert "inferido" in ev.parse_warnings[0].lower()


def test_parser_missing_banco_only_capital_amount_raises_specific_error():
    """Solo código sin monto parseable → sin inferencia posible."""
    with pytest.raises(AccountingParseError) as excinfo:
        parse_accounting_text("544113410519\nsin montos", CTX)
    assert excinfo.value.error_code == "ACCOUNTING_PARSE_FAILED"


def test_parser_missing_everything_raises_generic_error():
    with pytest.raises(AccountingParseError) as excinfo:
        parse_accounting_text("sin datos contables", CTX)
    exc = excinfo.value
    assert exc.error_code == "ACCOUNTING_PARSE_FAILED"
    assert exc.detected_codes == []


def test_parser_error_includes_preview_for_debug():
    with pytest.raises(AccountingParseError) as excinfo:
        parse_accounting_text("544113410519 sin monto adjunto", CTX)
    assert excinfo.value.text_preview
    assert len(excinfo.value.text_preview) <= 500
