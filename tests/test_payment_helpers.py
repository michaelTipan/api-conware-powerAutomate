import pytest
from datetime import datetime, date

try:
    from app.application.services.payment_helpers import (
        parse_bank_date,
        parse_bank_amount,
        extract_client_from_bank_row,
        parse_statement_name,
        extract_from_pdf_fallback,
        extract_credit_id_from_extract_pdf_text,
        find_best_amortization_table,
    )
except ImportError:
    parse_bank_date = None
    parse_bank_amount = None
    extract_client_from_bank_row = None
    parse_statement_name = None
    extract_from_pdf_fallback = None
    find_best_amortization_table = None
    extract_credit_id_from_extract_pdf_text = None

# 1.5 Parseo Banco/Excel
def test_parse_bank_date():
    if parse_bank_date is None:
        pytest.fail("payment_helpers no implementado")
        
    process_date = date(2026, 5, 10)
    
    # datetime real de excel
    assert parse_bank_date(datetime(2025, 12, 30), process_date) == date(2025, 12, 30)
    # texto completo
    assert parse_bank_date("30/12/2025", process_date) == date(2025, 12, 30)
    # texto visible
    assert parse_bank_date("30-dic", process_date) == date(2026, 12, 30)
    assert parse_bank_date("01-ene", process_date) == date(2026, 1, 1)

    with pytest.raises(ValueError, match="bank_date_parse_error"):
        parse_bank_date("fecha-invalida", process_date)

def test_parse_bank_amount():
    if parse_bank_amount is None:
        pytest.fail("payment_helpers no implementado")
        
    # numero directo
    assert parse_bank_amount(25443565) == 25443565.0
    # string colombiano
    assert parse_bank_amount("$ 25.443.565,00") == 25443565.0
    assert parse_bank_amount("25.443.565,00") == 25443565.0
    
    with pytest.raises(ValueError, match="bank_amount_parse_error"):
        parse_bank_amount("monto-invalido")

def test_extract_client_from_bank_row():
    if extract_client_from_bank_row is None:
        pytest.fail("payment_helpers no implementado")
        
    assert extract_client_from_bank_row("GEOEXCON", "Transaccion larga") == "GEOEXCON"
    assert extract_client_from_bank_row("", "PAGO GEOEXCON ABONO") == "PAGO GEOEXCON ABONO"
    assert extract_client_from_bank_row(None, "PAGO GEOEXCON ABONO") == "PAGO GEOEXCON ABONO"
    
    with pytest.raises(ValueError, match="customer_not_found"):
        extract_client_from_bank_row("", "")

# 1.6 Parseo Extractos
def test_parse_statement_name():
    if parse_statement_name is None:
        pytest.fail("payment_helpers no implementado")
        
    # Estandar nuevo
    fecha, cred = parse_statement_name("Extracto 2025-12-23 CREDITO # 254.pdf")
    assert fecha == date(2025, 12, 23)
    assert cred == "254"
    
    fecha2, cred2 = parse_statement_name("extracto 2025-12-22 credito # 265.pdf")
    assert fecha2 == date(2025, 12, 22)
    assert cred2 == "265"
    
    # Estandar viejo
    fecha3, cred3 = parse_statement_name("Extracto Abril 20-03-2026 Obligacio # 265.pdf")
    assert fecha3 == date(2026, 3, 20)
    assert cred3 == "265"
    
    fecha4, cred4 = parse_statement_name("Extracto Mayo 23-04-2026 Obligacion # 258.pdf")
    assert fecha4 == date(2026, 4, 23)
    assert cred4 == "258"
    
    # Invalido
    assert parse_statement_name("Cualquier cosa.pdf") == (None, None)

def test_extract_from_pdf_fallback():
    if extract_from_pdf_fallback is None:
        pytest.fail("payment_helpers no implementado")
        
    # Hacemos un test con texto simulado porque extraer del PDF real requeriria I/O.
    # El helper deberia aceptar texto extraido del PDF (que podemos simular).
    pdf_text = "No. Obligacion 254\nFecha Limite de pago: 2025-12-23\nTOTAL A PAGAR $ 25.443.565,00\nIntereses de mora: 15.000,00"
    
    cred, fecha, total = extract_from_pdf_fallback(pdf_text)
    assert cred == "254"
    assert fecha == date(2025, 12, 23)
    assert total == 25443565.0
    # No extrae mora.

# 1.7 Localizar tabla amortizacion
def test_find_best_amortization_table():
    if find_best_amortization_table is None:
        pytest.fail("payment_helpers no implementado")
        
    files = [
        "~$Tabla de amortizacion.xlsx", # Ignorar
        "Modelo de extracto Equinorte 264.xlsx", # Prioridad media
        "Soporte pago.xlsx" # Descartar
    ]
    
    best = find_best_amortization_table(files, "Equinorte", "264")
    assert best == "Modelo de extracto Equinorte 264.xlsx"
    
    files2 = [
        "Tabla de amortizacion GEOEXCON.xlsx", # Prioridad alta
        "Modelo de extracto.xlsx" # Prioridad media
    ]
    assert find_best_amortization_table(files2, "GEOEXCON", "123") == "Tabla de amortizacion GEOEXCON.xlsx"
    
    # Ambigüedad
    files3 = [
        "Tabla de amortizacion 1.xlsx",
        "Tabla de amortizacion 2.xlsx"
    ]
    with pytest.raises(ValueError, match="amortization_table_ambiguous"):
        find_best_amortization_table(files3, "Cliente", "123")
        
    # No found
    with pytest.raises(ValueError, match="amortization_table_not_found"):
        find_best_amortization_table(["Soporte.xlsx"], "Cliente", "123")


def test_find_best_amortization_table_ignores_pdf_even_if_name_has_tabla():
    if find_best_amortization_table is None:
        pytest.fail("payment_helpers no implementado")

    files = [
        "TABLA DE AMORTIZACION ACIMOR.pdf",
        "Tabla de amortizacion ACIMOR ACIMOR.xlsx",
    ]
    assert find_best_amortization_table(files, "ACIMOR", "ACIMOR") == "Tabla de amortizacion ACIMOR ACIMOR.xlsx"

    with pytest.raises(ValueError, match="amortization_table_not_found"):
        find_best_amortization_table(["TABLA DE AMORTIZACION ACIMOR.pdf"], "ACIMOR", "ACIMOR")


def test_extract_credit_id_from_extract_pdf_text_patterns():
    if extract_credit_id_from_extract_pdf_text is None:
        pytest.fail("payment_helpers no implementado")

    assert extract_credit_id_from_extract_pdf_text("No, Obligación GB-1022-082") == "82"
    assert extract_credit_id_from_extract_pdf_text("No. Obligacion GB-1022-082") == "82"
    assert extract_credit_id_from_extract_pdf_text("Obligación # 82") == "82"
    assert extract_credit_id_from_extract_pdf_text("Obligacion 82") == "82"
    assert extract_credit_id_from_extract_pdf_text("NIT 901.555.123-4\nsin obligación explícita") is None
