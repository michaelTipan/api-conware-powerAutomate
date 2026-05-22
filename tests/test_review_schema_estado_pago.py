"""EstadoPago.ADELANTADO en conjuntos operativos de Finalize."""

from app.application.services.review_schema import EstadoPago


def test_adelantado_in_counters_positive_total():
    assert EstadoPago.ADELANTADO in EstadoPago.COUNTERS_POSITIVE_TOTAL


def test_adelantado_in_secretary_and_ruta():
    assert EstadoPago.ADELANTADO in EstadoPago.SECRETARY_AND_RUTA


def test_adelantado_in_clears_pending():
    assert EstadoPago.ADELANTADO in EstadoPago.CLEARS_PENDING
