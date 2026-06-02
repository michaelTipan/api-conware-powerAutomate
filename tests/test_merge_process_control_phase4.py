from __future__ import annotations

import asyncio
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

import pytest

from app.application.use_cases.merge_composite_validado_pdfs import (
    _auto_detect_bank_ready_for_merge,
    merge_composite_validado_pdfs,
)


def _snap(
    *,
    estado: str,
    is_active: bool,
    hist: str,
    email: str,
):
    return SimpleNamespace(
        control_file_path="CTL/x.xlsx",
        estado_proceso=estado,
        is_active=is_active,
        process_key="payment-validation|banco_bogota|2026-06-01",
        validation_file_path="",
        historical_file_path=hist,
        secretary_file_path="",
        email_pdf_path=email,
        notify_idempotency_key="N",
        merge_manifest_path="",
        merge_idempotency_key="",
        bank_code="",
        bank_name="",
    )


def test_merge_auto_detects_bogota_when_only_bogota_ready():
    async def run():
        with patch(
            "app.application.use_cases.payment_validation_process_control.read_process_control_snapshot",
            new_callable=AsyncMock,
            side_effect=lambda _g, _s, _d, *, bank_code: (
                _snap(
                    estado="PENDIENTE_ASIENTOS",
                    is_active=True,
                    hist="H/a.xlsx",
                    email="E/m.pdf",
                )
                if bank_code == "banco_bogota"
                else _snap(estado="FINALIZADO", is_active=True, hist="H/b.xlsx", email="")
            ),
        ):
            detected, ready = await _auto_detect_bank_ready_for_merge(None, "s", "d")
            assert detected == "banco_bogota"
            assert ready == ["banco_bogota"]

    asyncio.run(run())


def test_merge_auto_detects_bancolombia_when_only_bancolombia_ready():
    async def run():
        with patch(
            "app.application.use_cases.payment_validation_process_control.read_process_control_snapshot",
            new_callable=AsyncMock,
            side_effect=lambda _g, _s, _d, *, bank_code: (
                _snap(
                    estado="PENDIENTE_ASIENTOS",
                    is_active=True,
                    hist="H/a.xlsx",
                    email="E/m.pdf",
                )
                if bank_code == "banco_bancolombia"
                else _snap(estado="FINALIZADO", is_active=True, hist="H/b.xlsx", email="")
            ),
        ):
            detected, ready = await _auto_detect_bank_ready_for_merge(None, "s", "d")
            assert detected == "banco_bancolombia"
            assert ready == ["banco_bancolombia"]

    asyncio.run(run())


def test_merge_auto_detect_returns_none_when_no_ready():
    async def run():
        with patch(
            "app.application.use_cases.payment_validation_process_control.read_process_control_snapshot",
            new_callable=AsyncMock,
            return_value=_snap(estado="FINALIZADO", is_active=True, hist="H/a.xlsx", email=""),
        ):
            detected, ready = await _auto_detect_bank_ready_for_merge(None, "s", "d")
            assert detected is None
            assert ready == []

    asyncio.run(run())


def test_merge_auto_detect_returns_none_when_both_ready():
    async def run():
        with patch(
            "app.application.use_cases.payment_validation_process_control.read_process_control_snapshot",
            new_callable=AsyncMock,
            return_value=_snap(
                estado="PENDIENTE_ASIENTOS",
                is_active=True,
                hist="H/a.xlsx",
                email="E/m.pdf",
            ),
        ):
            detected, ready = await _auto_detect_bank_ready_for_merge(None, "s", "d")
            assert detected is None
            assert ready == ["banco_bogota", "banco_bancolombia"]

    asyncio.run(run())


def test_merge_rejects_invalid_bank_code():
    async def run():
        with pytest.raises(ValueError, match="invalid_bank_code"):
            await merge_composite_validado_pdfs(None, bank_code="nope")  # type: ignore[arg-type]

    asyncio.run(run())

