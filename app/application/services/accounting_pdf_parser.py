"""
Parser de texto extraído de asientos contables (sin OCR) para aplicación de pagos.
"""

from __future__ import annotations

import io
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from typing import Any

from pypdf import PdfReader

ACCOUNT_VALOR_PAGADO_CLIENTE = "544111100505"
ACCOUNT_CAPITAL = "544113410519"
ACCOUNT_INTERESES = "544113430501"
ACCOUNT_MORA = "544141502030"
ACCOUNT_SALDOS_MENORES = "544153159505"

ACCOUNT_PREFIX = "544"

ACCOUNT_CODES = (
    ACCOUNT_VALOR_PAGADO_CLIENTE,
    ACCOUNT_CAPITAL,
    ACCOUNT_INTERESES,
    ACCOUNT_MORA,
    ACCOUNT_SALDOS_MENORES,
)

# Sufijo de 8 dígitos (código corto pypdf) → cuenta canónica de 12 dígitos
SHORT_SUFFIX_TO_ACCOUNT: dict[str, str] = {
    "11100505": ACCOUNT_VALOR_PAGADO_CLIENTE,
    "13410519": ACCOUNT_CAPITAL,
    "13430501": ACCOUNT_INTERESES,
    "41502030": ACCOUNT_MORA,
    "53159505": ACCOUNT_SALDOS_MENORES,
}

RETENCIONES_LABEL = "retenciones intereses facturas"
WARNING_BANK_INFERRED = "BANK_VALUE_INFERRED_OR_MISSING"

_LINEA_TOKEN = r"(?:Linea|Línea|línea|LINEA)"
_AMOUNT_TOKEN = r"\(?[\d]{1,3}(?:,[\d]{3})*(?:\.[\d]+)?\)?"

# Monto antes de PAGO … Linea + código completo (12 dígitos)
_RE_AMOUNT_BEFORE_LINEA_FULL = re.compile(
    rf"({_AMOUNT_TOKEN})\s*PAGO[^\n]*?{_LINEA_TOKEN}\s*(\d{{12}})",
    re.IGNORECASE,
)
# Monto antes de PAGO … Linea 544 [1] {8 dígitos} (extracción pypdf HBI)
_RE_AMOUNT_BEFORE_LINEA_SPLIT = re.compile(
    rf"({_AMOUNT_TOKEN})\s*PAGO[^\n]*?{_LINEA_TOKEN}\s*544\s+(?:1\s+)?(\d{{8}})\b",
    re.IGNORECASE,
)
# Código completo seguido de monto en la misma línea
_RE_CODE_BEFORE_AMOUNT = re.compile(
    rf"(\d{{12}})[^\d\-()\n]*({_AMOUNT_TOKEN})",
    re.IGNORECASE,
)
# Línea resumen con el mismo monto duplicado (sin PAGO / Linea)
_RE_SUMMARY_DUPLICATE_AMOUNT = re.compile(
    rf"^\s*({_AMOUNT_TOKEN})\s+\1\s*$",
    re.IGNORECASE,
)


class AccountingParseError(ValueError):
    """Texto de asiento sin valor pagado cliente identificable."""

    def __init__(
        self,
        message: str,
        *,
        error_code: str = "ACCOUNTING_PARSE_FAILED",
        detected_codes: list[str] | None = None,
        amounts_before_code: bool = False,
        text_preview: str | None = None,
        parser_mode: str = "",
    ) -> None:
        super().__init__(message)
        self.error_code = error_code
        self.detected_codes = detected_codes or []
        self.amounts_before_code = amounts_before_code
        self.text_preview = text_preview
        self.parser_mode = parser_mode


class PdfTextNotExtractableError(ValueError):
    """PDF sin texto extraíble (no OCR)."""


@dataclass(frozen=True)
class PaymentApplicationEvent:
    id_pago: str
    cliente: str
    credito: str
    asiento_pdf_path: str
    comprobante: str
    fecha_asiento: date | None
    valor_pagado_cliente: float
    capital: float
    intereses: float
    mora: float
    retenciones: float
    saldos_menores: float
    raw_text: str
    parse_warnings: tuple[str, ...] = ()
    parser_mode: str = ""
    detected_codes: tuple[str, ...] = ()


def _parse_amount_token(token: str) -> float | None:
    t = token.strip().replace(" ", "")
    if not t:
        return None
    negative = t.startswith("(") and t.endswith(")")
    t = t.strip("()")
    t = t.replace(",", "")
    if t.count(".") > 1:
        parts = t.split(".")
        t = "".join(parts[:-1]) + "." + parts[-1]
    try:
        n = float(t)
    except ValueError:
        return None
    return -n if negative else n


def _normalize_preview(text: str, limit: int = 500) -> str:
    collapsed = re.sub(r"\s+", " ", (text or "").strip())
    return collapsed[:limit]


def _is_known_account_code(code: str) -> bool:
    return code in ACCOUNT_CODES


def _reconstruct_full_code(short_digits: str) -> str | None:
    """544 + sufijo de 8 dígitos → código contable de 12 dígitos."""
    suffix = re.sub(r"\D", "", short_digits)
    if len(suffix) != 8:
        return None
    mapped = SHORT_SUFFIX_TO_ACCOUNT.get(suffix)
    if mapped:
        return mapped
    full = f"{ACCOUNT_PREFIX}{suffix}"
    return full if _is_known_account_code(full) else None


def _is_summary_total_line(line: str) -> bool:
    """Ej. ``48,497,027.00 48,497,027.00`` al pie del asiento — no es línea contable."""
    stripped = line.strip()
    if not stripped or re.search(_LINEA_TOKEN, stripped, re.IGNORECASE):
        return False
    if re.search(r"PAGO", stripped, re.IGNORECASE):
        return False
    return _RE_SUMMARY_DUPLICATE_AMOUNT.match(stripped) is not None


def _line_has_accounting_movement(line: str) -> bool:
    if _is_summary_total_line(line):
        return False
    if re.search(r"PAGO", line, re.IGNORECASE) and re.search(
        _LINEA_TOKEN, line, re.IGNORECASE
    ):
        return True
    if re.search(rf"{_LINEA_TOKEN}\s*544\s", line, re.IGNORECASE):
        return True
    if re.search(r"\d{12}", line):
        return True
    return False


def _merge_matches(
    raw_matches: list[tuple[str, float, int, int, bool, bool]],
) -> tuple[dict[str, float], bool, bool]:
    """Agrupa por código; evita solapamiento de span; suma líneas distintas."""
    accepted: list[tuple[str, float, int, int]] = []
    amounts_before_code = False
    used_split = False
    used_full = False

    for code, amount, start, end, before, split in sorted(raw_matches, key=lambda x: x[2]):
        if any(code == c0 and not (end <= s0 or start >= e0) for c0, _a0, s0, e0 in accepted):
            continue
        accepted.append((code, amount, start, end))
        if before:
            amounts_before_code = True
        if split:
            used_split = True
        else:
            used_full = True

    totals: dict[str, float] = defaultdict(float)
    for code, amount, _s, _e in accepted:
        totals[code] += amount

    return dict(totals), amounts_before_code, used_split, used_full


def _extract_amounts_by_account_code(
    text: str,
) -> tuple[dict[str, float], bool, str, list[str]]:
    """
    Suma montos por código contable.
    Retorna (totales, montos_antes_código, parser_mode, códigos_detectados).
    """
    raw_matches: list[tuple[str, float, int, int, bool, bool]] = []
    line_offset = 0

    for line in (text or "").splitlines():
        if not _line_has_accounting_movement(line):
            line_offset += len(line) + 1
            continue

        for m in _RE_AMOUNT_BEFORE_LINEA_SPLIT.finditer(line):
            code = _reconstruct_full_code(m.group(2))
            if not code:
                continue
            parsed = _parse_amount_token(m.group(1))
            if parsed is None or parsed <= 0:
                continue
            raw_matches.append(
                (code, parsed, line_offset + m.start(), line_offset + m.end(), True, True)
            )

        for m in _RE_AMOUNT_BEFORE_LINEA_FULL.finditer(line):
            code = m.group(2)
            if not _is_known_account_code(code):
                continue
            parsed = _parse_amount_token(m.group(1))
            if parsed is None or parsed <= 0:
                continue
            raw_matches.append(
                (code, parsed, line_offset + m.start(), line_offset + m.end(), True, False)
            )

        for m in _RE_CODE_BEFORE_AMOUNT.finditer(line):
            code = m.group(1)
            if not _is_known_account_code(code):
                continue
            parsed = _parse_amount_token(m.group(2))
            if parsed is None or parsed <= 0:
                continue
            raw_matches.append(
                (code, parsed, line_offset + m.start(), line_offset + m.end(), False, False)
            )

        line_offset += len(line) + 1

    totals, amounts_before, used_split, used_full = _merge_matches(raw_matches)

    if used_split and not used_full:
        parser_mode = "split_code_pypdf"
    elif used_full and not used_split:
        parser_mode = "full_code"
    elif used_split and used_full:
        parser_mode = "mixed"
    else:
        parser_mode = ""

    detected = sorted(k for k, v in totals.items() if v > 0)
    return totals, amounts_before, parser_mode, detected


def _amount_after_label(text: str, label: str) -> float:
    pattern = rf"{re.escape(label)}[^\d\-()]*(\(?[\d.,]+\)?)"
    total = 0.0
    for m in re.finditer(pattern, text, flags=re.IGNORECASE):
        parsed = _parse_amount_token(m.group(1))
        if parsed is not None and parsed > 0:
            total += parsed
    return total


def _extract_comprobante(text: str) -> str:
    m = re.search(
        r"comprobante[^\d]*(\d[\d\-]*)",
        text,
        flags=re.IGNORECASE,
    )
    return m.group(1).strip() if m else ""


def _extract_fecha_asiento(text: str) -> date | None:
    for pat in (
        r"fecha[^\d]*(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})",
        r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})",
    ):
        m = re.search(pat, text, flags=re.IGNORECASE)
        if not m:
            continue
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        try:
            return date(y, mo, d)
        except ValueError:
            continue
    return None


def _is_adjustment_entry(totals: dict[str, float]) -> bool:
    """Asiento de abono/saldos menores sin línea banco ni desglose intereses/mora."""
    bank = totals.get(ACCOUNT_VALOR_PAGADO_CLIENTE, 0.0)
    intereses = totals.get(ACCOUNT_INTERESES, 0.0)
    mora = totals.get(ACCOUNT_MORA, 0.0)
    capital = totals.get(ACCOUNT_CAPITAL, 0.0)
    saldos = totals.get(ACCOUNT_SALDOS_MENORES, 0.0)
    if bank > 0:
        return False
    if intereses > 0 or mora > 0:
        return False
    return saldos > 0 or capital > 0


def _infer_valor_pagado_cliente(
    totals: dict[str, float],
) -> tuple[float, tuple[str, ...]]:
    bank = totals.get(ACCOUNT_VALOR_PAGADO_CLIENTE, 0.0)
    if bank > 0:
        return bank, ()

    capital = totals.get(ACCOUNT_CAPITAL, 0.0)
    intereses = totals.get(ACCOUNT_INTERESES, 0.0)
    mora = totals.get(ACCOUNT_MORA, 0.0)
    saldos = totals.get(ACCOUNT_SALDOS_MENORES, 0.0)

    if _is_adjustment_entry(totals):
        if saldos > 0 and (capital <= 0 or abs(capital - saldos) < 0.01):
            amount = saldos if capital <= 0 else max(capital, saldos)
            return amount, (WARNING_BANK_INFERRED,)
        if capital > 0 and saldos <= 0:
            return capital, (WARNING_BANK_INFERRED,)

    component_sum = capital + intereses + mora + saldos
    if component_sum > 0 and (intereses > 0 or mora > 0):
        warning = (
            f"valor_pagado_cliente inferido desde líneas contables "
            f"(sin {ACCOUNT_VALOR_PAGADO_CLIENTE})"
        )
        return component_sum, (warning,)

    if capital > 0:
        return capital, (WARNING_BANK_INFERRED,)

    return 0.0, ()


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """Extrae texto plano con pypdf; falla si no hay contenido utilizable."""
    if not pdf_bytes or len(pdf_bytes) < 32:
        raise PdfTextNotExtractableError("PDF vacío o demasiado pequeño.")
    reader = PdfReader(io.BytesIO(pdf_bytes))
    chunks: list[str] = []
    for page in reader.pages:
        part = page.extract_text() or ""
        if part.strip():
            chunks.append(part)
    text = "\n".join(chunks).strip()
    if len(text) < 10:
        raise PdfTextNotExtractableError("PDF sin texto extraíble (posible escaneo; no se usa OCR).")
    return text


def parse_accounting_text(text: str, context: dict[str, Any]) -> PaymentApplicationEvent:
    """
    Parsea texto plano del asiento. ``context`` debe incluir id_pago, cliente, credito, asiento_pdf_path.
    """
    raw = text or ""
    totals, amounts_before_code, parser_mode, detected = _extract_amounts_by_account_code(raw)

    valor_pagado, parse_warnings = _infer_valor_pagado_cliente(totals)

    if valor_pagado <= 0:
        preview = _normalize_preview(raw)
        if detected:
            raise AccountingParseError(
                f"No se encontró valor pagado cliente ({ACCOUNT_VALOR_PAGADO_CLIENTE}) "
                f"en el asiento; códigos con monto: {', '.join(detected)}.",
                error_code="MISSING_BANK_VALUE_BUT_HAS_ACCOUNTING_LINES",
                detected_codes=detected,
                amounts_before_code=amounts_before_code,
                text_preview=preview,
                parser_mode=parser_mode,
            )
        raise AccountingParseError(
            f"No se encontró valor pagado cliente ({ACCOUNT_VALOR_PAGADO_CLIENTE}) en el asiento.",
            error_code="ACCOUNTING_PARSE_FAILED",
            detected_codes=detected,
            amounts_before_code=amounts_before_code,
            text_preview=preview,
            parser_mode=parser_mode,
        )

    return PaymentApplicationEvent(
        id_pago=str(context.get("id_pago") or "").strip(),
        cliente=str(context.get("cliente") or "").strip(),
        credito=str(context.get("credito") or "").strip(),
        asiento_pdf_path=str(context.get("asiento_pdf_path") or "").strip(),
        comprobante=_extract_comprobante(raw),
        fecha_asiento=_extract_fecha_asiento(raw),
        valor_pagado_cliente=valor_pagado,
        capital=totals.get(ACCOUNT_CAPITAL, 0.0),
        intereses=totals.get(ACCOUNT_INTERESES, 0.0),
        mora=totals.get(ACCOUNT_MORA, 0.0),
        retenciones=_amount_after_label(raw, RETENCIONES_LABEL),
        saldos_menores=totals.get(ACCOUNT_SALDOS_MENORES, 0.0),
        raw_text=raw,
        parse_warnings=parse_warnings,
        parser_mode=parser_mode,
        detected_codes=tuple(detected),
    )
