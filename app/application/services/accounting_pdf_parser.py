"""
Parser de texto extraído de asientos contables (sin OCR) para aplicación de pagos.

Reglas de negocio (cuenta corta / sufijo 8 dígitos salvo indicación):
- 11100505 → recaudo Banco Bogotá (valor pagado cliente)
- 11300502 / 113005002 → recaudo Bancolombia (valor pagado cliente)
- 1341* → capital
- 1343* → intereses corrientes
- 1355* → retenciones practicadas
- 41502030 → mora
- 53159505 (544…) → saldos menores (asientos de ajuste HBI)
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

# Sufijo 8 dígitos (código corto pypdf) → cuenta canónica de 12 dígitos (compatibilidad HBI)
SHORT_SUFFIX_TO_ACCOUNT: dict[str, str] = {
    "11100505": ACCOUNT_VALOR_PAGADO_CLIENTE,
    "11300502": ACCOUNT_VALOR_PAGADO_CLIENTE,
    "13410519": ACCOUNT_CAPITAL,
    "13430501": ACCOUNT_INTERESES,
    "41502030": ACCOUNT_MORA,
    "53159505": ACCOUNT_SALDOS_MENORES,
}

BANK_BOGOTA_SUFFIX = "11100505"
BANK_BANCOLOMBIA_SUFFIX = "11300502"
MORA_SUFFIX = "41502030"
SALDOS_MENORES_SUFFIX = "53159505"

RETENCIONES_LABEL = "retenciones intereses facturas"
WARNING_BANK_INFERRED = "BANK_VALUE_INFERRED_OR_MISSING"

_BUCKET_BANK = "bank"
_BUCKET_CAPITAL = "capital"
_BUCKET_INTERESES = "intereses"
_BUCKET_MORA = "mora"
_BUCKET_RETENCIONES = "retenciones"
_BUCKET_SALDOS = "saldos_menores"

_LINEA_TOKEN = r"(?:Linea|Línea|línea|LINEA)"
_AMOUNT_TOKEN = r"\(?[\d]{1,3}(?:,[\d]{3})*(?:\.[\d]+)?\)?"

_RE_AMOUNT_BEFORE_LINEA_FULL = re.compile(
    rf"({_AMOUNT_TOKEN})\s*PAGO[^\n]*?{_LINEA_TOKEN}\s*(\d{{12}})",
    re.IGNORECASE,
)
_RE_AMOUNT_BEFORE_LINEA_SPLIT = re.compile(
    rf"({_AMOUNT_TOKEN})\s*PAGO[^\n]*?{_LINEA_TOKEN}\s*544\s+(?:1\s+)?(\d{{8,9}})\b",
    re.IGNORECASE,
)
_RE_CODE_BEFORE_AMOUNT = re.compile(
    rf"(\d{{8,12}})[^\d\-()\n]*({_AMOUNT_TOKEN})",
    re.IGNORECASE,
)
_RE_ACCOUNT_AT_LINE_START = re.compile(
    rf"^\s*(\d{{8,12}})\b[^\d\-()\n]*({_AMOUNT_TOKEN})",
    re.IGNORECASE,
)
_RE_ACCOUNT_LINE_AMOUNT_AT_END = re.compile(
    rf"^\s*(\d{{8,12}})\b.*?({_AMOUNT_TOKEN})\s*$",
    re.IGNORECASE,
)
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


def _is_plausible_line_amount(token: str, parsed: float) -> bool:
    """Evita confundir la columna ``1`` del asiento con un monto."""
    if parsed <= 0:
        return False
    t = token.strip()
    if "," in t or "." in t:
        return True
    if parsed >= 1000:
        return True
    return False


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


def _normalize_account_suffix(raw_digits: str) -> str | None:
    """Normaliza a sufijo de 8 dígitos para clasificación por prefijo."""
    d = re.sub(r"\D", "", raw_digits or "")
    if not d:
        return None
    if len(d) == 12 and d.startswith(ACCOUNT_PREFIX):
        return d[-8:]
    if d in ("113005002",):
        return BANK_BANCOLOMBIA_SUFFIX
    if len(d) == 8:
        return d
    if 9 <= len(d) <= 12:
        return d[-8:]
    return None


def _suffix_business_bucket(suffix: str) -> str | None:
    """Clasifica sufijo 8 dígitos según reglas de negocio."""
    if suffix == BANK_BOGOTA_SUFFIX:
        return _BUCKET_BANK
    if suffix == BANK_BANCOLOMBIA_SUFFIX:
        return _BUCKET_BANK
    if suffix.startswith("1341"):
        return _BUCKET_CAPITAL
    if suffix.startswith("1343"):
        return _BUCKET_INTERESES
    if suffix.startswith("1355"):
        return _BUCKET_RETENCIONES
    if suffix == MORA_SUFFIX:
        return _BUCKET_MORA
    if suffix == SALDOS_MENORES_SUFFIX:
        return _BUCKET_SALDOS
    canon = SHORT_SUFFIX_TO_ACCOUNT.get(suffix)
    if canon == ACCOUNT_VALOR_PAGADO_CLIENTE:
        return _BUCKET_BANK
    if canon == ACCOUNT_CAPITAL:
        return _BUCKET_CAPITAL
    if canon == ACCOUNT_INTERESES:
        return _BUCKET_INTERESES
    if canon == ACCOUNT_MORA:
        return _BUCKET_MORA
    if canon == ACCOUNT_SALDOS_MENORES:
        return _BUCKET_SALDOS
    return None


def _is_classifiable_suffix(suffix: str) -> bool:
    return _suffix_business_bucket(suffix) is not None


def _parse_account_suffix(raw_digits: str) -> str | None:
    suffix = _normalize_account_suffix(raw_digits)
    if suffix and _is_classifiable_suffix(suffix):
        return suffix
    return None


def _is_known_account_code(code: str) -> bool:
    suffix = _parse_account_suffix(code)
    return suffix is not None


def _reconstruct_full_code(short_digits: str) -> str | None:
    """Compatibilidad: devuelve código canónico 12 dígitos si existe en mapa HBI."""
    suffix = _parse_account_suffix(short_digits)
    if not suffix:
        return None
    return SHORT_SUFFIX_TO_ACCOUNT.get(suffix) or (
        f"{ACCOUNT_PREFIX}{suffix}"
        if f"{ACCOUNT_PREFIX}{suffix}" in ACCOUNT_CODES
        else None
    )


def _aggregate_suffix_totals(
    suffix_totals: dict[str, float],
) -> tuple[dict[str, float], list[str], float]:
    """Agrega por sufijo → totales canónicos (claves ACCOUNT_*) y códigos detectados."""
    buckets: dict[str, float] = defaultdict(float)
    for suffix, amount in suffix_totals.items():
        bucket = _suffix_business_bucket(suffix)
        if bucket:
            buckets[bucket] += amount

    totals = {
        ACCOUNT_VALOR_PAGADO_CLIENTE: buckets[_BUCKET_BANK],
        ACCOUNT_CAPITAL: buckets[_BUCKET_CAPITAL],
        ACCOUNT_INTERESES: buckets[_BUCKET_INTERESES],
        ACCOUNT_MORA: buckets[_BUCKET_MORA],
        ACCOUNT_SALDOS_MENORES: buckets[_BUCKET_SALDOS],
    }
    retenciones_suffix = buckets[_BUCKET_RETENCIONES]

    detected: list[str] = []
    if totals[ACCOUNT_VALOR_PAGADO_CLIENTE] > 0:
        detected.append(ACCOUNT_VALOR_PAGADO_CLIENTE)
    if totals[ACCOUNT_CAPITAL] > 0:
        detected.append(ACCOUNT_CAPITAL)
    if totals[ACCOUNT_INTERESES] > 0:
        detected.append(ACCOUNT_INTERESES)
    if totals[ACCOUNT_MORA] > 0:
        detected.append(ACCOUNT_MORA)
    if totals[ACCOUNT_SALDOS_MENORES] > 0:
        detected.append(ACCOUNT_SALDOS_MENORES)
    if retenciones_suffix > 0:
        detected.append("1355")

    return totals, detected, retenciones_suffix


def _resolve_retenciones(suffix_retenciones: float, raw_text: str) -> float:
    from_label = _amount_after_label(raw_text, RETENCIONES_LABEL)
    if suffix_retenciones > 0:
        return suffix_retenciones
    return from_label


def _is_summary_total_line(line: str) -> bool:
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
    if re.search(r"\d{8,12}", line):
        return True
    return False


def _merge_suffix_matches(
    raw_matches: list[tuple[str, float, int, int, bool, bool]],
) -> tuple[dict[str, float], bool, bool]:
    accepted: list[tuple[str, float, int, int]] = []
    amounts_before_code = False
    used_split = False
    used_full = False

    for suffix, amount, start, end, before, split in sorted(raw_matches, key=lambda x: x[2]):
        if any(
            suffix == s0 and not (end <= st or start >= e0)
            for s0, _a0, st, e0 in accepted
        ):
            continue
        accepted.append((suffix, amount, start, end))
        if before:
            amounts_before_code = True
        if split:
            used_split = True
        else:
            used_full = True

    totals: dict[str, float] = defaultdict(float)
    for suffix, amount, _s, _e in accepted:
        totals[suffix] += amount

    return dict(totals), amounts_before_code, used_split, used_full


def _extract_amounts_by_account_code(
    text: str,
) -> tuple[dict[str, float], bool, str, list[str], float]:
    """
    Suma montos por sufijo de cuenta y agrega a totales canónicos.
    Retorna (totales, montos_antes_código, parser_mode, códigos_detectados, retenciones_sufijo).
    """
    raw_matches: list[tuple[str, float, int, int, bool, bool]] = []
    line_offset = 0

    for line in (text or "").splitlines():
        if not _line_has_accounting_movement(line):
            line_offset += len(line) + 1
            continue

        for m in _RE_AMOUNT_BEFORE_LINEA_SPLIT.finditer(line):
            suffix = _parse_account_suffix(m.group(2))
            if not suffix:
                continue
            parsed = _parse_amount_token(m.group(1))
            if parsed is None or not _is_plausible_line_amount(m.group(1), parsed):
                continue
            raw_matches.append(
                (suffix, parsed, line_offset + m.start(), line_offset + m.end(), True, True)
            )

        for m in _RE_AMOUNT_BEFORE_LINEA_FULL.finditer(line):
            suffix = _parse_account_suffix(m.group(2))
            if not suffix:
                continue
            parsed = _parse_amount_token(m.group(1))
            if parsed is None or not _is_plausible_line_amount(m.group(1), parsed):
                continue
            raw_matches.append(
                (suffix, parsed, line_offset + m.start(), line_offset + m.end(), True, False)
            )

        for m in _RE_CODE_BEFORE_AMOUNT.finditer(line):
            suffix = _parse_account_suffix(m.group(1))
            if not suffix:
                continue
            parsed = _parse_amount_token(m.group(2))
            if parsed is None or not _is_plausible_line_amount(m.group(2), parsed):
                continue
            raw_matches.append(
                (suffix, parsed, line_offset + m.start(), line_offset + m.end(), False, False)
            )

        for m in _RE_ACCOUNT_AT_LINE_START.finditer(line):
            suffix = _parse_account_suffix(m.group(1))
            if not suffix:
                continue
            parsed = _parse_amount_token(m.group(2))
            if parsed is None or not _is_plausible_line_amount(m.group(2), parsed):
                continue
            raw_matches.append(
                (suffix, parsed, line_offset + m.start(), line_offset + m.end(), False, False)
            )

        for m in _RE_ACCOUNT_LINE_AMOUNT_AT_END.finditer(line):
            suffix = _parse_account_suffix(m.group(1))
            if not suffix:
                continue
            parsed = _parse_amount_token(m.group(2))
            if parsed is None or not _is_plausible_line_amount(m.group(2), parsed):
                continue
            raw_matches.append(
                (suffix, parsed, line_offset + m.start(), line_offset + m.end(), False, False)
            )

        line_offset += len(line) + 1

    suffix_totals, amounts_before, used_split, used_full = _merge_suffix_matches(raw_matches)
    totals, detected, _ret_suffix = _aggregate_suffix_totals(suffix_totals)

    if used_split and not used_full:
        parser_mode = "split_code_pypdf"
    elif used_full and not used_split:
        parser_mode = "full_code"
    elif used_split and used_full:
        parser_mode = "mixed"
    else:
        parser_mode = "prefix_table" if suffix_totals else ""

    return totals, amounts_before, parser_mode, detected, _ret_suffix


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
            f"(sin línea banco {BANK_BOGOTA_SUFFIX}/{BANK_BANCOLOMBIA_SUFFIX})"
        )
        return component_sum, (warning,)

    if capital > 0:
        return capital, (WARNING_BANK_INFERRED,)

    return 0.0, ()


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
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
    raw = text or ""
    totals, amounts_before_code, parser_mode, detected, retenciones_suffix = (
        _extract_amounts_by_account_code(raw)
    )
    retenciones = _resolve_retenciones(retenciones_suffix, raw)

    valor_pagado, parse_warnings = _infer_valor_pagado_cliente(totals)

    if valor_pagado <= 0:
        preview = _normalize_preview(raw)
        if detected:
            raise AccountingParseError(
                f"No se encontró recaudo bancario ({BANK_BOGOTA_SUFFIX} / {BANK_BANCOLOMBIA_SUFFIX}) "
                f"en el asiento; códigos con monto: {', '.join(detected)}.",
                error_code="MISSING_BANK_VALUE_BUT_HAS_ACCOUNTING_LINES",
                detected_codes=detected,
                amounts_before_code=amounts_before_code,
                text_preview=preview,
                parser_mode=parser_mode,
            )
        raise AccountingParseError(
            f"No se encontró recaudo bancario ({BANK_BOGOTA_SUFFIX} / {BANK_BANCOLOMBIA_SUFFIX}) "
            "en el asiento.",
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
        retenciones=retenciones,
        saldos_menores=totals.get(ACCOUNT_SALDOS_MENORES, 0.0),
        raw_text=raw,
        parse_warnings=parse_warnings,
        parser_mode=parser_mode,
        detected_codes=tuple(detected),
    )
