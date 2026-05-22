from datetime import datetime, date
import re
import unicodedata
from io import BytesIO

def _normalize_str(text: str) -> str:
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return text.strip().lower()

def parse_bank_date(raw_date: any, process_date: date) -> date:
    if isinstance(raw_date, datetime):
        return raw_date.date()
    if isinstance(raw_date, date):
        return raw_date
        
    if not isinstance(raw_date, str):
        raise ValueError("bank_date_parse_error")
        
    text = raw_date.strip()
    
    # Try text match like "30/12/2025" or "30-12-2025"
    m_full = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$", text)
    if m_full:
        return date(int(m_full.group(3)), int(m_full.group(2)), int(m_full.group(1)))
        
    # Try text match like "30-dic"
    m_short = re.match(r"^(\d{1,2})-(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)$", text.lower())
    if m_short:
        day = int(m_short.group(1))
        months = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
        month = months.index(m_short.group(2)) + 1
        return date(process_date.year, month, day)
        
    raise ValueError("bank_date_parse_error")

def parse_bank_amount(raw_amount: any) -> float:
    if isinstance(raw_amount, (int, float)):
        return float(raw_amount)
        
    if not isinstance(raw_amount, str):
        raise ValueError("bank_amount_parse_error")
        
    text = raw_amount.strip()
    if not text:
        raise ValueError("bank_amount_parse_error")
        
    # Eliminar $, espacios, y puntos de miles
    text = text.replace("$", "").replace(" ", "").replace(".", "")
    # Reemplazar coma de decimales por punto
    text = text.replace(",", ".")
    
    try:
        return float(text)
    except ValueError:
        raise ValueError("bank_amount_parse_error")

def extract_client_from_bank_row(concepto: str | None, transaccion: str | None) -> str:
    c = str(concepto).strip() if concepto else ""
    t = str(transaccion).strip() if transaccion else ""
    
    if c and c.lower() != "nan":
        return c
    if t and t.lower() != "nan":
        return t
        
    raise ValueError("customer_not_found")

def parse_statement_name(filename: str) -> tuple[date | None, str | None]:
    # Extracto YYYY-MM-DD CREDITO # NNN.pdf
    m_new = re.search(r"(?i)extracto\s+(\d{4}-\d{2}-\d{2})\s+cr[eé]dito\s*#\s*(\d+)", filename)
    if m_new:
        d = datetime.strptime(m_new.group(1), "%Y-%m-%d").date()
        return d, m_new.group(2)
        
    # Extracto Abril 20-03-2026 Obligacio # 265.pdf
    m_old = re.search(r"(?i)extracto\s+[a-z]+\s+(\d{2}-\d{2}-\d{4})\s+obligaci[oó]n?\s*#\s*(\d+)", filename)
    if m_old:
        d = datetime.strptime(m_old.group(1), "%d-%m-%Y").date()
        return d, m_old.group(2)
        
    return None, None

_FECHA_NEAR_WINDOW = 120


def _parse_limite_date_token(token: str) -> date | None:
    """
    Interpreta fechas de extracto: ISO yyyy-mm-dd o europeo dd/mm/yyyy y dd-mm-yyyy
    (día primero; nunca mm/dd/yyyy).
    """
    token = token.strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", token)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$", token)
    if not m:
        return None
    day, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if mo < 1 or mo > 12 or day < 1 or day > 31:
        return None
    try:
        return date(y, mo, day)
    except ValueError:
        return None


def _first_date_in_ordered_scan(snippet: str) -> date | None:
    """Primera fecha reconocible en orden de lectura izquierda-derecha."""
    for m in re.finditer(r"\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[/-]\d{1,2}[/-]\d{4}", snippet):
        cand = _parse_limite_date_token(m.group(0))
        if cand:
            return cand
    return None


def _extract_near_explicit_labels(pdf_text: str) -> date | None:
    """
    Prioriza texto inmediatamente posterior a etiquetas de vencimiento / fecha límite de pago.
    Evita confundir con desembolso u otros campos cuando la etiqueta explícita está presente.
    """
    # Orden: más específico primero; (?i) = mayúsculas/minúsculas; tildes opcionales en límite
    label_exprs = (
        r"fecha\s+l[ií]mite\s+de\s+pago",
        r"fecha\s+l[ií]mite\s+pago",
        r"fecha\s+de\s+vencimiento",
        r"\bvencimiento\b",
    )
    for expr in label_exprs:
        rx = re.compile(expr, re.IGNORECASE)
        for mat in rx.finditer(pdf_text):
            tail = pdf_text[mat.end() : mat.end() + _FECHA_NEAR_WINDOW]
            anchor_d = _first_date_in_ordered_scan(tail)
            if anchor_d is not None:
                return anchor_d
    return None


def _extract_fecha_limite_loose_fallback(pdf_text: str) -> date | None:
    """
    Si no hubo etiqueta reconocida, intenta líneas típicas de extracto sin anclar tanto.
    Solo si no hay mejor evidencia — evitar patrones (yyyy/mm/dd tras 'limité') ambiguos.
    """
    m = re.search(
        r"(?i)fecha\s+l[ií]mite\s+de\s+pago[^\d]{0,20}(\d{4}-\d{1,2}-\d{1,2})", pdf_text
    )
    if m:
        cand = _parse_limite_date_token(m.group(1))
        if cand:
            return cand
    m = re.search(
        r"(?i)fecha\s+l[ií]mite\s+de\s+pago[^\d]{0,20}(\d{1,2}[/-]\d{1,2}[/-]\d{4})",
        pdf_text,
    )
    if m:
        cand = _parse_limite_date_token(m.group(1))
        if cand:
            return cand
    for mat in re.finditer(r"(?i)fecha\s+l[ií]mite\b", pdf_text):
        tail = pdf_text[mat.end() : mat.end() + 72]
        if re.search(r"desembolso", tail[:40], re.IGNORECASE):
            continue
        d = _first_date_in_ordered_scan(tail)
        if d:
            return d
    return None


def extract_fecha_limite_pago_from_pdf_text(pdf_text: str) -> date | None:
    """
    Obtiene la fecha límite de pago desde el texto de un extracto HBI.
    Devuelve None si no hay coincidencia clara (no lanzar: el llamador excluye el PDF).

    Preferencia explícita: fechas cercanas a «Fecha límite de pago», «Fecha limite pago»,
    «Fecha de vencimiento», etc. Formatos yyyy-mm-dd, dd/mm/yyyy y dd-mm-yyyy (día/mes año).
    """
    if not pdf_text:
        return None
    d = _extract_near_explicit_labels(pdf_text)
    if d is not None:
        return d
    return _extract_fecha_limite_loose_fallback(pdf_text)


def extract_fecha_limite_pago_from_pdf(pdf_bytes: bytes) -> date | None:
    """Igual que extract_fecha_limite_pago_from_pdf_text pero leyendo bytes del PDF."""
    try:
        text = _read_pdf_text(pdf_bytes)
    except ValueError:
        return None
    return extract_fecha_limite_pago_from_pdf_text(text)


def extract_from_pdf_fallback(pdf_text: str) -> tuple[str | None, date | None, float | None]:
    cred, d, total = None, None, None
    
    m_cred = re.search(r"(?i)No\.?\s*Obligaci[oó]n\s+(\d+)", pdf_text)
    if m_cred:
        cred = m_cred.group(1)

    d = extract_fecha_limite_pago_from_pdf_text(pdf_text)
    m_total = re.search(r"(?i)TOTAL\s+A\s+PAGAR\s*\$?\s*([\d\.\,]+)", pdf_text)
    if m_total:
        try:
            val_str = m_total.group(1).replace(".", "").replace(",", ".")
            total = float(val_str)
        except ValueError:
            pass
            
    return cred, d, total


def _is_amortization_excel_filename(filename: str) -> bool:
    """Tabla operativa = solo Excel (no PDFs informativos con 'tabla' en el nombre)."""
    if not filename or filename.startswith("~$"):
        return False
    lower = filename.lower()
    return lower.endswith(".xlsx") or lower.endswith(".xlsm") or lower.endswith(".xls")


def filter_amortization_excel_filenames(files: list[str]) -> list[str]:
    return [f for f in files if f and _is_amortization_excel_filename(f)]


def extract_credit_id_from_extract_pdf_text(pdf_text: str) -> str | None:
    """
    Número de obligación/crédito legible en texto de extracto (modo cliente plano / refuerzo).
    Ej.: «No, Obligación GB-1022-082» → «82» (último segmento numérico significativo).
    No sustituye la carpeta de crédito cuando esa carpeta existe explícitamente en SharePoint.
    """
    if not pdf_text or not pdf_text.strip():
        return None
    # Patrones explícitos primero (evitar NIT/cuentas largas sin contexto «obligación»)
    m = re.search(r"(?is)(?:obligaci[oó]n)\s*#\s*(\d{1,8})\b", pdf_text)
    if m:
        return str(int(m.group(1))) if m.group(1).isdigit() else m.group(1)
    m = re.search(r"(?is)\b(?:obligaci[oó]n)\s+(\d{1,8})\b(?!\s*\.?\d)", pdf_text)
    if m:
        return str(int(m.group(1))) if m.group(1).isdigit() else m.group(1)
    # No. / No, Obligacion GB-1022-082
    m = re.search(
        r"(?is)(?:no\.?\s*,?\s*)(?:obligaci[oó]n)\s+([a-z]{1,6}(?:-\d+)+)",
        pdf_text,
    )
    if m:
        parts = m.group(1).upper().split("-")
        tail = parts[-1]
        if tail.isdigit():
            return str(int(tail))
        return tail or None
    m = re.search(r"(?is)(?:obligaci[oó]n)\s+([a-z]{1,6}(?:-\d+)+)", pdf_text)
    if m:
        parts = m.group(1).upper().split("-")
        tail = parts[-1]
        if tail.isdigit():
            return str(int(tail))
        return tail or None
    return None


def extract_credit_id_from_extract_pdf(pdf_bytes: bytes) -> str | None:
    text: str | None = None
    try:
        text = _read_pdf_text(pdf_bytes)
    except Exception:
        raw = pdf_bytes.decode(errors="ignore")
        if raw.startswith("PDF_MOCK:"):
            text = raw
        else:
            return None
    return extract_credit_id_from_extract_pdf_text(text or "")


def find_best_amortization_table(files: list[str], client_name: str, credit_id: str) -> str:
    candidates = []
    
    c_norm = _normalize_str(client_name)
    
    for f in files:
        if f.startswith("~$"):
            continue
        if not _is_amortization_excel_filename(f):
            continue
            
        f_norm = _normalize_str(f)
        score = 0
        
        has_tabla = "tabla" in f_norm
        has_amort = "amortizacion" in f_norm
        has_modelo = "modelo" in f_norm
        has_extracto = "extracto" in f_norm
        
        if has_tabla and has_amort:
            score += 10
        elif has_modelo and has_extracto:
            score += 5
            
        if score == 0:
            continue
            
        if credit_id in f_norm:
            score += 5
            
        if c_norm and c_norm in f_norm:
            score += 3
            
        candidates.append((score, f))
        
    if not candidates:
        raise ValueError("amortization_table_not_found")
        
    # Sort by score descending
    candidates.sort(key=lambda x: x[0], reverse=True)
    
    best_score = candidates[0][0]
    best_files = [f for s, f in candidates if s == best_score]
    
    if len(best_files) > 1:
        raise ValueError("amortization_table_ambiguous")
        
    return best_files[0]


def _read_pdf_text(pdf_bytes: bytes) -> str:
    """Extrae texto plano de un PDF. Lanza ValueError si el PDF no tiene texto."""
    try:
        from pypdf import PdfReader
    except ImportError:
        from PyPDF2 import PdfReader  # type: ignore[no-redef]
    reader = PdfReader(BytesIO(pdf_bytes))
    text_parts: list[str] = []
    for page in reader.pages:
        t = page.extract_text() or ""
        if t:
            text_parts.append(t)
    text = "\n".join(text_parts)
    if not text.strip():
        raise ValueError("pdf_no_text")
    return text


def extract_total_a_pagar_from_pdf(pdf_bytes: bytes) -> float:
    """
    Extrae el valor TOTAL A PAGAR de un PDF de extracto HBI.

    Soporta formatos latinos:
      - 18.826.879
      - 6.500.739
      - 1.234.567,89

    Lanza ValueError('extract_amount_not_found') si no puede parsear el monto.
    """
    text = _read_pdf_text(pdf_bytes)
    patterns = [
        r"TOTAL\s*A\s*PAGAR\s*[:\-\s]*\$?\s*([\d\.\,]+)",
        r"TOTAL\s*A\s*PAGAR\s*\$?\s*([\d\.\,]+)",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            raw = m.group(1).strip()
            # Formato latino: puntos = separadores de miles, coma = decimal
            # Ej: "18.826.879" -> 18826879, "1.234.567,89" -> 1234567.89
            if "," in raw:
                # Tiene decimal con coma
                raw = raw.replace(".", "").replace(",", ".")
            else:
                # Solo puntos como miles
                raw = raw.replace(".", "")
            try:
                return float(raw)
            except ValueError:
                pass
    raise ValueError("extract_amount_not_found")
