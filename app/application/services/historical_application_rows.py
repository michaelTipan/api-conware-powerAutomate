"""
Lectura compartida de filas PAGO/ABONO validadas en el histórico (cartera_validada).

Notify y Merge deben usar estas funciones para no divergir en reglas de inclusión.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

from app.application.services.review_schema import (
    DistribucionAbonosCols,
    DistribucionCols,
    ReviewSheets,
    TipoAplicacion,
    normalize_credito_digits,
    require_validar_abono_value,
)
from app.application.use_cases.send_validar_extractos_notification import (
    _accent_fold_upper,
    _excel_cell_display,
    _find_distribucion_header_row,
    _find_distribucion_sheet,
    _get_col_distrib,
    _norm_key,
    _norm_sheet_name,
    distrib_row_included_for_validar_extractos,
)


def _get_col_abono(header_map: dict[str, int], *candidates: str) -> int | None:
    for c in candidates:
        key = _accent_fold_upper(c)
        if key in header_map:
            return header_map[key]
    return None


def _find_distribucion_abonos_sheet(wb: Any) -> Any | None:
    target = _norm_sheet_name(ReviewSheets.DISTRIBUCION_ABONOS)
    for ws in wb.worksheets:
        if _norm_sheet_name(ws.title) == target:
            return ws
    return None


def _find_abonos_header_row(ws: Any) -> tuple[int, dict[str, int]]:
    marker = _accent_fold_upper(DistribucionAbonosCols.ID_PAGO)
    max_col = ws.max_column or 1
    max_scan = min(ws.max_row or 1, 50)
    for row_idx in range(1, max_scan + 1):
        raw_vals = [ws.cell(row=row_idx, column=c).value for c in range(1, max_col + 1)]
        cells = [_norm_key(str(v)) if v is not None and str(v).strip() else "" for v in raw_vals]
        if not any(cells):
            continue
        folded = [_accent_fold_upper(c) if c else "" for c in cells]
        if marker not in folded:
            continue
        header_map: dict[str, int] = {}
        for col, name in enumerate(cells, start=1):
            if name:
                header_map[_accent_fold_upper(name)] = col
        return row_idx, header_map
    raise ValueError("missing_distribucion_abonos_headers")


def _coerce_historical_date(value: Any) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _coerce_historical_amount(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text or text.startswith("="):
        return None
    try:
        return float(text)
    except ValueError:
        try:
            return float(text.replace(".", "").replace(",", "."))
        except ValueError:
            return None


def abono_row_included_for_notify_merge(
    ws: Any,
    row: int,
    header_map: dict[str, int],
) -> bool:
    col_va = _get_col_abono(
        header_map,
        DistribucionAbonosCols.VALIDAR_ABONO,
        "Validar abono",
        "VALIDAR ABONO",
    )
    if col_va is None:
        return False
    try:
        norm_va = require_validar_abono_value(ws.cell(row=row, column=col_va).value)
    except ValueError:
        return False
    if norm_va != "SI":
        return False

    col_tipo = _get_col_abono(header_map, DistribucionAbonosCols.TIPO_APLICACION, "TipoAplicacion")
    if col_tipo is not None:
        tipo = str(ws.cell(row=row, column=col_tipo).value or "").strip().upper()
        if tipo and tipo != TipoAplicacion.ABONO.value:
            return False

    col_req = _get_col_abono(header_map, DistribucionAbonosCols.REQUIERE_EXTRACTO, "RequiereExtracto")
    if col_req is not None:
        req = str(ws.cell(row=row, column=col_req).value or "").strip().upper()
        if req and req != "NO":
            return False

    col_ra = _get_col_abono(
        header_map,
        DistribucionAbonosCols.RUTA_ASIENTOS_CONTABLES,
        "RutaAsientosContables",
        "RUTA_ASIENTOS_CONTABLES",
    )
    if col_ra is None:
        return False
    ruta_as = _excel_cell_display(ws.cell(row=row, column=col_ra).value).strip()
    if not ruta_as:
        return False

    col_cred = _get_col_abono(header_map, DistribucionAbonosCols.CREDITO, "Crédito", "Credito", "CREDITO")
    if col_cred is not None:
        cred = _excel_cell_display(ws.cell(row=row, column=col_cred).value).strip()
        if not cred:
            return False

    return True


def read_validated_payment_rows(
    wb: Any,
    *,
    legacy_estado_token: str,
) -> list[dict[str, Any]]:
    ws = _find_distribucion_sheet(wb)
    h_row, header_map = _find_distribucion_header_row(ws)
    col_estado = _get_col_distrib(
        header_map,
        DistribucionCols.ESTADO_PAGO,
        "Estado pago",
        "Estado",
        "Estado línea",
        "Estado linea",
        "ESTADO LINEA",
    )
    col_ruta = _get_col_distrib(header_map, DistribucionCols.RUTA, "RUTA", "Rutas", "RUTAS")
    col_ruta_asientos = _get_col_distrib(
        header_map,
        DistribucionCols.RUTA_ASIENTOS_CONTABLES,
        "RUTA_ASIENTOS_CONTABLES",
        "Ruta asientos contables",
        "RutaAsientosContables",
    )
    col_id = _get_col_distrib(
        header_map,
        DistribucionCols.ID_PAGO,
        "ID pago",
        "ID PAGO",
        "Id pago",
        "ID_PAGO",
        "Id Pago",
    )
    col_cliente = _get_col_distrib(
        header_map,
        DistribucionCols.CLIENTE,
        "CLIENTE",
        "Nombre Cliente",
        "Nombre cliente",
    )
    col_credito = _get_col_distrib(
        header_map,
        DistribucionCols.CREDITO,
        "Credito",
        "CREDITO",
        "CRÉDITO",
    )
    col_monto = _get_col_distrib(header_map, DistribucionCols.MONTO_BANCO, "Monto banco", "MONTO BANCO")
    col_fecha = _get_col_distrib(header_map, DistribucionCols.FECHA_BANCO, "Fecha banco", "FECHA BANCO")
    col_vp = _get_col_distrib(header_map, DistribucionCols.VALIDAR_PAGO, "Validar pago", "VALIDAR PAGO")

    if col_estado is None and col_vp is None:
        raise ValueError("missing_distribucion_status_column")
    if col_ruta is None:
        raise ValueError("missing_distribucion_route_column")

    rows: list[dict[str, Any]] = []
    last = ws.max_row or h_row
    for r in range(h_row + 1, last + 1):
        if not distrib_row_included_for_validar_extractos(ws, r, header_map, legacy_estado_token):
            continue
        id_raw = ws.cell(row=r, column=col_id).value if col_id else None
        id_str = _excel_cell_display(id_raw).strip() if col_id else ""
        if not id_str:
            id_str = f"__sin_id_fila_{r}__"
        cred_raw = ws.cell(row=r, column=col_credito).value if col_credito else None
        cred_digits = normalize_credito_digits(cred_raw) if cred_raw is not None else ""
        if not cred_digits and cred_raw is not None:
            cred_digits = re.sub(r"\D", "", str(cred_raw).strip())
        cliente = ""
        if col_cliente:
            cliente = _excel_cell_display(ws.cell(row=r, column=col_cliente).value).strip()
        monto = _coerce_historical_amount(ws.cell(row=r, column=col_monto).value) if col_monto else None
        fecha = _coerce_historical_date(ws.cell(row=r, column=col_fecha).value) if col_fecha else None
        rows.append(
            {
                "tipo_aplicacion": TipoAplicacion.PAGO.value,
                "requiere_extracto": True,
                "excel_row": r,
                "id_pago": id_str,
                "cliente": cliente,
                "credito_raw": cred_raw,
                "credito_label": _excel_cell_display(cred_raw).strip() if cred_raw is not None else "",
                "credito_digits": cred_digits,
                "monto_banco": monto,
                "fecha_banco": fecha,
                "ruta_cell": ws.cell(row=r, column=col_ruta).value if col_ruta else None,
                "ruta_asientos_cell": (
                    ws.cell(row=r, column=col_ruta_asientos).value if col_ruta_asientos else None
                ),
            }
        )
    return rows


def read_validated_abono_rows(wb: Any) -> list[dict[str, Any]]:
    ws = _find_distribucion_abonos_sheet(wb)
    if ws is None:
        return []
    try:
        h_row, header_map = _find_abonos_header_row(ws)
    except ValueError:
        return []

    col_id = _get_col_abono(
        header_map,
        DistribucionAbonosCols.ID_PAGO,
        "ID pago",
        "ID PAGO",
        "Id Pago",
    )
    col_cliente = _get_col_abono(header_map, DistribucionAbonosCols.CLIENTE, "CLIENTE", "Cliente")
    col_credito = _get_col_abono(header_map, DistribucionAbonosCols.CREDITO, "Crédito", "Credito", "CREDITO")
    col_monto = _get_col_abono(header_map, DistribucionAbonosCols.MONTO_BANCO, "Monto banco", "MONTO BANCO")
    col_fecha = _get_col_abono(header_map, DistribucionAbonosCols.FECHA_BANCO, "Fecha banco", "FECHA BANCO")
    col_ra = _get_col_abono(
        header_map,
        DistribucionAbonosCols.RUTA_ASIENTOS_CONTABLES,
        "RutaAsientosContables",
        "RUTA_ASIENTOS_CONTABLES",
    )
    col_cred_norm = _get_col_abono(
        header_map,
        DistribucionAbonosCols.CREDITO_NORMALIZADO,
        "CreditoNormalizado",
        "CREDITO_NORMALIZADO",
    )
    col_obs = _get_col_abono(header_map, DistribucionAbonosCols.OBSERVACION, "Observación", "Observacion")

    rows: list[dict[str, Any]] = []
    last = ws.max_row or h_row
    for r in range(h_row + 1, last + 1):
        if not abono_row_included_for_notify_merge(ws, r, header_map):
            continue
        id_str = _excel_cell_display(ws.cell(row=r, column=col_id).value).strip() if col_id else ""
        if not id_str:
            id_str = f"__sin_id_fila_{r}__"
        cred_raw = ws.cell(row=r, column=col_credito).value if col_credito else None
        cred_norm = ""
        if col_cred_norm:
            cred_norm = _excel_cell_display(ws.cell(row=r, column=col_cred_norm).value).strip()
        if not cred_norm:
            cred_norm = normalize_credito_digits(cred_raw) or (
                re.sub(r"\D", "", str(cred_raw).strip()) if cred_raw is not None else ""
            )
        cliente = _excel_cell_display(ws.cell(row=r, column=col_cliente).value).strip() if col_cliente else ""
        rows.append(
            {
                "tipo_aplicacion": TipoAplicacion.ABONO.value,
                "requiere_extracto": False,
                "excel_row": r,
                "id_pago": id_str,
                "cliente": cliente,
                "credito_raw": cred_raw,
                "credito_label": _excel_cell_display(cred_raw).strip() if cred_raw is not None else "",
                "credito_digits": cred_norm,
                "monto_banco": _coerce_historical_amount(ws.cell(row=r, column=col_monto).value) if col_monto else None,
                "fecha_banco": _coerce_historical_date(ws.cell(row=r, column=col_fecha).value) if col_fecha else None,
                "observacion": _excel_cell_display(ws.cell(row=r, column=col_obs).value).strip() if col_obs else "",
                "ruta_asientos_cell": ws.cell(row=r, column=col_ra).value if col_ra else None,
            }
        )
    return rows


@dataclass(frozen=True)
class AbonoEmailGroup:
    id_pago: str
    cliente: str
    monto_banco: float | None
    fecha_banco: date | None
    creditos_seleccionados: tuple[str, ...]


def group_abono_rows_for_email(abono_rows: list[dict[str, Any]]) -> list[AbonoEmailGroup]:
    by_id: dict[str, list[dict[str, Any]]] = {}
    for row in abono_rows:
        id_pago = str(row.get("id_pago") or "").strip()
        if not id_pago:
            continue
        by_id.setdefault(id_pago, []).append(row)

    groups: list[AbonoEmailGroup] = []
    for id_pago in sorted(by_id.keys()):
        members = by_id[id_pago]
        ref = members[0]
        cliente = str(ref.get("cliente") or "").strip()
        monto = ref.get("monto_banco")
        fecha = ref.get("fecha_banco")
        for row in members[1:]:
            if str(row.get("cliente") or "").strip() != cliente:
                raise ValueError(f"abono_group_inconsistent|id_pago={id_pago}")
            if row.get("monto_banco") != monto:
                raise ValueError(f"abono_group_inconsistent|id_pago={id_pago}")
            if row.get("fecha_banco") != fecha:
                raise ValueError(f"abono_group_inconsistent|id_pago={id_pago}")
        creditos: list[str] = []
        seen: set[str] = set()
        for row in sorted(members, key=lambda x: str(x.get("credito_digits") or x.get("credito_label") or "")):
            c = str(row.get("credito_digits") or row.get("credito_label") or "").strip()
            if c and c not in seen:
                seen.add(c)
                creditos.append(c)
        groups.append(
            AbonoEmailGroup(
                id_pago=id_pago,
                cliente=cliente,
                monto_banco=monto if isinstance(monto, (int, float)) else None,
                fecha_banco=fecha if isinstance(fecha, date) else None,
                creditos_seleccionados=tuple(creditos),
            )
        )
    return groups


def group_rows_by_id_pago(rows: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    groups: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        id_pago = str(row.get("id_pago") or "").strip() or f"__sin_id_fila_{row.get('excel_row')}__"
        groups.setdefault(id_pago, []).append(row)
    return groups


def detect_application_type_group_conflict(
    payment_groups: dict[str, list[dict[str, Any]]],
    abono_groups: dict[str, list[dict[str, Any]]],
) -> None:
    overlap = set(payment_groups.keys()) & set(abono_groups.keys())
    if overlap:
        ids = ",".join(sorted(overlap))
        raise ValueError(f"application_type_group_conflict|id_pagos={ids}")
