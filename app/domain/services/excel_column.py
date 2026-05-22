import base64
import io

import pandas as pd


def _normalize_name(value: str) -> str:
    return " ".join(value.strip().lower().split())


def parse_excel_names(
    excel_b64: str,
    column_letter: str = "C",
) -> tuple[list[str], list[str], int, list[str]]:
    """
    Lee la columna indicada del Excel (base64).

    Retorna:
    - unique_names: primeras apariciones, sin repetir
    - duplicates: valores que aparecieron más de una vez (normalizado)
    - total_rows: filas de datos en esa columna (incluye vacías)
    - values_in_row_order: solo celdas no vacías, en orden de fila, normalizadas
    """
    raw_bytes = base64.b64decode(excel_b64)
    dataframe = pd.read_excel(io.BytesIO(raw_bytes), engine="openpyxl")

    column_idx = ord(column_letter.upper()) - ord("A")
    if column_idx < 0 or column_idx >= len(dataframe.columns):
        raise ValueError(f"Column {column_letter} does not exist in the Excel file.")

    series = dataframe.iloc[:, column_idx]
    total_rows = int(series.shape[0])

    values_in_row_order: list[str] = []
    for cell in series.tolist():
        if pd.isna(cell):
            continue
        value = str(cell).strip()
        if not value or value.lower() == "nan":
            continue
        values_in_row_order.append(_normalize_name(value))

    seen: set[str] = set()
    duplicates: set[str] = set()
    unique_names: list[str] = []
    for name in values_in_row_order:
        if name in seen:
            duplicates.add(name)
            continue
        seen.add(name)
        unique_names.append(name)

    return unique_names, sorted(duplicates), total_rows, values_in_row_order
