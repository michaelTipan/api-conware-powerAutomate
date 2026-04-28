import io
import json

import pandas as pd


def parse_excel_from_bytes(excel_bytes: bytes) -> tuple[list[str], list[dict[str, object]]]:
    """
    Lee la primera hoja del Excel y devuelve todas las columnas y filas como diccionarios.

    Los valores nulos en celdas se serializan como null en JSON (vía to_json).
    """
    dataframe = pd.read_excel(io.BytesIO(excel_bytes), engine="openpyxl")
    columns = [str(c) for c in dataframe.columns]
    rows: list[dict[str, object]] = json.loads(
        dataframe.to_json(orient="records", date_format="iso", default_handler=str)
    )
    return columns, rows
