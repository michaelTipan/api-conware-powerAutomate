import base64
import io

import pandas as pd

from app.excel_parser import parse_excel_names


def _to_excel_b64(df: pd.DataFrame) -> str:
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def test_parse_excel_names_column_c_and_duplicates():
    dataframe = pd.DataFrame(
        {
            "A": [1, 2, 3, 4],
            "B": ["x", "y", "z", "w"],
            "C": [" Ana ", "BRUNO", "ana", None],
        }
    )
    excel_b64 = _to_excel_b64(dataframe)

    unique_names, duplicates, total_rows, in_order = parse_excel_names(excel_b64, "C")

    assert unique_names == ["ana", "bruno"]
    assert duplicates == ["ana"]
    assert total_rows == 4
    assert in_order == ["ana", "bruno", "ana"]
