import asyncio
import base64
import io

import pandas as pd

from app.excel_parser import parse_excel_from_bytes
from app.main import parse_excel_json
from app.models import ParseExcelJsonBody


def _df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    return buffer.getvalue()


def test_parse_excel_reads_all_columns_and_rows():
    dataframe = pd.DataFrame(
        {
            "A": [1, 2, 3, 4],
            "B": ["x", "y", "z", "w"],
            "C": [" Ana ", "BRUNO", "ana", None],
        }
    )
    excel_bytes = _df_to_xlsx_bytes(dataframe)

    columns, rows = parse_excel_from_bytes(excel_bytes)

    assert columns == ["A", "B", "C"]
    assert len(rows) == 4
    assert rows[0]["C"] == " Ana "
    assert rows[1]["C"] == "BRUNO"
    assert rows[3]["C"] is None


def test_parse_excel_json_endpoint_accepts_base64_body():
    dataframe = pd.DataFrame({"A": [1, 2], "B": ["a", "b"]})
    excel_bytes = _df_to_xlsx_bytes(dataframe)
    b64 = base64.b64encode(excel_bytes).decode("ascii")

    async def _call() -> None:
        result = await parse_excel_json(
            ParseExcelJsonBody(file_name="test.xlsx", excel_base64=b64, run_id="r1")
        )
        assert result.status == "ok"
        assert result.file_name == "test.xlsx"
        assert result.run_id == "r1"
        assert result.row_count == 2
        assert result.columns == ["A", "B"]

    asyncio.run(_call())
