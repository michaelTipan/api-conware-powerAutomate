import time

from app.domain.services.excel_column import parse_excel_names
from app.models import ParseExcelRequest, ParseExcelResponse


def run_parse_excel(payload: ParseExcelRequest) -> ParseExcelResponse:
    started_at = time.perf_counter()
    unique_names, duplicates, total_rows, values_in_row_order = parse_excel_names(
        payload.excel_base64,
        payload.column_letter,
    )
    elapsed_ms = int((time.perf_counter() - started_at) * 1000)
    return ParseExcelResponse(
        status="ok",
        run_id=payload.run_id,
        file_name=payload.file_name,
        column_letter=payload.column_letter.upper(),
        total_rows=total_rows,
        non_empty_count=len(values_in_row_order),
        unique_count=len(unique_names),
        duplicate_count=len(duplicates),
        unique_values=unique_names,
        values_in_row_order=values_in_row_order,
        duplicates=duplicates,
        elapsed_ms=elapsed_ms,
    )
