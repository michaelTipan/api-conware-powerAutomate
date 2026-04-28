from typing import Any

from pydantic import BaseModel, Field, field_validator


class ParseExcelJsonBody(BaseModel):
    """Cuerpo JSON para clientes que no pueden enviar multipart (p. ej. HTTP estándar en Power Automate)."""

    file_name: str = "upload.xlsx"
    excel_base64: str
    run_id: str = ""

    @field_validator("excel_base64")
    @classmethod
    def strip_whitespace(cls, v: str) -> str:
        return v.strip()


class ParseExcelResponse(BaseModel):
    status: str
    run_id: str
    file_name: str
    row_count: int
    columns: list[str] = Field(default_factory=list)
    rows: list[dict[str, Any]] = Field(default_factory=list)
    elapsed_ms: int = 0
