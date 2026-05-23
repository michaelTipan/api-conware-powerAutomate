from pydantic import BaseModel, Field


class ParseExcelRequest(BaseModel):
    file_name: str
    excel_base64: str
    run_id: str = ""
    column_letter: str = "C"


class ParseExcelResponse(BaseModel):
    status: str
    run_id: str
    file_name: str
    column_letter: str
    total_rows: int
    non_empty_count: int
    unique_count: int
    duplicate_count: int
    unique_values: list[str] = Field(default_factory=list)
    values_in_row_order: list[str] = Field(default_factory=list)
    duplicates: list[str] = Field(default_factory=list)
    elapsed_ms: int = 0


class GraphUploadRequest(BaseModel):
    content_base64: str


class NotifyValidarExtractosRequest(BaseModel):
    """
    `historical_file_path`: ruta relativa al root del drive de validación (la misma que
    `result.historical_file_path` de Finalize). No usar `webUrl` ni URL pública.

    Obligatorio para ejecutar el job: si falta, está vacío o es solo espacios, el job falla
    con `missing_historical_file_path` (Power Automate debe enviar siempre el path devuelto por Finalize).

    `to` / `cc`: overrides opcionales respecto a CORREOS.xlsx (columnas EMISOR / RECEPTORES).
    """

    historical_file_path: str | None = Field(
        default=None,
        description="Ruta relativa al root del drive (igual que Finalize). Obligatorio para el job.",
    )
    to: str | None = None
    cc: str | None = None


class MergeCompositeValidadoRequest(BaseModel):
    """Opciones para POST merge-composite-validado-pdfs. Body opcional (compatibilidad PA)."""

    force_rebuild: bool = Field(
        default=False,
        description=(
            "Si true, regenera el PDF consolidado aunque ya exista en la carpeta de salida "
            "(útil para corregir consolidados con extractos duplicados)."
        ),
    )
