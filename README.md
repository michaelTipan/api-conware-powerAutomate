# Excel Parser API

Servicio Python que recibe un Excel en base64 (desde Power Automate u otro cliente), lee una columna configurable y devuelve los valores extraidos y metricas. Disenado para invocarse por HTTP.

## Stack
- FastAPI
- Pandas + OpenPyXL
- Render (deploy)

## Ejecucion local
1. Crear entorno virtual e instalar: `pip install -r requirements.txt`
2. Iniciar API: `uvicorn app.main:app --reload`

## Endpoint principal
- `POST /parse-excel`
- Entrada JSON:
  - `file_name`
  - `excel_base64`
  - `run_id` (opcional)
  - `column_letter` (opcional, default `C`)

## Respuesta
- `unique_values`: valores unicos (orden de primera aparicion)
- `values_in_row_order`: todas las celdas no vacias de la columna, en orden de fila
- `duplicates`: valores repetidos (tras normalizar)
- Conteos: `total_rows`, `non_empty_count`, `unique_count`, `duplicate_count`
- `elapsed_ms`

## Alias
- `POST /validate-excel` — mismo comportamiento que `/parse-excel` (obsoleto).

## Documentacion adicional
- Flujo Power Automate: `docs/power-automate-flow.md`
- Deploy en Render: `docs/deploy-render.md`
