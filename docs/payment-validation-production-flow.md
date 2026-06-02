# Flujo productivo — validación de pagos (Render)

Runtime: **Render** (`uvicorn app.main:app`). Control operativo: **un Excel por banco** en `00 CONTROL`.

## Controles oficiales

- `control_proceso_validacion_pagos_banco_bogota.xlsx`
- `control_proceso_validacion_pagos_banco_bancolombia.xlsx`

Ruta base de control: configurable vía `GRAPH_PAYMENT_VALIDATION_CONTROL_PATH` o `PAYMENT_VALIDATION_BASE_FOLDER` + `PAYMENT_VALIDATION_CONTROL_FOLDER` (ver `docs/render-env-variables.md`). Valor por defecto en código: `…/02 COMWARE - VALIDACION PAGOS/00 CONTROL`.

El archivo `control_merge_pdfs.xlsx` **no** forma parte del flujo. El endpoint antiguo `POST /graph/sharepoint/validate-payment-report` fue retirado; use Generate/Finalize.

## Orden operativo

| Paso | Método | Ruta | Job status |
|------|--------|------|------------|
| 0 Setup control | POST | `/graph/sharepoint/payment-validation/setup/merge-control-workbook` | — |
| 1 Generate | POST | `/graph/sharepoint/payment-validation/generate/queue` | `GET .../payment-validation/jobs/{job_id}` |
| 2 Finalize | POST | `/graph/sharepoint/payment-validation/finalize/queue` | `GET .../payment-validation/jobs/{job_id}` |
| 3 Notify | POST | `/graph/sharepoint/notify-validar-extractos-email` | `GET .../notify-validar-extractos-email/jobs/{job_id}` |
| 4 Merge | POST | `/graph/sharepoint/merge-composite-validado-pdfs` | `GET .../merge-composite-validado-pdfs/jobs/{job_id}` |
| 5 Dry-run | POST | `/graph/sharepoint/payment-validation/amortization/dry-run/queue` | `GET .../payment-validation/jobs/{job_id}` |
| 6 Apply | POST | `/graph/sharepoint/payment-validation/amortization/apply/queue` | `GET .../payment-validation/jobs/{job_id}` |

Setup opcional (una vez o reparación): `POST .../setup/payment-followup-workbooks`, `POST .../setup/ibr-workbook`.

## Body típico

- **Generate / Finalize:** `bank_code` (`banco_bogota` | `banco_bancolombia`); Generate sin `bank_code` usa Bogotá por defecto.
- **Notify / Merge / Dry-run / Apply:** body vacío `{}` si un solo banco está listo; si no, `bank_code` explícito.
- Overrides manuales (pruebas): `historical_file_path`, `merge_manifest_path`, `report_date_iso`.

## Endpoints de operación (no son pasos del flujo diario)

- `GET /health`
- `POST /graph/sharepoint/ensure-asientos-contables-folders` (+ job GET)
- Utilidades Graph/SharePoint bajo `/graph/sharepoint/*` (resolve-env, item-content, etc.)
- `POST /parse-excel` — flujo Power Automate de parseo de columna (independiente del flujo de validación de pagos)

## OpenAPI

`GET /docs` lista solo rutas registradas en la app. Tras el retiro de `validate-payment-report`, no debe aparecer ese path.
