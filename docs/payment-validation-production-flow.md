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

### Dry-run y ABONO (Fase 4 + Fase 5)

El dry-run lee el manifest extendido de Merge y distingue **PAGO** y **ABONO**:

- **PAGO:** comportamiento histórico (extracto, fecha límite, `due_date_row`, IBR, planning de filas).
- **ABONO:** preflight documental, **cuadre financiero por `ID Pago`** y planificación de fila de aplicación:
  - exige un asiento contable por cada crédito seleccionado;
  - **no** exige extracto (`extracto_pdf_paths` puede estar vacío);
  - **no** busca fecha límite, `due_date_row` ni IBR (`schedule_resolution_status=NOT_REQUIRED`);
  - suma `valor_pagado_cliente` una vez por PDF de asiento y lo compara con `monto_banco` (tolerancia `0.02`);
  - si el cuadre pasa, planifica `application_row` como la **siguiente fila libre** de Aplicación de Pagos (`NEXT_AVAILABLE_PAYMENT_ROW`);
  - la fecha de la fila proviene del asiento (`fecha_asiento`) o, en su defecto, `fecha_banco` del grupo;
  - los valores escritos provienen del asiento contable (misma semántica que PAGO en Aplicación de Pagos);
  - bloquea el grupo si falta un asiento, hay PDF duplicado, el cuadre no pasa o la tabla no es escribible.

Manifest legacy (sin `tipo_aplicacion`) se interpreta como **PAGO**.

**Apply ABONO (Fase 5):** escribe la fila planificada, extiende fórmulas O:P, registra `_AUTOMATION_LOG` y mueve asientos a `PROCESADOS/` tras verificación. **No modifica IBR** ni celdas de cuota contractual. Reintento con la misma huella PDF → `SKIPPED_IDEMPOTENT` (sin fila duplicada).

**Apply mixto:** si un ABONO del mismo `ProcessKey` está bloqueado, Apply no escribe ninguna tabla (fail-closed global). No hay transacción distribuida en SharePoint: un apply parcial deja `AMORTIZACION_PARCIAL` y el retry completa solo tablas pendientes vía idempotencia fina.

Tras **Apply** exitoso (tabla verificada y eventos `APPLIED`):

- Las tablas de amortización quedan actualizadas en SharePoint.
- Cada PDF de asiento usado se mueve a `PROCESADOS/` **dentro** de la carpeta del crédito (`ASIENTOS CONTABLES CRED {n}/PROCESADOS/`), con nombre trazable (`asiento_{fecha}_{bank_code}_credito-{n}_pago-{id_pago}[_evento-{i}].pdf`).
- Los asientos no usados en ese apply permanecen en la carpeta original del crédito.
- El resultado del job incluye `accounting_pdfs_moves` y contadores (`accounting_pdfs_moved_count`, etc.).

Reintento de **Apply** con el mismo proceso (`EstadoProceso=AMORTIZACION_APLICADA` y `ApplyIdempotencyKey` igual a `ProcessKey` en control):

- Respuesta idempotente (`already_applied=true`, `apply_wrote_changes=false`) **antes** de dry-run/preflight.
- No descarga PDFs de asiento en la ruta original (aunque ya estén en `PROCESADOS/`).
- No mueve asientos ni reescribe tablas.

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
