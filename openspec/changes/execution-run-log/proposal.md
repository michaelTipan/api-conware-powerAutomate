# Proposal: Execution Run Log — trazabilidad consolidada por corrida

## Intent

Un **único JSON por ejecución completa** en SharePoint `04 LOGS`, sin subcarpetas ni ZIP, que permita a soporte técnico diagnosticar cualquier corrida en producción (múltiples por día, misma banco) sin Render ni historial PA. **100 % aditivo**: contratos HTTP, estados de control, manifest canónico y lógica financiera intactos cuando `EXECUTION_RUN_LOG_ENABLED=false`.

## Scope

### In Scope
- Servicio central `execution_run_log.py` + sanitización.
- Archivo: `execution_log_{bank_code}_{run_timestamp_utc}_{execution_id_short}.json`.
- `execution_id` al encolar Generate; `process_id` separado (nullable hasta éxito Generate).
- Columnas control **aditivas**: `ExecutionId`, `ExecutionLogPath`.
- Hooks en routers (QUEUED) y use cases (STARTED/terminal).
- Concurrencia: eTag + reintento; dedup por `event_id`; `revision`.
- Snapshot manifest resumido **dentro** del JSON.
- Feature flag default `false`; campos opcionales en `result`.
- Tests listados en §20.

### Out of Scope
- Subcarpetas `events/`, ZIP, retención automática, cambio de manifest canónico, persistencia de jobs, OpenTelemetry, cambios Power Automate.

## Capabilities

### New Capabilities
- `execution-run-log`: auditoría JSON consolidada por `execution_id` en `04 LOGS`.

### Modified Capabilities
- `payment-validation-process-control`: columnas opcionales `ExecutionId`, `ExecutionLogPath` (setup idempotente; sin cambio de semántica operativa).

---

## Reporte de propuesta (25 puntos)

### 1. Auditoría del flujo actual de `ProcessId`

| Aspecto | Estado actual |
|---------|---------------|
| Creación | UUID en `generate_payment_validation` **línea ~3360**, tras validaciones e idempotencia |
| Persistencia | Columna `ProcessId` en control fila 2 al **éxito** de Generate |
| Lectura downstream | `ProcessControlSnapshot` **no parsea** `ProcessId`, job IDs ni `LastCompletedStep` |
| Finalize | **No referencia** `ProcessId` en código |
| Idempotencia | `already_generated` retorna `process_id: ""` sin nuevo UUID |
| Correlación hoy | `process_key` (banco+fecha reporte), rutas en control, `job_id` por paso (distinto cada endpoint) |

**Conclusión:** `ProcessId` identifica el proceso de negocio de una corrida exitosa de Generate, pero **no** cubre fallos tempranos ni correlación antes de Generate; no es leído por pasos posteriores.

### 2. Momento exacto de creación de `execution_id`

En `POST /generate/queue`, **inmediatamente después** de validar `bank_code` y `process_date`, **antes** de `background_tasks.add_task`:

1. `execution_id = uuid4()`
2. `run_timestamp_utc` compacto (`YYYYMMDDTHHmmssZ`)
3. `initialize_execution_log(...)` con evento `QUEUED` (si flag activo)
4. Pasar `execution_id` y `execution_log_path` al background task

Si `initialize_execution_log` falla: `logger.warning`; job sigue; `execution_log_status=WRITE_FAILED` en result.

### 3. Mecanismo de propagación (fuente canónica)

| Alternativa | Veredicto |
|-------------|-----------|
| Control Excel `ExecutionId` + `ExecutionLogPath` | **Canónica** |
| `ProcessId` | Secundaria; nullable; no sustituye `execution_id` |
| `process_key` | Solo informativo; **prohibido** correlacionar sola |
| Rutas de archivos | Ambiguas entre corridas |
| Body endpoints | Sin cambios de contrato |
| Job metadata RAM | Se pierde en reinicio Render |

**Flujo:**
1. Generate (éxito) escribe `ExecutionId`, `ExecutionLogPath` en control junto con `ProcessId`.
2. Finalize/Notify/Merge/Dry-run/Apply: tras resolver `bank_code`, `read_process_control_snapshot` → `execution_id`, `execution_log_path`.
3. Si faltan: **no** inferir por banco+fecha; warning + skip log (negocio continúa).
4. Idempotencia `already_generated` / `already_merged` / `already_applied`: append `SKIPPED_IDEMPOTENT` o `ALREADY_COMPLETED` al log **existente** de control (no nuevo archivo).

### 4. Nombre definitivo del archivo

```text
04 LOGS/execution_log_{bank_code}_{run_timestamp_utc}_{execution_id_short}.json
```

- `run_timestamp_utc`: UTC al `initialize` (ej. `20260612T101530Z`)
- `execution_id_short`: primeros 8 hex de `execution_id` sin guiones (ej. `a81f2c9e`)
- Ejemplo: `execution_log_banco_bogota_20260612T101530Z_a81f2c9e.json`

### 5. Schema JSON v1 completo

Raíz (`execution_log_schema_version = 1`):

```json
{
  "schema_version": 1,
  "execution_id": "uuid",
  "process_id": null,
  "process_key": null,
  "power_automate_run_id": null,

  "bank": { "code": "banco_bogota", "name": "Banco de Bogotá" },

  "dates": {
    "run_started_at": "2026-06-12T10:15:30.123456+00:00",
    "run_finished_at": null,
    "process_date": "2026-06-12",
    "report_date": null,
    "bank_transaction_dates": []
  },

  "summary": {
    "overall_status": "RUNNING",
    "current_step": "GENERATE",
    "last_attempted_step": "GENERATE",
    "last_successful_step": null,
    "failed_step": null,
    "can_retry": false,
    "requires_support": false,
    "warnings_count": 0,
    "errors_count": 0
  },

  "events": [],
  "artifacts": {},
  "affected_entities": [],
  "metrics": {},
  "manifest_snapshot": null,

  "logging": {
    "revision": 1,
    "last_updated_at": "ISO-8601 UTC",
    "write_status": "OK",
    "etag": null
  }
}
```

**Evento** (`events[]`):

```json
{
  "event_id": "uuid:GENERATE:1:FAILED:job-uuid",
  "sequence": 3,
  "step": "GENERATE",
  "substep": null,
  "attempt": 1,
  "status": "FAILED",
  "job_id": "uuid",
  "power_automate_run_id": null,
  "queued_at": "ISO-8601",
  "started_at": "ISO-8601",
  "finished_at": "ISO-8601",
  "duration_ms": 1200,
  "severity": "error",
  "error": {
    "error_code": "review_folder_not_empty",
    "exception_type": "ValueError",
    "technical_message": "sanitized",
    "user_message": null,
    "next_action": null
  },
  "metrics": {},
  "artifacts": [],
  "affected_entities": []
}
```

**Artefacto** (en `artifacts` keyed por role o en evento):

```json
{
  "role": "VALIDATION_FILE",
  "path": "sharepoint/rel/path",
  "file_name": "file.xlsx",
  "client": null,
  "credit": null,
  "application_type": null,
  "action": "CREATED",
  "status": "SUCCEEDED",
  "size_bytes": null,
  "sha256": null
}
```

**Entidad afectada** (`affected_entities[]`, clave `client|credit|application_type`):

```json
{
  "client": "EQUINORTE",
  "credit": "258",
  "application_type": "PAGO",
  "status": "PENDING_INPUT",
  "missing_documents": ["ACCOUNTING_PDF"],
  "warnings": [],
  "errors": ["missing_asiento_contable_pdf"]
}
```

### 6. Estados globales y de eventos

**Evento (`status`):** `QUEUED`, `STARTED`, `SUCCEEDED`, `FAILED`, `BLOCKED`, `PARTIAL`, `COMPLETED_WITH_WARNINGS`, `SKIPPED_IDEMPOTENT`, `ALREADY_COMPLETED`.

**Global (`summary.overall_status`):**

| Valor | Regla |
|-------|-------|
| `RUNNING` | Existe STARTED sin terminal en paso actual y sin fallo terminal global |
| `SUCCEEDED` | APPLY terminal `SUCCEEDED` o `ALREADY_COMPLETED` sin `FAILED` posterior |
| `FAILED` | Último intento terminal de algún paso es `FAILED` y no hay retry exitoso posterior |
| `BLOCKED` | Último terminal `BLOCKED` (dry-run/apply gate) sin recuperación |
| `PARTIAL` | Terminal `PARTIAL` en MERGE o APPLY sin fallo total |
| `COMPLETED_WITH_WARNINGS` | Terminal con warnings sin error bloqueante |
| `ABANDONED` | **Solo inferido offline** si `RUNNING` + último evento STARTED sin terminal y `last_updated_at` > umbral (ej. 6h); **nunca** durante ejecución activa |

**Resumen (no solo `last_completed_step`):**
- `current_step`: paso con STARTED sin terminal más reciente, o null si flujo terminó
- `last_attempted_step`: paso del evento más reciente
- `last_successful_step`: paso del último `SUCCEEDED`/`ALREADY_COMPLETED`/`SKIPPED_IDEMPOTENT` válido
- `failed_step`: paso del último `FAILED`/`BLOCKED` si no fue superado por retry exitoso

`can_retry` / `requires_support`: derivar de `operational_message_policy` existente cuando disponible.

### 7. Reglas de intentos y reintentos

- `attempt` = 1 + count eventos mismo `step`+`substep` con `status=STARTED` antes de este evento.
- Reintentos **no borran** intentos anteriores (ej. MERGE attempt 1 FAILED, attempt 2 SUCCEEDED).
- `SKIPPED_IDEMPOTENT` / `ALREADY_COMPLETED`: `attempt` no incrementa corrida nueva; se append al log canónico del control.

### 8. Reglas de deduplicación

`event_id = "{execution_id}:{step}:{substep or 'main'}:{attempt}:{status}:{job_id}"`

Antes de append: si `event_id` existe en `events[]`, **skip** (hook duplicado).

### 9. Estrategia de concurrencia y eTag

**Estado Graph hoy:** `MsGraphClient.put_bytes` sin `If-Match`; `get` devuelve `eTag` en metadatos (ya usado en amortización).

**Ampliación mínima aditiva:**
- `GraphApiPort.put_bytes(..., if_match: str | None = None)`
- `MsGraphClient`: header `If-Match` si presente; 412 → conflicto
- `execution_run_log`: guardar `logging.etag` tras cada lectura

**`record_execution_event` (función pública transaccional):**
1. GET metadata (eTag) + GET bytes JSON
2. Validar schema; dedup `event_id`
3. Append evento; asignar `sequence`; merge artifacts/entities/metrics
4. Recalcular `summary`; incrementar `revision`
5. PUT con `If-Match`; si 412: reintentar hasta 3 veces
6. Si sin eTag support: mismo flujo sin header (documentar limitación)

**No** usar lock en memoria como garantía.

### 10. Estrategia de actualización del resumen

Tras cada evento terminal o STARTED:
- Recalcular campos §6 desde `events[]` (determinista, función pura `recompute_summary(events)`).
- Actualizar `metrics` raíz = merge de métricas de eventos (último valor gana por clave).
- `artifacts` raíz = merge por `role+path` (último estado gana).
- `affected_entities` = merge por clave lógica §5.
- `run_finished_at` cuando `overall_status` ∈ {SUCCEEDED, FAILED, BLOCKED, PARTIAL, COMPLETED_WITH_WARNINGS, ABANDONED inferido}.

### 11. Snapshot de manifest

Campo raíz `manifest_snapshot` (no archivo extra), poblado en evento MERGE terminal:

```json
{
  "manifest_status": "PARTIAL",
  "eligible_for_dry_run": false,
  "complete_groups_count": 1,
  "incomplete_groups_count": 1,
  "failed_groups_count": 0,
  "outputs": [],
  "incomplete_groups": []
}
```

Copiar subset de `manifest_payload` actual (`merge_composite_validado_pdfs.py` ~1685); sin duplicar arrays enormes. Manifest canónico `merge_manifest_{bank}_{date}.json` **sin cambios**.

### 12. Artefactos registrados por paso

| Paso | Roles (`artifacts`) |
|------|---------------------|
| GENERATE | `VALIDATION_FILE`, `PROCESS_CONTROL_FILE`, `BANK_INPUT_FILE` |
| FINALIZE | `VALIDATION_FILE`, `HISTORICAL_FILE`, `SECRETARY_SUPPORT_FILE`, `AUDIT_WORKBOOK` |
| NOTIFY | `EMAIL_PDF`, `SECRETARY_SUPPORT_FILE`, `PROCESS_CONTROL_FILE` |
| MERGE | `HISTORICAL_FILE`, `EMAIL_PDF`, `MERGE_OUTPUT_PDF`, `MERGE_MANIFEST`, `PROCESS_CONTROL_FILE` |
| DRY_RUN | `MERGE_MANIFEST`, `HISTORICAL_FILE`, `ACCOUNTING_PDF` |
| APPLY | `AMORTIZATION_TABLE`, `ACCOUNTING_PDF`, `PROCESSED_ACCOUNTING_PDF`, `PROCESS_CONTROL_FILE` |

Acciones: `READ`, `CREATED`, `REUSED`, `UPDATED`, `UPLOADED`, `VERIFIED`, `MOVED`, `SKIPPED`, `FAILED`.

### 13. Métricas registradas por paso

| Paso | Métricas (cuando disponibles) |
|------|-------------------------------|
| GENERATE | `bank_movements_count`, `payment_groups_count`, `abono_*_groups_count`, `validated_rows` |
| FINALIZE | `support_rows`, `validated_rows`, `errors_count` |
| NOTIFY | `warnings_count` (control no actualizado) |
| MERGE | `outputs_count`, `incomplete_groups_count`, `*_skipped_count` |
| DRY_RUN | `tables_planned_count`, `warnings_count`, `errors_count`, entidades incompletas |
| APPLY | `tables_uploaded_count`, `tables_verified_count`, `tables_failed_count`, `ibr_updates_count`, `ibr_updates_skipped_count`, `accounting_pdfs_processed_count`, `accounting_pdfs_moved_count`, `accounting_pdfs_move_warnings_count` |

### 14. Sanitización de datos

`sanitize_for_execution_log(obj)` antes de serializar:
- Denylist keys: `authorization`, `client_secret`, `access_token`, `password`, `excel_base64`, `pdf_base64`, `GRAPH_*`
- Strip stack traces con patrones de secretos
- Max length strings (ej. 4000 chars technical_message)
- Permitir: rutas, nombres, conteos, códigos, clientes/créditos, hashes ya calculados

### 15. Manejo de fallos de logging

- `try/except` en cada hook; **nunca** raise al use case
- `logger.warning` estructurado: `execution_id`, `job_id`, `step`, error
- **No** fire-and-forget; await dentro del job
- Result aditivo:

```json
{
  "execution_id": "uuid",
  "execution_log_path": "...",
  "execution_log_status": "OK",
  "execution_log_warning": null
}
```

o `WRITE_FAILED` + warning humano.

### 16. Campos opcionales en `result`

`execution_id`, `execution_log_path`, `execution_log_status`, `execution_log_warning` — solo si flag activo; omitibles si flag off.

### 17. Impacto Parse JSON Power Automate

- Jobs `GET /jobs/{id}`: `result` con `additionalProperties` implícito (dict Python) → PA **no debería** romperse si solo lee campos conocidos.
- Flujos que Parse JSON con schema estricto sobre `result` completo: **verificar** en Flujos 1–4; campos nuevos son opcionales.
- **No** agregar a respuestas `202 queue` (solo `job_id`, `status`) para no tocar PA enqueue.
- Recomendación: smoke test PA con flag on leyendo solo campos existentes.

### 18. Archivos a crear

- `app/application/services/execution_run_log.py`
- `tests/test_execution_run_log.py`
- `tests/test_execution_run_log_hooks.py` (opcional, mocks Graph)

### 19. Archivos a modificar

- `app/domain/ports/graph.py`, `app/adapters/secondary/ms_graph_client.py` (eTag opcional)
- `app/application/config/payment_validation_settings.py` (flag)
- `app/application/use_cases/setup_merge_control_workbook.py` (columnas)
- `app/application/use_cases/payment_validation_process_control.py` (snapshot)
- `app/adapters/primary/http/routers/payment_validation.py` (hooks queue + jobs)
- `app/adapters/primary/http/routers/sharepoint.py` (notify/merge hooks)
- `app/application/use_cases/payment_validation_generate.py`
- `app/application/use_cases/payment_validation_finalize.py`
- `app/application/use_cases/send_validar_extractos_notification.py`
- `app/application/use_cases/merge_composite_validado_pdfs.py`
- `app/application/use_cases/amortization_fill_dry_run.py`
- `app/application/use_cases/amortization_fill_apply.py`

### 20. Tests requeridos

1. Generate fallido pre-ProcessId → log con `process_id=null`
2. Dos corridas mismo banco/día → dos archivos distintos
3. STARTED sin terminal visible en JSON
4. Flujo completo exitoso
5. Fallo Finalize
6. Notify OK, control no actualizado → warning/partial
7. Merge PARTIAL
8. Merge FAILED luego SUCCEEDED (2 intentos)
9. Dry-run BLOCKED
10. Apply PARTIAL
11. Apply exitoso
12. Apply ALREADY_COMPLETED
13. Eventos duplicados no duplican
14. `revision` incrementa
15. `manifest_snapshot` en mismo JSON
16. Entidades deduplicadas
17. Fallo log no falla negocio
18. Sin secretos en JSON
19. Contratos request/response legacy sin cambios (flag off)
20. Manifest canónico tests verdes
21. Control tests verdes
22. Schema PA-compatible (test snapshot JSON result keys)
23. Suite regresión completa `pytest`

### 21. Feature flag y rollout

```text
EXECUTION_RUN_LOG_ENABLED=false  # default
```

1. Deploy con `false` → regresión
2. Ambiente validación `true` → smoke 1 corrida completa
3. Producción `true` tras aprobación secretaría/soporte
4. Variable **opcional** (no obligatoria en Render)

### 22. Riesgos residuales

| Riesgo | Mitigación |
|--------|------------|
| Control sin ExecutionId (corrida legacy) | Skip log; warning |
| Conflicto eTag bajo carga | Reintentos + revision |
| JSON grande (APPLY) | Límites en arrays; resumen |
| Columna control nueva en workbook viejo | Setup idempotente repara columnas |
| Secretaría edita log JSON | Permisos SharePoint (§ retención) |

### 23. Plan de rollback

1. `EXECUTION_RUN_LOG_ENABLED=false` + redeploy
2. JSON en LOGS inertes
3. Revert deploy anterior si hooks causan error (flag off primero)

### 24. Fases de implementación

| Fase | Entregable |
|------|------------|
| **A** | Servicio + schema + tests unitarios + eTag |
| **B** | Columnas control + snapshot + setup idempotente |
| **C** | Hooks Generate/Finalize (queue + use case) |
| **D** | Hooks Notify/Merge |
| **E** | Hooks Dry-run/Apply + manifest_snapshot |
| **F** | Smoke PA + producción flag on |

Cada fase: flag off por defecto; tests verdes.

### 25. Commit sugerido

```text
feat(traceability): add execution run log JSON per validation flow

Introduce execution_id, SharePoint consolidated JSON in 04 LOGS,
control columns ExecutionId/ExecutionLogPath, and optional result
fields behind EXECUTION_RUN_LOG_ENABLED (default false).
```

---

## Approach (resumen)

Función pública `record_execution_event(...)` encapsula lectura, dedup, resumen, eTag y escritura. Routers registran `QUEUED`; use cases registran `STARTED` al inicio y terminal al final (try/except en router para FAILED temprano).

## Affected Areas

| Area | Impact |
|------|--------|
| `app/application/services/execution_run_log.py` | New |
| Control workbook + process_control | Modified (columnas aditivas) |
| Routers + 6 use cases | Modified (hooks) |
| Graph client/port | Modified (If-Match opcional) |

## Risks

Ver §22.

## Rollback Plan

Ver §23.

## Dependencies

Graph read/write LOGS; setup control idempotente; `resolve_logs_folder_path()`.

## Success Criteria

Ver criterios de aceptación del usuario (§ reporte) + regresión con flag `false` idéntica a hoy.
