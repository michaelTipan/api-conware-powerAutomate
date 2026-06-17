# Tasks: Execution Run Log

> **42 tareas** organizadas en **7 fases** (A–G).
> Todas las fases se implementan detrás de `EXECUTION_RUN_LOG_ENABLED=false`.

---

## Regla transversal: Fallos de logging vs. fallos de reserva

### Fallos no bloqueantes (logging)

No detienen el negocio. Ante fallo:

- escritura del execution log → `WRITE_FAILED`, warning, negocio continúa;
- checkpoints → omitido, `possible_gaps=true`;
- manifest snapshot → omitido;
- columnas `ExecutionId`/`ExecutionLogPath` → retry best-effort;
- polling → devolver estado real del job sin logging;
- terminal repair → `possible_gaps=true`;
- actualización de `logging_health` → warning.

### Fallos de reserva/concurrencia (bloqueantes)

Pueden impedir la creación de una nueva ejecución. Son condiciones fail-closed:

- no se puede determinar si el lock está ocupado;
- se agotan reintentos 412 en la adquisición de reserva;
- archivo de reserva tiene estructura inválida;
- Graph no ofrece la garantía condicional verificada (A0 fallido);
- dos reservas no pueden distinguirse con seguridad.

En estos casos:

- fail-closed para iniciar una nueva corrida;
- no crear segundo `execution_id`;
- no encolar un segundo Generate;
- mantener intacto el proceso existente.

El lock distribuido NO es un componente de logging. Es un mecanismo de concurrencia que protege la creación de ejecuciones.

---

## Phase A — Fundamentos, Graph/eTag y servicio central

- [ ] A0 — Verificar soporte real de If-Match en Graph PUT /content
  - Objetivo: Spike de integración aislado para confirmar que Graph devuelve 412 ante eTag obsoleto en el endpoint usado por `put_bytes`. Decisión de compuerta: `GRAPH_IF_MATCH_VERIFIED = true`.
  - Archivos creados: `tests/test_graph_if_match_validation.py` (marcado `@pytest.mark.graph_integration`)
  - Archivos modificados: ninguno productivo
  - Dependencias: ninguna
  - Implementación:
    1. Auditar `GraphApiPort.put_bytes` — confirmar que NO acepta `if_match` hoy.
    2. Auditar `MsGraphClient.put_bytes` — confirmar endpoint exacto: `PUT /sites/{site}/drives/{drive}/root:/{path}:/content`.
    3. Auditar respuesta Graph: confirmar que PUT devuelve `eTag` en body JSON (`driveItem`).
    4. Escribir spike de integración (script temporal, helper exclusivo de test, o llamada HTTP controlada). NO depender de una firma productiva que no existe todavía (A3 la agrega después):
       - Crear archivo de prueba via llamada HTTP directa al endpoint Graph.
       - Leer metadata via GET driveItem → extraer `eTag`.
       - PUT con `If-Match: {eTag}` correcto → confirmar 200 OK.
       - PUT con `If-Match: {eTag_antiguo}` → confirmar 412 Precondition Failed.
       - Confirmar que el contenido NO fue sobrescrito por el PUT rechazado.
    5. Documentar: método HTTP, endpoint, header exacto, formato eTag, comportamiento 412.
    6. Marcar tests: `@pytest.mark.graph_integration` — no se ejecutan en `pytest -q` normal.
  - Tests (marcados `graph_integration`):
    - `test_put_with_valid_etag_succeeds`
    - `test_put_with_stale_etag_returns_412`
    - `test_412_does_not_overwrite_content`
    - `test_etag_returned_in_response`
  - Comandos separados:
    - `python -m pytest -q` — suite unitaria normal, NO ejecuta A0.
    - `python -m pytest -m graph_integration -q` — solo spike A0.
  - Criterios de aceptación:
    - Evidencia escrita de: eTag inicial, actualización con eTag válido, 412 con eTag obsoleto, contenido no sobrescrito.
    - Si NO respeta: documentar alternativa y devolver a diseño. Detener A3, B5 y C3.
    - Decisión `GRAPH_IF_MATCH_VERIFIED` registrada como criterio conceptual.
  - Riesgos:
    - Graph podría ignorar If-Match en PUT /content (bloquea lock distribuido).
    - Endpoint sharepoint-backed vs onedrive-backed podría diferir.
    - Requiere credenciales y ambiente de validación para ejecutarse.
  - Rollback: No hay cambios productivos. Borrar test si se descarta.

---

- [ ] A1 — Modelos y enums
  - Objetivo: Crear todos los modelos de dominio definidos en diseño como dataclasses y StrEnum puros sin lógica Graph.
  - Archivos creados: `app/application/services/execution_context.py`, `tests/test_execution_context.py`
  - Archivos modificados: ninguno
  - Dependencias: ninguna
  - Implementación:
    1. `ExecutionEventStatus(StrEnum)` — 12 valores.
    2. `OverallStatus(StrEnum)` — 7 valores.
    3. `ExecutionCheckpoint(StrEnum)` — 25 valores.
    4. `ExecutionContext` frozen dataclass — sin `checkpoint_key`.
    5. `ExecutionError` frozen dataclass.
    6. `ExecutionArtifact` frozen dataclass.
    7. `AffectedEntity` frozen dataclass con tuples.
    8. `ExecutionLogWriteResult` dataclass (ok, revision, write_status, warning).
    9. Modelos auxiliares: `ManifestSnapshot`, `PollingObservation`, `ExecutionSummary`, `LoggingHealth`.
    10. Validar serialización ISO-8601 UTC con `utc_now_iso()`.
  - Tests:
    - Instantiation de cada dataclass y enum.
    - Frozen validation (no mutable).
    - Serialización de fechas UTC.
    - Valores de StrEnum exactos.
  - Criterios de aceptación:
    - Todos los modelos del diseño §Models and Types creados.
    - Sin imports de Graph, httpx, asyncio.
    - Tests unitarios puros verdes.
  - Riesgos: Ninguno significativo (tipos puros).
  - Rollback: Borrar archivo creado.

---

- [ ] A2 — Sanitización y truncamiento
  - Objetivo: Crear el módulo de sanitización para eliminar secretos y truncar colecciones antes de serializar al JSON de ejecución.
  - Archivos creados: `app/application/services/execution_log_sanitizer.py`, `tests/test_execution_log_sanitizer.py`
  - Archivos modificados: ninguno
  - Dependencias: ninguna
  - Implementación:
    1. `sanitize_for_execution_log(obj, depth=0, max_depth=8)` — recursivo.
    2. Denylist case-insensitive substring: `authorization`, `access_token`, `refresh_token`, `client_secret`, `password`, `api_key`, `excel_base64`, `pdf_base64`, `cookie`, `set-cookie`.
    3. Max string 4000 chars.
    4. Max list scan 500 antes de truncar.
    5. `bytes` → `"<bytes>"`.
    6. Excepciones → type name + message sanitizada.
    7. URLs: strip query `token=`/`access_token=`.
    8. Preservar: paths, client, credit, file names, counts, error codes.
    9. `truncate_collection(items, max_items=100)` → wrapper `TruncatedList`.
    10. Schema TruncatedList: `{items, total_count, included_count, truncated}`.
  - Tests:
    - `test_denylist_keys_removed` — cada key de la denylist.
    - `test_denylist_case_insensitive` — `Authorization`, `ACCESS_TOKEN`.
    - `test_url_token_stripped` — `?token=xxx&other=keep`.
    - `test_bytes_replaced` — `b"data"` → `"<bytes>"`.
    - `test_exception_sanitized` — type + message only.
    - `test_max_depth` — objeto anidado a 9 niveles.
    - `test_max_string_length` — string de 5000 chars → 4000.
    - `test_truncated_list_wrapper` — 150 items → 100 + `truncated=true`.
    - `test_preserves_business_fields` — client, credit, path, error_code intactos.
  - Criterios de aceptación:
    - Ningún secreto pasa sanitización.
    - TruncatedList schema consistente.
    - Tests puros sin I/O.
  - Riesgos: Falso positivo en denylist (key legítima contiene substring). Mitigar con preservación explícita.
  - Rollback: Borrar archivos creados.

---

- [ ] A3 — Extensión Graph (If-Match + metadata)
  - Objetivo: Extender `GraphApiPort` y `MsGraphClient` de forma aditiva para soportar escritura condicional y lectura de metadata.
  - Archivos creados: ninguno
  - Archivos modificados: `app/domain/ports/graph.py`, `app/adapters/secondary/ms_graph_client.py`, todos los fakes/mocks en tests que implementan `GraphApiPort`
  - Dependencias: A0 (debe confirmar que If-Match funciona)
  - Implementación:
    1. `GraphApiPort.put_bytes(endpoint, content, content_type, if_match: str | None = None)` — parámetro opcional, backward-compatible.
    2. `MsGraphClient.put_bytes` — si `if_match` presente: `headers["If-Match"] = if_match`.
    3. Nuevo error `GraphPreconditionFailedError(HTTPStatusError)` para 412.
    4. `MsGraphClient.put_bytes` — ante 412: NO reintentar internamente (el caller decide). Raise `GraphPreconditionFailedError`.
    5. `GraphApiPort.get_drive_item_metadata(site_id, drive_id, rel_path) -> DriveItemMetadata` — nuevo método.
    6. `MsGraphClient.get_drive_item_metadata` — `GET /sites/{site}/drives/{drive}/root:/{path}:` → extraer `eTag`, `size`, `lastModifiedDateTime`.
    7. `DriveItemMetadata` dataclass: `etag: str`, `size: int`, `last_modified: str`.
    8. Actualizar TODOS los fakes en tests para aceptar `if_match=None` sin romper.
  - Tests:
    - `test_put_bytes_without_if_match_unchanged` — comportamiento legacy idéntico.
    - `test_put_bytes_with_if_match_sends_header` — mock httpx, verificar header.
    - `test_put_bytes_412_raises_precondition_failed` — response 412 → exception.
    - `test_get_drive_item_metadata_returns_etag` — mock response.
    - `test_existing_callers_unaffected` — compilación de todas las llamadas existentes.
  - Criterios de aceptación:
    - `python -m compileall app -q` verde.
    - `pytest` completo verde (ningún test existente roto).
    - Todas las llamadas existentes a `put_bytes` siguen funcionando sin `if_match`.
  - Riesgos:
    - Fakes en tests con firma diferente → scan completo de `put_bytes` en tests.
    - `get_drive_item_metadata` endpoint podría variar por tipo de drive.
  - Rollback: Revert de los 3 archivos modificados.

---

- [ ] A4 — Documento schema v1
  - Objetivo: Implementar creación, validación, lectura y serialización del documento JSON de ejecución v1.
  - Archivos creados: funciones dentro de `app/application/services/execution_run_log.py` (esqueleto)
  - Archivos modificados: ninguno
  - Dependencias: A1 (modelos), A2 (sanitizer)
  - Implementación:
    1. `create_empty_execution_document(execution_id, bank_code, process_date, ...)` → dict con todos los campos raíz del schema v1.
    2. `validate_execution_log_document(doc) -> ExecutionLogDocument` — validar `schema_version`, campos requeridos.
    3. Colecciones como `TruncatedList` wrappers: `artifacts`, `affected_entities`.
    4. `events` como lista directa (nunca truncada).
    5. `logging_health` con defaults: `status=OK`, `possible_gaps=false`, `write_failures_count=0`.
    6. `logging.revision` empieza en 1.
    7. Serialización JSON con `default=str` para datetimes.
    8. Lectura: parsear JSON → validar schema → devolver dict tipado.
  - Tests:
    - `test_create_empty_document_has_all_fields`.
    - `test_validate_rejects_wrong_schema_version`.
    - `test_collections_use_truncated_list_wrapper`.
    - `test_events_is_plain_list`.
    - `test_serialization_roundtrip`.
  - Criterios de aceptación:
    - Documento vacío pasa validación.
    - Schema v1 completo según proposal §5.
    - Sin I/O de Graph (funciones puras sobre dicts).
  - Riesgos: Ninguno significativo.
  - Rollback: Borrar funciones del esqueleto.

---

- [ ] A5 — Merge concurrente
  - Objetivo: Implementar y probar la función de merge de documentos para resolver conflictos eTag.
  - Archivos creados: tests en `tests/test_execution_run_log_merge.py`
  - Archivos modificados: `app/application/services/execution_run_log.py` (agregar función)
  - Dependencias: A4 (schema)
  - Implementación:
    1. `merge_execution_documents(base, incoming) -> dict`.
    2. Events: union por `event_id`; sort por `sequence`, luego `started_at`; reasignar `sequence` 1..N si conflicto.
    3. Artifacts: merge por `artifact_id = role|normalized_path`; incoming gana.
    4. Entities: key `client|credit|application_type`.
    5. Metrics: per-key max numérico o last-wins para strings.
    6. `job_polling`: merge por `job_id`; merge `observed_statuses` counts.
    7. `logging_health.possible_gaps = true` si merge detecta secuencias solapadas.
    8. Llamar `recompute_execution_summary` sobre resultado.
  - Tests:
    - `test_merge_deduplicates_events_by_event_id`.
    - `test_merge_resequences_after_union`.
    - `test_merge_artifacts_incoming_wins`.
    - `test_merge_entities_by_logical_key`.
    - `test_merge_polling_by_job_id`.
    - `test_merge_sets_possible_gaps`.
    - `test_merge_calls_recompute_summary`.
  - Criterios de aceptación:
    - No last-write-wins sobre documento completo.
    - Eventos nunca perdidos ni duplicados.
    - `possible_gaps` activado en conflicto.
  - Riesgos: Merge complejo puede tener edge cases con secuencias cruzadas.
  - Rollback: Revert de la función y tests.

---

- [ ] A6 — Summary y clasificación terminal
  - Objetivo: Implementar funciones puras de clasificación de eventos terminales y recomputación del summary.
  - Archivos creados: tests en `tests/test_execution_run_log_summary.py`
  - Archivos modificados: `app/application/services/execution_run_log.py` (agregar funciones)
  - Dependencias: A1 (enums)
  - Implementación:
    1. `classify_terminal_event(step, result, error) -> ExecutionEventStatus` — tabla de 20+ mapeos del diseño.
    2. `recompute_execution_summary(document) -> dict`.
    3. Priority: SUCCEEDED > FAILED/BLOCKED > PARTIAL > COMPLETED_WITH_WARNINGS > RUNNING > WAITING_FOR_NEXT_STEP.
    4. Retry exitoso limpia `failed_step`.
    5. Matriz `WAITING_FOR_NEXT_STEP` (exacta del diseño aprobado):
       | Etapa completada | `waiting_for`              | `next_expected_step` |
       |------------------|----------------------------|----------------------|
       | GENERATE         | SECRETARY_REVIEW           | FINALIZE             |
       | FINALIZE         | NOTIFY_EXECUTION           | NOTIFY               |
       | NOTIFY           | ACCOUNTING_DOCUMENT_UPLOAD | MERGE                |
       | MERGE            | DRY_RUN_EXECUTION          | DRY_RUN              |
       | DRY_RUN listo    | APPLY_EXECUTION            | APPLY                |
       | APPLY            | null                       | null                 |
    6. `flow_completed_at` solo tras Apply success/ALREADY_COMPLETED.
    7. Post-completion: `overall_status` permanece SUCCEEDED; incrementar `post_completion_events_count`.
    8. `last_terminal_event_at`, `last_activity_at` correctamente calculados.
  - Tests:
    - `test_classify_generate_success` → SUCCEEDED.
    - `test_classify_generate_already_generated` → SKIPPED_IDEMPOTENT.
    - `test_classify_merge_partial` → PARTIAL.
    - `test_classify_dry_run_blocked` → BLOCKED.
    - `test_classify_apply_already_applied` → ALREADY_COMPLETED.
    - `test_summary_waiting_after_generate` — `waiting_for=SECRETARY_REVIEW`, `next_expected_step=FINALIZE`.
    - `test_summary_waiting_after_finalize` — `waiting_for=NOTIFY_EXECUTION`, `next_expected_step=NOTIFY`.
    - `test_summary_waiting_after_notify` — `waiting_for=ACCOUNTING_DOCUMENT_UPLOAD`, `next_expected_step=MERGE`.
    - `test_summary_waiting_after_merge` — `waiting_for=DRY_RUN_EXECUTION`, `next_expected_step=DRY_RUN`.
    - `test_summary_waiting_after_dry_run` — `waiting_for=APPLY_EXECUTION`, `next_expected_step=APPLY`.
    - `test_summary_succeeded_only_after_apply`.
    - `test_summary_retry_clears_failed_step`.
    - `test_summary_post_completion_stays_succeeded`.
    - `test_summary_flow_completed_at_only_on_apply`.
    - `test_summary_recoverable_failure_no_flow_completed`.
  - Criterios de aceptación:
    - Toda la tabla de clasificación del diseño cubierta.
    - Matriz WAITING coincide exactamente con el diseño aprobado.
    - Funciones puras sin I/O.
  - Riesgos: Complejidad de precedencia entre estados.
  - Rollback: Revert funciones y tests.

---

- [ ] A7 — Persistencia central
  - Objetivo: Implementar el módulo de persistencia con eTag retries, merge en 412, y manejo de errores sin propagación al negocio. Solo funciones de escritura/lectura central. Polling y repair se implementan en Phase F.
  - Archivos creados: `app/application/services/execution_run_log.py` (completar), `tests/test_execution_run_log_persistence.py`
  - Archivos modificados: ninguno productivo adicional
  - Dependencias: A3 (Graph If-Match), A4 (schema), A5 (merge), A6 (summary)
  - Implementación:
    1. `initialize_execution_log(graph, site_id, drive_id, execution_id, bank_code, ...) -> ExecutionLogWriteResult`.
    2. `record_execution_event(graph, context, status, **kwargs) -> ExecutionLogWriteResult`.
    3. `reconstruct_execution_log_from_control(...) -> ExecutionLogWriteResult`.
    4. Write loop: max 5 retries; backoff `0.2 * (2**n)` capped 2s.
    5. Ante 412: `get_drive_item_metadata` + GET bytes → merge → retry PUT.
    6. Ante 404: init si reconstructing.
    7. JSON corrupto: tratar como vacío + `possible_gaps=true`.
    8. eTag NUNCA almacenado en JSON.
    9. Todo fallo de logging → `ExecutionLogWriteResult(ok=False, write_status="WRITE_FAILED")`. NUNCA raise.
    10. Interfaces tipadas para polling (stubs): `record_job_poll(...)` y `ensure_terminal_job_event_logged(...)` con firma definida y `raise NotImplementedError` o pass-through, para que A8 y C pueda depender de las firmas. La implementación completa se realiza en F1–F3.
  - Tests:
    - `test_initialize_creates_document` — fake Graph.
    - `test_record_event_appends_and_increments_revision`.
    - `test_412_triggers_reread_merge_retry`.
    - `test_max_retries_returns_write_failed`.
    - `test_corrupt_json_treated_as_empty`.
    - `test_404_on_read_initializes`.
    - `test_logging_failure_never_raises`.
    - `test_etag_not_stored_in_json`.
  - Criterios de aceptación:
    - Ninguna excepción de logging se propaga al caller.
    - eTag retry funcional con fake Graph.
    - Merge invocado en conflicto 412.
    - Firmas de polling definidas pero NO implementadas (→ F1–F3).
  - Riesgos: Complejidad del loop de reintentos. Fake Graph debe simular 412 fielmente.
  - Rollback: Revert del módulo.

---

- [ ] A8 — Wrapper execution_step
  - Objetivo: Crear el context manager que envuelve cada ejecución de paso registrando STARTED, terminal, checkpoints, y manejando flag off.
  - Archivos creados: `app/application/services/execution_step.py`, `tests/test_execution_step.py`
  - Archivos modificados: ninguno
  - Dependencias: A1 (modelos), A7 (persistencia)
  - Implementación:
    1. `execution_step(graph, context, site_id, drive_id)` — async context manager.
    2. `TraceHandle` con: `checkpoint(label, checkpoint_key=..., metrics=...)`, `complete_from_result(result)`.
    3. Entrada: `has_terminal_for_invocation()` antes de STARTED (prevenir doble).
    4. Excepción: FAILED si no `terminal_recorded`.
    5. `terminal_recorded` flag previene doble terminal.
    6. `event_id` dedup como backstop.
    7. Flag false: yield no-op `TraceHandle`, sin Graph calls.
    8. Exception re-raised sin modificar.
    9. Fallo de logging NO modifica la excepción original.
  - Tests:
    - `test_wrapper_records_started`.
    - `test_wrapper_records_terminal_on_success`.
    - `test_wrapper_records_failed_on_exception`.
    - `test_wrapper_reraises_original_exception`.
    - `test_logging_failure_does_not_change_exception`.
    - `test_complete_from_result_classifies`.
    - `test_checkpoint_with_key`.
    - `test_no_op_when_flag_false`.
    - `test_no_double_terminal`.
  - Criterios de aceptación:
    - Original exception siempre re-raised.
    - Flag false = cero llamadas a Graph.
    - Single terminal por invocación.
  - Riesgos: Edge case donde logging falla en __aexit__ y modifica traceback.
  - Rollback: Borrar archivos creados.

---

## Phase B — Control Excel y archivos de reserva

- [ ] B1 — Feature flag
  - Objetivo: Agregar `EXECUTION_RUN_LOG_ENABLED` con default false y verificar que el sistema actual no cambia.
  - Archivos creados: `tests/test_execution_run_log_flag.py`
  - Archivos modificados: `app/application/config/payment_validation_settings.py`
  - Dependencias: ninguna
  - Implementación:
    1. `execution_run_log_enabled() -> bool` — `os.getenv("EXECUTION_RUN_LOG_ENABLED", "false").strip().lower() in ("1", "true", "yes")`.
    2. Colocar junto a funciones de configuración existentes.
  - Tests:
    - `test_flag_default_false`.
    - `test_flag_true_values` — "1", "true", "yes", "TRUE", " True ".
    - `test_flag_false_values` — "0", "false", "no", "anything", "".
    - `test_flag_false_no_graph_logging_calls` — mock assert zero calls.
    - `test_flag_false_wrappers_are_noop`.
  - Criterios de aceptación:
    - Default false sin variable.
    - Comportamiento actual idéntico con flag false.
  - Riesgos: Ninguno.
  - Rollback: Revert del archivo modificado.

---

- [ ] B2 — Columnas del control Excel
  - Objetivo: Agregar `ExecutionId` y `ExecutionLogPath` al final de `PROCESS_CONTROL_EXTENSION_COLUMNS`, ocultas, protegidas, reparación idempotente.
  - Archivos creados: `tests/test_execution_control_columns.py`
  - Archivos modificados: `app/application/use_cases/setup_merge_control_workbook.py`
  - Dependencias: ninguna
  - Implementación:
    1. Agregar `"ExecutionId"`, `"ExecutionLogPath"` al final de `PROCESS_CONTROL_EXTENSION_COLUMNS`.
    2. Setup: crear/reparar columnas si faltan.
    3. Columnas ocultas (`column.hidden = True`).
    4. Columnas protegidas (bloqueadas para edición manual).
    5. No tocar row 2 data existente.
    6. Controles legacy sin las columnas: setup las agrega idempotentemente.
  - Tests:
    - `test_setup_adds_execution_columns`.
    - `test_columns_are_hidden`.
    - `test_columns_are_protected`.
    - `test_existing_row2_data_unchanged`.
    - `test_idempotent_repair_no_duplicates`.
    - `test_legacy_workbook_repaired`.
    - `test_columns_readable_by_openpyxl`.
    - `test_no_operational_columns_changed`.
  - Criterios de aceptación:
    - Columnas al final, ocultas, protegidas.
    - Legacy workbooks reparados sin pérdida de datos.
    - Suite de setup existente verde.
  - Riesgos: Offset de columnas podría afectar lecturas por índice. Verificar que lecturas usan nombres.
  - Rollback: Revert de `setup_merge_control_workbook.py`.

---

- [ ] B3 — Snapshot del control
  - Objetivo: Extender `ProcessControlSnapshot` para incluir `execution_id` y `execution_log_path` como campos opcionales.
  - Archivos creados: ninguno
  - Archivos modificados: `app/application/use_cases/payment_validation_process_control.py`, tests asociados
  - Dependencias: B2 (columnas existen)
  - Implementación:
    1. Agregar `execution_id: str | None = None` y `execution_log_path: str | None = None` a `ProcessControlSnapshot`.
    2. Lector: parsear las columnas si existen; `None` si vacías.
    3. NUNCA inferir de banco o fecha.
    4. Empty → `None`, no string vacío.
  - Tests:
    - `test_snapshot_reads_execution_id_when_present`.
    - `test_snapshot_returns_none_when_empty`.
    - `test_snapshot_does_not_infer_from_bank`.
    - `test_existing_snapshot_fields_unchanged`.
  - Criterios de aceptación:
    - Campos opcionales no rompen lecturas existentes.
    - `ProcessControlSnapshot` backward-compatible.
  - Riesgos: Consumidores que usan unpacking posicional. Verificar que todos usen keyword/attribute access.
  - Rollback: Revert del archivo modificado.

---

- [ ] B4 — Archivos de reserva precreados en Setup
  - Objetivo: Setup precrea idempotentemente los archivos de reserva de lock para cada banco en `00 CONTROL/`.
  - Archivos creados: `tests/test_execution_reservation_files.py`
  - Archivos modificados: `app/application/use_cases/setup_merge_control_workbook.py`
  - Dependencias: A3 (Graph put_bytes)
  - Implementación:
    1. `setup_merge_control_workbook` crea `execution_run_reservation_banco_bogota.json` en `00 CONTROL/`.
    2. Crea `execution_run_reservation_banco_bancolombia.json` en `00 CONTROL/`.
    3. Schema inicial: `{schema_version: 1, bank_code, locked: false, owner_execution_id: null, owner_invocation_id: null, acquired_at: null, expires_at: null, revision: 0}`.
    4. Si archivo ya existe: NO sobrescribir (idempotente).
    5. Verificar: intentar GET primero, crear solo si 404.
  - Política de flag:
    - Setup es una acción administrativa de infraestructura.
    - Precrea los archivos incluso con `EXECUTION_RUN_LOG_ENABLED=false`.
    - El runtime con flag false NO lee ni escribe estos archivos.
    - Los endpoints principales NO generan llamadas Graph adicionales con flag false.
    - Esto evita contradicción con la regresión flag-off: Setup crea infraestructura, flag controla runtime.
  - Tests:
    - `test_setup_creates_reservation_file_bogota`.
    - `test_setup_creates_reservation_file_bancolombia`.
    - `test_filenames_have_no_invalid_characters` — no `:`, no `?`, no `*`.
    - `test_exactly_one_file_per_bank`.
    - `test_existing_file_not_overwritten`.
    - `test_initial_schema_correct` — locked=false, revision=0.
    - `test_setup_creates_regardless_of_flag` — infraestructura vs runtime.
    - `test_runtime_flag_false_no_reservation_reads`.
  - Criterios de aceptación:
    - Nombres con `_` no `:`.
    - Exactamente un archivo por banco.
    - No sobrescribir.
    - Setup existente verde.
    - Runtime con flag false = cero I/O sobre archivos de reserva.
  - Riesgos: Setup podría fallar si permisos SharePoint insuficientes.
  - Rollback: Revert de `setup_merge_control_workbook.py`.

---

- [ ] B5 — Lock local + reserva Graph condicional
  - Objetivo: Crear módulo de lock per-bank con asyncio + adquisición condicional vía eTag en archivo de reserva Graph. Este NO es un componente de logging; es un mecanismo de concurrencia que protege la creación de ejecuciones.
  - Archivos creados: `app/application/services/execution_bank_lock.py`, `tests/test_execution_bank_lock.py`
  - Archivos modificados: ninguno
  - Dependencias: A0 (If-Match verificado), A3 (Graph eTag), B4 (archivos precreados)
  - Implementación:
    1. `_bank_locks: dict[str, asyncio.Lock]` — registry lazy con lock de inicialización.
    2. `acquire_bank_execution_lock(bank_code)` → asyncio.Lock context.
    3. `acquire_graph_reservation(graph, site_id, drive_id, bank_code, execution_id, invocation_id)`:
       - Leer archivo + eTag.
       - Evaluar `locked` y `expires_at`.
       - Si libre o vencido: construir reserva con TTL 5 min.
       - PUT con If-Match.
       - Ante 412: releer, reevaluar, retry (max 3).
       - Ante lock ocupado no vencido: `active_process_exists` equivalente.
       - Agotados reintentos: fail-closed.
    4. `release_graph_reservation(graph, site_id, drive_id, bank_code)`:
       - Leer + eTag.
       - Escribir `locked=false, owner_execution_id=null, owner_invocation_id=null`.
       - PUT con If-Match.
       - NO borrar archivo.
    5. TTL lógico: comparar `expires_at` vs `utc_now()`.
    6. NO renovación durante Generate.
  - Fallos de reserva (bloqueantes):
    - Si no se puede determinar estado del lock: fail-closed.
    - Si se agotan reintentos 412: fail-closed.
    - Si archivo de reserva tiene estructura inválida: fail-closed.
    - No crear segundo execution_id ante ambigüedad.
    - No encolar segundo Generate.
  - Tests:
    - `test_acquire_free_lock_succeeds`.
    - `test_acquire_occupied_lock_fails_closed`.
    - `test_acquire_expired_lock_recovers`.
    - `test_two_concurrent_acquires_only_one_wins`.
    - `test_412_retries_with_reread`.
    - `test_max_retries_fail_closed`.
    - `test_release_sets_locked_false`.
    - `test_release_does_not_delete_file`.
    - `test_no_renewal_method_exists`.
    - `test_ttl_comparison_logic`.
    - `test_flag_false_skips_graph_reservation`.
    - `test_invalid_reservation_structure_fails_closed`.
    - `test_undetermined_lock_state_fails_closed`.
  - Criterios de aceptación:
    - Solo una adquisición concurrente gana.
    - 412 provoca relectura, no abort inmediato.
    - Release no elimina archivo.
    - No existe renovación.
    - Fail-closed ante ambigüedad.
  - Riesgos: Timing window entre GET y PUT. Mitigado por eTag.
  - Rollback: Borrar archivos creados.

---

- [ ] B6 — ensure_execution_control_link
  - Objetivo: Función best-effort para vincular ExecutionId/ExecutionLogPath en control cuando faltó inicialmente.
  - Archivos creados: `tests/test_ensure_execution_control_link.py`
  - Archivos modificados: `app/application/use_cases/payment_validation_process_control.py`
  - Dependencias: B2 (columnas), B3 (snapshot)
  - Implementación:
    1. `ensure_execution_control_link(graph, site_id, drive_id, bank_code, execution_id, execution_log_path) -> bool`.
    2. Leer control snapshot.
    3. Si ya vinculado con mismo execution_id: return True.
    4. Si vacío: escribir ExecutionId + ExecutionLogPath.
    5. Si diferente execution_id: NO sobrescribir (otro proceso).
    6. Fallo Graph: return False, no raise.
    7. NO modificar columnas operativas.
  - Tests:
    - `test_link_already_present_returns_true`.
    - `test_link_missing_writes_and_returns_true`.
    - `test_link_different_execution_does_not_overwrite`.
    - `test_graph_failure_returns_false`.
    - `test_operational_columns_unchanged`.
    - `test_idempotent_retry`.
  - Criterios de aceptación:
    - Best-effort: nunca raise.
    - No modifica columnas operativas.
    - Idempotente.
  - Riesgos: Race condition si dos procesos intentan vincular simultáneamente.
  - Rollback: Revert del archivo modificado.

---

## Phase C — Generate y Finalize

- [ ] C1 — Metadata de trace en job stores
  - Objetivo: Helper canónico para agregar trace metadata a ambos stores de jobs (JobManager + `_validation_jobs`) sin duplicar lógica.
  - Archivos creados: ninguno (helper dentro de execution_context.py o execution_run_log.py)
  - Archivos modificados: `app/application/services/execution_context.py`
  - Dependencias: A1 (modelos)
  - Implementación:
    1. `EXECUTION_TRACE_KEYS` constante.
    2. `attach_execution_trace(job: dict, context: ExecutionContext, write_result: ExecutionLogWriteResult | None)`.
    3. Helper lee context y write_result → escribe keys en job dict.
    4. Helper compartido por ambos routers.
  - Tests:
    - `test_attach_trace_sets_all_keys`.
    - `test_attach_trace_handles_none_context`.
    - `test_attach_trace_handles_write_failed`.
  - Criterios de aceptación:
    - Un solo helper, no dos copias.
    - Keys aditivas, no destructivas.
  - Riesgos: Ninguno significativo.
  - Rollback: Revert.

---

- [ ] C2 — Reserva de invocación atómica
  - Objetivo: Implementar `reserve_execution_invocation` con attempt calculation basada en `max(attempt)`, invocation_id, y eTag loop optimista.
  - Archivos creados: `tests/test_reserve_execution_invocation.py`
  - Archivos modificados: `app/application/services/execution_context.py`
  - Dependencias: A3 (Graph eTag), A7 (persistencia)
  - Implementación:
    1. `reserve_execution_invocation(graph, site_id, drive_id, *, execution_id, execution_log_path, job_id, step, substep, bank_code, process_date, ...) -> ExecutionContext`.
    2. Leer JSON + eTag.
    3. Calcular `attempt = 1 + max(attempt existente para step + substep)`. Considerar TODOS los eventos del mismo step/substep que tengan attempt: REQUEST_RECEIVED, QUEUED, STARTED, y terminales. NO usar `count(STARTED)` porque puede producir attempts duplicados cuando dos invocaciones están QUEUED antes del primer STARTED.
    4. Generar `invocation_id = uuid4()`.
    5. Agregar evento de reserva (REQUEST_RECEIVED o QUEUED) con el attempt calculado. La escritura condicional persiste el attempt antes de devolver ExecutionContext.
    6. PUT con If-Match.
    7. Ante 412: releer, recalcular attempt (puede haber cambiado), retry.
    8. Devolver ExecutionContext definitivo con attempt correcto.
  - Tests:
    - `test_first_reserve_merge_gets_attempt_1`.
    - `test_second_reserve_before_started_gets_attempt_2`.
    - `test_412_rereads_and_gets_updated_attempt`.
    - `test_terminal_reuses_exactly_reserved_attempt`.
    - `test_no_two_invocations_share_same_attempt_for_step`.
    - `test_reserve_appends_queued_event`.
    - `test_reserve_creates_immutable_context`.
  - Criterios de aceptación:
    - Dos llamadas concurrentes al mismo step → attempts diferentes.
    - Attempt persistido en JSON ANTES de devolver context.
    - eTag loop con recálculo.
    - Context devuelto inmutable.
    - No existen dos `invocation_id` distintos con el mismo attempt para el mismo step/substep.
  - Riesgos: Complejidad del loop optimista. Recálculo correcto tras 412 es crítico.
  - Rollback: Revert.

---

- [ ] C3 — Generate nuevo con reservation flow
  - Objetivo: Integrar reservation, log, vínculo control, lock y enqueue detrás del flag en `queue_generate`.
  - Archivos creados: ninguno
  - Archivos modificados: `app/adapters/primary/http/routers/payment_validation.py`
  - Dependencias: A7, A8, B1, B2, B3, B5, B6, C1, C2
  - Implementación:
    1. Dentro de `queue_generate`, tras validaciones, si flag on:
       - acquire bank lock (reserva/concurrencia — bloqueante).
       - read control snapshot.
       - decide: new / reuse / reject.
       - `execution_id = uuid4()`.
       - `initialize_execution_log` (logging — no bloqueante).
       - `ensure_execution_control_link` (logging — best-effort).
       - `attach_execution_trace` al job.
       - enqueue background task.
       - release lock en `finally`.
    2. Si flag off: flujo actual sin cambios.
    3. Fallo de JSON (logging): WRITE_FAILED, warning, enqueue procede.
    4. Fallo de control link (logging): warning, enqueue procede.
    5. Fallo de lock/reserva (concurrencia): fail-closed, NO encolar, NO crear segundo execution_id.
    6. Fallo de enqueue: release lock, registrar FAILED/REQUEST_REJECTED si log existe.
  - Tests:
    - `test_generate_new_creates_execution_id`.
    - `test_generate_json_failure_still_enqueues`.
    - `test_generate_control_failure_still_enqueues`.
    - `test_generate_both_logging_fail_still_enqueues`.
    - `test_generate_reservation_failure_does_not_enqueue`.
    - `test_generate_enqueue_failure_releases_lock`.
    - `test_generate_flag_false_unchanged`.
    - `test_generate_lock_released_in_finally`.
  - Criterios de aceptación:
    - Generate nunca falla por causa exclusiva del logging.
    - Fallo de reserva SÍ detiene el enqueue (concurrencia, no logging).
    - Lock siempre liberado.
    - Flag false = comportamiento idéntico.
  - Riesgos: Complejidad de la integración. Muchas dependencias.
  - Rollback: Revert de `payment_validation.py`. Flag false = safety net.

---

- [ ] C4 — Generate idempotente y proceso activo
  - Objetivo: Manejar correctamente los casos de idempotencia y proceso activo sin crear logs huérfanos.
  - Archivos creados: ninguno
  - Archivos modificados: `app/adapters/primary/http/routers/payment_validation.py`
  - Dependencias: C3
  - Implementación:
    1. `already_generated` con `process_key` match: reutilizar `ExecutionId`/`ExecutionLogPath` del control.
    2. Append `SKIPPED_IDEMPOTENT` al log existente.
    3. No crear nuevo archivo ni nuevo execution_id.
    4. `active_process_exists`: NO crear execution_id nuevo.
    5. Si `ExecutionId` en control activo: MAY append REQUEST_REJECTED.
  - Tests:
    - `test_idempotent_generate_reuses_log`.
    - `test_idempotent_appends_skipped_event`.
    - `test_idempotent_no_new_file`.
    - `test_active_process_no_new_execution_id`.
    - `test_active_process_may_reject_to_log`.
  - Criterios de aceptación:
    - No orphan logs.
    - Idempotencia preservada.
    - Reutilización de contexto.
  - Riesgos: Control sin ExecutionId → debe skip silenciosamente.
  - Rollback: Revert. Flag false = safety.

---

- [ ] C5 — Runner Generate y checkpoints
  - Objetivo: Integrar `execution_step` wrapper y checkpoints en `_run_generate_job` y el use case Generate.
  - Archivos creados: ninguno
  - Archivos modificados: `app/adapters/primary/http/routers/payment_validation.py`, `app/application/use_cases/payment_validation_generate.py`
  - Dependencias: A8 (wrapper), C3
  - Implementación:
    1. En `_run_generate_job`: envolver con `execution_step(graph, context, ...)`.
    2. Pasar `trace` handle al use case como parámetro opcional.
    3. En use case: `trace.checkpoint(GENERATE_BANK_FILE_READ, ...)`, `GENERATE_WORKBOOK_UPLOADED`, `GENERATE_CONTROL_UPDATED`.
    4. Terminal via `complete_from_result(result)`.
    5. Exception: wrapper registra FAILED automáticamente.
    6. Resultado financiero sin cambios.
  - Tests:
    - `test_runner_uses_execution_step`.
    - `test_checkpoints_emitted`.
    - `test_terminal_classified_correctly`.
    - `test_exception_records_failed`.
    - `test_result_unchanged_with_flag_on`.
  - Criterios de aceptación:
    - Resultado de Generate idéntico.
    - Checkpoints registrados correctamente.
    - Exception re-raised.
  - Riesgos: Pasar trace al use case puede requerir cambio de firma.
  - Rollback: Revert de ambos archivos.

---

- [ ] C6 — Finalize hooks y checkpoints
  - Objetivo: Integrar trazabilidad en Finalize: context desde control, reserva de attempt, wrapper, checkpoints, terminal, retry, already_finalized.
  - Archivos creados: ninguno
  - Archivos modificados: `app/adapters/primary/http/routers/payment_validation.py`, `app/application/use_cases/payment_validation_finalize.py`
  - Dependencias: A8, B3, C2
  - Implementación:
    1. En `queue_finalize`: leer control → ExecutionId/ExecutionLogPath.
    2. `reserve_execution_invocation(step=FINALIZE)`.
    3. REQUEST_RECEIVED, QUEUED.
    4. En `_run_finalize_job`: `execution_step` wrapper.
    5. Checkpoints: `FINALIZE_REVIEW_READ`, `FINALIZE_VALIDATION_COMPLETED`, `FINALIZE_HISTORICAL_UPLOADED`, `FINALIZE_SUPPORT_UPLOADED`.
    6. Terminal: SUCCEEDED, FAILED, ALREADY_COMPLETED.
    7. Legacy control sin ExecutionId: Finalize continúa, skip log, warning técnico.
  - Tests:
    - `test_finalize_reads_context_from_control`.
    - `test_finalize_reserves_attempt`.
    - `test_finalize_wrapper_records_events`.
    - `test_finalize_retry_increments_attempt`.
    - `test_finalize_already_finalized`.
    - `test_finalize_legacy_no_execution_id_continues`.
    - `test_finalize_legacy_no_invented_log`.
  - Criterios de aceptación:
    - Finalize sin ExecutionId continúa sin log.
    - Retry incrementa attempt.
    - No inventa log.
  - Riesgos: Control legacy puede tener columnas faltantes.
  - Rollback: Revert de ambos archivos.

---

## Phase D — Notify y Merge

- [ ] D1 — Notify hooks y checkpoints
  - Objetivo: Integrar trazabilidad en los endpoints de notificación.
  - Archivos creados: ninguno
  - Archivos modificados: `app/adapters/primary/http/routers/sharepoint.py`, `app/application/use_cases/send_validar_extractos_notification.py`
  - Dependencias: A8, B3, C2
  - Implementación:
    1. `post_notify_*`: REQUEST_RECEIVED, QUEUED con context desde control.
    2. `_run_notify_*`: execution_step wrapper.
    3. Checkpoints: `NOTIFY_HISTORICAL_READ`, `NOTIFY_EMAIL_SENT`, `NOTIFY_CONTROL_UPDATED`.
    4. Terminal: SUCCEEDED, ALREADY_COMPLETED, COMPLETED_WITH_WARNINGS.
    5. Retry path: nuevo attempt.
  - Tests:
    - `test_notify_records_request_received`.
    - `test_notify_wrapper_events`.
    - `test_notify_already_completed`.
    - `test_notify_control_warning`.
    - `test_notify_retry_increments_attempt`.
  - Criterios de aceptación: Terminal correctamente clasificado. Retry funcional.
  - Riesgos: Dos routers de notify → asegurar ambos integrados.
  - Rollback: Revert de ambos archivos.

---

- [ ] D2 — Merge hooks, artifacts y entities
  - Objetivo: Integrar trazabilidad en Merge con artifacts, entities, y todos los estados terminales.
  - Archivos creados: ninguno
  - Archivos modificados: `app/adapters/primary/http/routers/sharepoint.py`, `app/application/use_cases/merge_composite_validado_pdfs.py`
  - Dependencias: A8, B3, C2
  - Implementación:
    1. `post_merge_*`: REQUEST_RECEIVED, QUEUED.
    2. `_run_merge_*`: execution_step wrapper.
    3. Checkpoints: `MERGE_PREVALIDATION_COMPLETED`, `MERGE_MANIFEST_UPLOADED`, `MERGE_CONTROL_UPDATED`.
    4. Terminal: SUCCEEDED (CONSOLIDADO), PARTIAL (MERGE_PARCIAL), BLOCKED (gate), ALREADY_COMPLETED, FAILED.
    5. Register artifacts: HISTORICAL_FILE, EMAIL_PDF, MERGE_OUTPUT_PDF, MERGE_MANIFEST, PROCESS_CONTROL_FILE.
    6. Register affected_entities por crédito/cliente/tipo.
  - Tests:
    - `test_merge_records_all_event_types`.
    - `test_merge_partial_status`.
    - `test_merge_blocked_status`.
    - `test_merge_artifacts_registered`.
    - `test_merge_entities_registered`.
    - `test_merge_retry_preserves_attempts`.
  - Criterios de aceptación: Todos los estados terminales de Merge cubiertos.
  - Riesgos: Merge es el use case más complejo.
  - Rollback: Revert de ambos archivos.

---

- [ ] D3 — Manifest snapshot
  - Objetivo: Guardar snapshot resumido del manifest dentro del execution log en evento terminal de Merge.
  - Archivos creados: ninguno
  - Archivos modificados: `app/application/use_cases/merge_composite_validado_pdfs.py`
  - Dependencias: D2
  - Implementación:
    1. En terminal de Merge: construir `ManifestSnapshot` desde `manifest_payload`.
    2. Almacenar: `manifest_status`, `eligible_for_dry_run`, `complete_groups_count`, `incomplete_groups_count`, `failed_groups_count`, `outputs` (TruncatedList), `incomplete_groups` (TruncatedList).
    3. Guardar en `execution_log.manifest_snapshot`.
    4. NO cambiar path canónico `merge_manifest_{bank}_{date}.json`.
    5. NO cambiar nombre ni formato del manifest existente.
    6. NO cambiar lectores de Dry-run ni Apply.
  - Tests:
    - `test_manifest_snapshot_populated_on_merge_terminal`.
    - `test_canonical_manifest_path_unchanged`.
    - `test_snapshot_truncates_large_lists`.
    - `test_dry_run_readers_unaffected`.
  - Criterios de aceptación: Manifest canónico intacto. Snapshot es copia resumida.
  - Riesgos: Ninguno si no se toca el manifest.
  - Rollback: Revert.

---

- [ ] D4 — Checkpoints con checkpoint_key para grupos
  - Objetivo: Emitir checkpoint_key diferenciado por crédito/grupo para evitar deduplicación incorrecta.
  - Archivos creados: ninguno
  - Archivos modificados: `app/application/use_cases/merge_composite_validado_pdfs.py`
  - Dependencias: D2
  - Implementación:
    1. `trace.checkpoint(MERGE_MANIFEST_UPLOADED, checkpoint_key=f"group_{group_id}")`.
    2. `event_id` incluye `checkpoint_label + checkpoint_key`.
    3. Dos créditos con mismo label → dos eventos distintos.
  - Tests:
    - `test_two_groups_same_label_different_events`.
    - `test_checkpoint_key_in_event_id`.
  - Criterios de aceptación: No deduplicar eventos legítimos de grupos distintos.
  - Riesgos: Ninguno.
  - Rollback: Revert.

---

## Phase E — Dry-run y Apply

- [ ] E1 — Dry-run independiente
  - Objetivo: Integrar trazabilidad en el endpoint independiente de Dry-run.
  - Archivos creados: ninguno
  - Archivos modificados: `app/adapters/primary/http/routers/payment_validation.py`, `app/application/use_cases/amortization_fill_dry_run.py`
  - Dependencias: A8, B3, C2
  - Implementación:
    1. `step=DRY_RUN`, `substep=None`.
    2. REQUEST_RECEIVED, QUEUED, STARTED, terminal.
    3. Checkpoints: `DRY_RUN_MANIFEST_READ`, `DRY_RUN_PLANNING_COMPLETED`.
    4. Terminal: SUCCEEDED, COMPLETED_WITH_WARNINGS, BLOCKED.
  - Tests:
    - `test_dry_run_records_events`.
    - `test_dry_run_blocked_terminal`.
    - `test_dry_run_step_substep_correct`.
  - Criterios de aceptación: step=DRY_RUN, substep=null.
  - Riesgos: Ninguno.
  - Rollback: Revert.

---

- [ ] E2 — Dry-run interno de Apply
  - Objetivo: Registrar el dry-run interno de Apply con `step=APPLY, substep=DRY_RUN`.
  - Archivos creados: ninguno
  - Archivos modificados: `app/application/use_cases/amortization_fill_apply.py`
  - Dependencias: E1
  - Implementación:
    1. Checkpoint `APPLY_INTERNAL_DRY_RUN_COMPLETED` con `substep=DRY_RUN`.
    2. No confundir con endpoint independiente.
  - Tests:
    - `test_apply_internal_dry_run_substep`.
    - `test_not_confused_with_independent_dry_run`.
  - Criterios de aceptación: substep=DRY_RUN dentro de step=APPLY.
  - Riesgos: Ninguno.
  - Rollback: Revert.

---

- [ ] E3 — Apply checkpoints por crédito
  - Objetivo: Registrar checkpoints per-credit con checkpoint_key en Apply.
  - Archivos creados: ninguno
  - Archivos modificados: `app/adapters/primary/http/routers/payment_validation.py`, `app/application/use_cases/amortization_fill_apply.py`
  - Dependencias: A8, E2
  - Implementación:
    1. `APPLY_TABLE_UPLOADED` con `checkpoint_key=f"{client}|{credit}"`.
    2. `APPLY_TABLE_VERIFIED`, `APPLY_IBR_WRITTEN`, `APPLY_ACCOUNTING_PDF_MOVED`, `APPLY_CONTROL_UPDATED`.
    3. Cada evento repetible con checkpoint_key único.
  - Tests:
    - `test_apply_checkpoint_per_credit`.
    - `test_two_credits_different_events`.
    - `test_checkpoint_key_format`.
  - Criterios de aceptación: Eventos por crédito no deduplicados entre sí.
  - Riesgos: Volumen de eventos alto. Mitigado: events nunca truncados.
  - Rollback: Revert.

---

- [ ] E4 — Apply terminal
  - Objetivo: Clasificar todos los terminales de Apply correctamente.
  - Archivos creados: ninguno
  - Archivos modificados: `app/adapters/primary/http/routers/payment_validation.py`, `app/application/use_cases/amortization_fill_apply.py`
  - Dependencias: A8, E3
  - Implementación:
    1. Terminal: SUCCEEDED, PARTIAL, BLOCKED, FAILED, ALREADY_COMPLETED.
    2. Retry después de interrupción: nuevo attempt.
    3. Post-completion events: no reabrir ejecución, incrementar `post_completion_events_count`.
    4. `ensure_execution_control_link` al finalizar (best-effort retry).
  - Tests:
    - `test_apply_succeeded`.
    - `test_apply_partial`.
    - `test_apply_blocked`.
    - `test_apply_already_completed`.
    - `test_apply_retry_after_interrupt`.
    - `test_apply_post_completion_no_reopen`.
    - `test_apply_retries_control_link`.
  - Criterios de aceptación: Todos los terminales clasificados. Post-completion no reabre.
  - Riesgos: Edge case de post-completion con diferentes execution_id.
  - Rollback: Revert.

---

## Phase F — Polling y reparación terminal

- [ ] F1 — Poll metadata acumulador e implementación de `record_job_poll`
  - Objetivo: Implementar la lógica completa de acumulación de poll metadata en memoria y la función `record_job_poll` (cuyo stub se definió en A7).
  - Archivos creados: `tests/test_execution_polling.py`
  - Archivos modificados: `app/application/services/execution_run_log.py` (completar stub), `app/adapters/primary/http/routers/payment_validation.py`, `app/adapters/primary/http/routers/sharepoint.py`
  - Dependencias: A7 (stubs definidos), C1 (trace keys)
  - Implementación:
    1. `_poll_meta` en job dict: `{poll_count, last_status, statuses: {status: {count, first_at, last_at}}}`.
    2. Incrementar en cada GET /jobs/{id}.
    3. Completar implementación de `record_job_poll(graph, context, job_snapshot, poll_meta)` (stub de A7).
    4. Persistir al execution log solo:
       - Primera consulta.
       - Cambio de estado.
       - Estado terminal.
       - Error HTTP.
       - Cada quinta consulta idéntica (`poll_count % 5 == 0`).
  - Tests:
    - `test_first_poll_persists`.
    - `test_status_change_persists`.
    - `test_fifth_identical_persists`.
    - `test_intermediate_identical_does_not_persist`.
    - `test_terminal_persists`.
    - `test_poll_meta_resets_after_restart`.
  - Criterios de aceptación: No write en cada poll. Transiciones capturadas.
  - Riesgos: `_poll_meta` lost on restart (aceptable per spec).
  - Rollback: Revert.

---

- [ ] F2 — Presupuesto de latencia de polling
  - Objetivo: Implementar timeouts best-effort para que logging no retrase GET /jobs/{id}.
  - Archivos creados: ninguno
  - Archivos modificados: `app/application/services/execution_run_log.py`
  - Dependencias: F1
  - Implementación:
    1. Primera/cambio: max 500ms `asyncio.wait_for`.
    2. Quinta idéntica: max 200ms.
    3. Terminal: max 2s.
    4. Si timeout: ignorar error, devolver estado real del job.
    5. Ningún GET /jobs/{id} suma más de 2s overhead.
  - Tests:
    - `test_slow_graph_does_not_block_poll_response` — fake Graph con delay.
    - `test_timeout_returns_real_job_status`.
    - `test_terminal_gets_longer_budget`.
  - Criterios de aceptación: Presupuesto de latencia respetado. Job status siempre devuelto.
  - Riesgos: `asyncio.wait_for` timeout cleanup.
  - Rollback: Revert.

---

- [ ] F3 — Terminal repair por prioridad e implementación de `ensure_terminal_job_event_logged`
  - Objetivo: Implementar la lógica completa de reparación de eventos terminales (cuyo stub se definió en A7) sin fabricar datos.
  - Archivos creados: `tests/test_terminal_repair.py`
  - Archivos modificados: `app/application/services/execution_run_log.py` (completar stub)
  - Dependencias: A7 (stub definido), C1 (trace keys), F1 (poll meta)
  - Implementación:
    1. Completar implementación de `ensure_terminal_job_event_logged(graph, job)` (stub de A7).
    2. Si job no terminal: return.
    3. Si terminal ya existe para job_id+invocation_id: return.
    4. Resolución de attempt:
       - Prioridad 1: metadata del job (trace fields).
       - Prioridad 2: evento existente con mismo job_id.
       - Prioridad 3: QUEUED/STARTED con mismo invocation_id.
    5. Sin correlación segura: NO crear terminal, warning técnico, `possible_gaps=true`.
    6. NO fabricar attempt, invocation_id ni eventos históricos.
  - Tests:
    - `test_repair_from_job_metadata`.
    - `test_repair_from_event_job_id`.
    - `test_repair_from_queued_invocation`.
    - `test_no_repair_when_uncorrelatable`.
    - `test_uncorrelatable_sets_possible_gaps`.
    - `test_no_fabricated_attempt`.
    - `test_already_terminal_no_duplicate`.
  - Criterios de aceptación: Nunca fabricar datos. Correlación o nothing.
  - Riesgos: Pérdida de eventos si metadata perdida.
  - Rollback: Revert.

---

- [ ] F4 — Job stores unificados para repair
  - Objetivo: Verificar que terminal repair funciona con ambos stores (JobManager + `_validation_jobs`).
  - Archivos creados: ninguno
  - Archivos modificados: `app/adapters/primary/http/routers/payment_validation.py`, `app/adapters/primary/http/routers/sharepoint.py`
  - Dependencias: F3
  - Implementación:
    1. Helper compartido que recibe el job dict del router local.
    2. `get_job_status`: llamar `ensure_terminal_job_event_logged` con job del store local.
    3. `*_job_status` en sharepoint: misma lógica.
    4. No duplicar la lógica de repair entre routers.
  - Tests:
    - `test_repair_works_from_job_manager`.
    - `test_repair_works_from_validation_jobs`.
    - `test_helper_shared_not_duplicated`.
  - Criterios de aceptación: Un solo helper. Ambos stores cubiertos.
  - Riesgos: Ninguno.
  - Rollback: Revert.

---

- [ ] F5 — 404, 409, 422 edge cases
  - Objetivo: Manejar edge cases HTTP con correlación segura o sin ella.
  - Archivos creados: ninguno
  - Archivos modificados: `app/adapters/primary/http/routers/payment_validation.py`, `app/adapters/primary/http/routers/sharepoint.py`
  - Dependencias: C3
  - Implementación:
    1. 409 en Generate con ExecutionId activo: MAY append REQUEST_REJECTED.
    2. 422 antes de router: log request_id a Render, NO crear execution_id.
    3. 404 job: Render warning, MAY warn execution log si metadata permite.
  - Tests:
    - `test_409_appends_rejected_to_active_log`.
    - `test_422_no_execution_id_created`.
    - `test_404_job_render_warning_only`.
  - Criterios de aceptación: No correlación insegura. No execution_id fabricado.
  - Riesgos: Ninguno.
  - Rollback: Revert.

---

## Phase G — Verificación preproducción y rollout

- [ ] G1 — Suite completa
  - Objetivo: Ejecutar toda la suite de tests y compilación.
  - Archivos creados: ninguno
  - Archivos modificados: ninguno
  - Dependencias: A–F completas
  - Implementación:
    1. `python -m pytest -q`.
    2. `python -m compileall app -q`.
    3. Resolver cualquier fallo.
  - Tests: Suite completa.
  - Criterios de aceptación: 0 fallos. 0 errores de compilación.
  - Riesgos: Fallos por interacción entre fases.
  - Rollback: N/A.

---

- [ ] G2 — Regresión con flag false
  - Objetivo: Confirmar que con flag false el sistema actual es idéntico.
  - Archivos creados: `tests/test_execution_run_log_regression.py`
  - Archivos modificados: ninguno
  - Dependencias: G1
  - Implementación:
    1. Flujo completo mock con `EXECUTION_RUN_LOG_ENABLED=false`.
    2. Verificar: mismos archivos, mismos estados, mismos contratos, mismos resultados.
    3. Mock assert: cero llamadas a Graph por logging.
    4. PA sin cambios.
  - Tests:
    - `test_flag_false_identical_generate_result`.
    - `test_flag_false_identical_finalize_result`.
    - `test_flag_false_zero_graph_logging_calls`.
    - `test_flag_false_control_unchanged`.
    - `test_flag_false_manifest_unchanged`.
  - Criterios de aceptación: Comportamiento bit-idéntico con flag false.
  - Riesgos: Ninguno.
  - Rollback: N/A.

---

- [ ] G3 — Smoke flag true: Banco Bogotá
  - Objetivo: Ejecutar flujo completo con flag true para Banco Bogotá y verificar JSON.
  - Archivos creados: ninguno
  - Archivos modificados: ninguno
  - Dependencias: G2
  - Implementación:
    1. Flag true en env de validación.
    2. Generate → Finalize → Notify → Merge → Dry-run → Apply.
    3. Verificar JSON en 04 LOGS: schema v1, events completos, summary SUCCEEDED, artifacts, entities.
    4. Verificar: un execution_id, un archivo, attempts correctos.
  - Tests: Smoke test manual o semi-automatizado.
  - Criterios de aceptación: JSON completo y válido. Un archivo. Summary SUCCEEDED.
  - Riesgos: Dependencia de ambiente de validación.
  - Rollback: Flag false.

---

- [ ] G4 — Smoke flag true: Bancolombia
  - Objetivo: Mismos criterios que G3 para Bancolombia.
  - Archivos creados: ninguno
  - Archivos modificados: ninguno
  - Dependencias: G3
  - Implementación: Misma que G3 con banco_bancolombia.
  - Tests: Smoke test.
  - Criterios de aceptación: Mismos que G3. Lock independiente del otro banco.
  - Riesgos: Mismos que G3.
  - Rollback: Flag false.

---

- [ ] G5 — Reintentos reales
  - Objetivo: Verificar que reintentos mantienen execution_id y generan attempts distintos.
  - Archivos creados: ninguno
  - Archivos modificados: ninguno
  - Dependencias: G3
  - Implementación:
    1. Finalize fallido → corregir → reintentar: mismo execution_id, attempt 2.
    2. Merge parcial → corregir → reintentar: mismo execution_id.
    3. Apply bloqueado → corregir → reintentar: mismo execution_id.
    4. Verificar: un solo archivo JSON, attempts distintos, events de ambos intentos.
  - Tests: Semi-automatizado con mocks o en validación.
  - Criterios de aceptación: Mismo execution_id. Attempts incrementales. Un archivo.
  - Riesgos: Requiere capacidad de simular fallos.
  - Rollback: Flag false.

---

- [ ] G6 — Fallos del sistema de log
  - Objetivo: Simular fallos de logging y verificar resiliencia.
  - Archivos creados: ninguno
  - Archivos modificados: ninguno
  - Dependencias: G2
  - Implementación:
    1. Simular timeout Graph en logging → job sigue.
    2. Simular 412 → merge y retry.
    3. Simular 404 en read → reconstruct.
    4. Simular JSON corrupto → tratar como vacío + possible_gaps.
    5. Simular control sin vínculo → ensure_control_link retry.
    6. Simular reinicio entre STARTED y terminal → attempt visible, nuevo attempt en retry.
  - Tests: Con fakes lentos o que fallan.
  - Criterios de aceptación: Business flow nunca interrumpido por logging. Resiliencia verificada.
  - Riesgos: Ninguno.
  - Rollback: N/A.

---

- [ ] G7 — Verificación Power Automate
  - Objetivo: Confirmar que PA no requiere cambios.
  - Archivos creados: ninguno
  - Archivos modificados: ninguno
  - Dependencias: G3
  - Implementación:
    1. Verificar: no se cambiaron bodies de request.
    2. No se agregaron headers obligatorios.
    3. Polling sin cambios.
    4. Parse JSON tolera campos raíz opcionales (`additionalProperties`).
    5. Mensajes y correos funcionando.
  - Tests: Smoke test con PA o verificación manual.
  - Criterios de aceptación: PA sin ninguna modificación.
  - Riesgos: Parse JSON estricto en PA podría rechazar campos nuevos. Mitigar: campos opcionales.
  - Rollback: Flag false elimina campos opcionales.

---

- [ ] G8 — Documentación de rollout
  - Objetivo: Documentar el plan de rollout y rollback.
  - Archivos creados: documentación dentro de openspec o README
  - Archivos modificados: ninguno productivo
  - Dependencias: G1-G7
  - Implementación:
    1. Deploy con `EXECUTION_RUN_LOG_ENABLED=false`.
    2. Ejecutar suite de regresión.
    3. Flag true en ambiente de validación.
    4. Smoke por banco.
    5. Habilitación en producción tras sign-off.
    6. Rollback: `false` + redeploy. JSON files inertes.
  - Tests: N/A.
  - Criterios de aceptación: Documento escrito. Plan claro. Rollback en 1 paso.
  - Riesgos: Ninguno.
  - Rollback: N/A.

---

## Dependencias

```
A0 ──→ A3
A1 + A2 ──→ A4 ──→ A5
A1 ──→ A6
A3 + A4 + A5 + A6 ──→ A7
A7 ──→ A8

A0 + A3 + B4 ──→ B5
B2 ──→ B3 ──→ B6

C2 depende de A3 + A7 (escritura condicional)
A7 + B1 + B2 + B3 + B5 + B6 + C1 + C2 ──→ C3
C3 ──→ C4 ──→ C5
C3 + C2 ──→ C6

C6 ──→ D1, D2
D2 ──→ D3, D4
D1 + D2 ──→ E1, E2
E2 ──→ E3 ──→ E4

C + D + E ──→ F1 ──→ F2, F3
F3 ──→ F4, F5

A–F ──→ G1 ──→ G2 ──→ G3 ──→ G4 ──→ G5
G2 ──→ G6
G3 ──→ G7
G1–G7 ──→ G8
```

**Compuertas de seguridad:**
- **A0** DEBE completarse (y verificar If-Match) antes de A3, B5, y cualquier tarea que use eTag.
- **B1** DEBE completarse antes de cualquier hook productivo.
- **G2** DEBE completarse antes de G3 (no activar flag sin regresión).
- **C2** depende de A3 Y A7 (necesita escritura condicional para persistir attempt).

---

## Política de commits

Commits por entregable coherente. Cada commit mantiene `pytest -q` y `compileall` verdes:

```
1. feat(traceability): add execution log models and sanitizer (A1 + A2)
2. feat(graph): verify If-Match support on PUT /content (A0)
3. feat(graph): support conditional file writes with eTag (A3)
4. feat(traceability): add execution log schema, merge, and summary (A4 + A5 + A6)
5. feat(traceability): add execution log persistence and step wrapper (A7 + A8)
6. feat(traceability): add feature flag and control column extensions (B1 + B2 + B3)
7. feat(traceability): add reservation infrastructure and bank lock (B4 + B5 + B6)
8. feat(traceability): trace generate and finalize jobs (C1 + C2 + C3 + C4 + C5 + C6)
9. feat(traceability): trace notify and merge jobs (D1 + D2 + D3 + D4)
10. feat(traceability): trace dry-run and apply jobs (E1 + E2 + E3 + E4)
11. feat(traceability): add job polling and terminal repair (F1 + F2 + F3 + F4 + F5)
12. test(traceability): preproduction regression and smoke tests (G1 + G2 + G3 + G4 + G5 + G6 + G7)
13. docs(traceability): rollout and rollback plan (G8)
```

---

## Criterio de finalización por fase

Una fase NO puede cerrarse si:

1. `python -m pytest -q` falla.
2. `python -m compileall app -q` falla.
3. Flag false cambia el comportamiento.
4. Se modifica contrato HTTP/JSON no autorizado.
5. Falla una prueba de idempotencia.
6. Se detecta doble efecto financiero.
7. La tarea A0 no confirma If-Match para fases que dependen de eTag.

---

## Compuertas de implementación (checklist pre-apply)

Antes de `/sdd-apply` confirmar:

- [ ] A0 está definido como spike ejecutable con `@pytest.mark.graph_integration`.
- [ ] If-Match real todavía debe verificarse (A0 es la primera tarea).
- [ ] Attempt se calcula con `1 + max(attempt)`, no con `count(STARTED)`.
- [ ] Matriz WAITING coincide exactamente con el diseño aprobado (6 filas).
- [ ] B5 distingue reserva bloqueante de logging no bloqueante.
- [ ] A7 define stubs de polling; F1–F3 los implementa. Sin duplicación.
- [ ] Conteo total: 42 tareas en 7 fases.
