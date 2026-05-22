## Exploration: contratos HTTP Power Automate tras Estado Pago + Validar Pago

### Current State

- **Cola payment-validation** (`POST /graph/sharepoint/payment-validation/generate|finalize/queue`): la respuesta inmediata sigue siendo `{"job_id": "<uuid>", "status": "queued"}` (sin cambios de forma respecto a este diseño).
- **Consulta de job** (`GET /graph/sharepoint/payment-validation/jobs/{job_id}`): el documento se enriquece con `user_message`, `next_action` y `severity` cuando `status` es `completed` o `failed` (`job_status_enrichment.enrich_job_for_http_response`). Estructura de nivel superior: `job_id`, `type`, `status`, `queued_at`, `updated_at`, y según estado `result` o `error`.
- **Generate (resultado en job completado)**: `generate_payment_validation` devuelve las mismas claves de alto nivel que antes: `process_id`, `validation_file`, `validation_file_path`, `validation_file_url`, `summary` (con `pagos_banco`, `errores`, `conditional_formatting`). El router añade `process_date` e `elapsed_ms` dentro de `result`.
- **Finalize (resultado en job completado)**: además de las claves existentes (`status`, `historical_file_path`, `historical_file_url`, `secretary_*`, `validated_rows`, `amortization_updated`, `bank_cleaned`, `history_file`, `validation_file`, `validation_file_path`, `process_date`), **se añadió** `audit_warnings` (lista/array) cuando el append de auditoría advierte algo.
- **Errores de job failed**: sigue el patrón `error: { type, message, ... }` y enriquecimiento con `error_code`, `user_message`, `next_action`, `severity`. Códigos nuevos de excepción (p. ej. `invalid_estado_pago`) pueden no estar en `_FINALIZE_MESSAGES` y caer en mensaje genérico (`_UNKNOWN_USER`).
- **Otros contratos documentados para PA** (`docs/power-automate-flow.md`, `app/models.py`): `POST /parse-excel` / `ParseExcelResponse` no están enlazados al caso de revisión de pagos; no cambian con esta change a nivel de modelo Pydantic.

### Affected Areas

- `app/adapters/primary/http/routers/payment_validation.py` — cuerpos `GenerateRequest` / `FinalizeRequest` sin cambios de campos; misma forma de `result` en job.
- `app/application/use_cases/payment_validation_finalize.py` — nueva clave `audit_warnings` en el dict de retorno.
- `app/application/job_status_enrichment.py` — textos de usuario/next_action para finalize/notify/merge; mensaje de `unresolved_status` aún describe tokens legacy (desalineación UX vs reglas nuevas, no cambio de nombres de campos JSON).
- `docs/power-automate-flow.md` — describe solo parse-excel, no el job de payment-validation.
- **SharePoint / Excel**: libros generados o finalizados tienen columnas renombradas y columna nueva (`Estado Pago`, `Validar Pago`). Cualquier flujo PA que lea el xlsx por **índice de columna fijo** o busque el texto `VALIDAR` en la columna de estado queda afectado **fuera** del JSON de la API.

### Approaches

1. **Considerar la API “compatible hacia atrás” si los flujos solo leen claves existentes** — los clientes ignoran `audit_warnings`.
   - Pros: sin cambios en PA para quien usa solo `historical_file_path`, URLs y contadores.
   - Cons: si un flujo usa un esquema JSON “cerrado” (rechaza propiedades extra), el nuevo campo puede romper el parse.
   - Effort: Low (validación manual de cada flujo).

2. **Actualizar esquemas Parse JSON / Compose en Power Automate** — añadir `audit_warnings` opcional y revisar ramas por `error_code`.
   - Pros: alineación explícita con producción; manejo de advertencias de auditoría.
   - Cons: trabajo en cada ambiente PA.
   - Effort: Medium.

3. **Alinear Excel y flujos que lean tablas SharePoint** — usar nombres de columna o tabla en lugar de letras fijas; actualizar condiciones que dependan del token `VALIDAR` en columna Estado.
   - Pros: evita roturas silenciosas en procesos no-HTTP.
   - Cons: puede requerir redesplegar muchos flujos.
   - Effort: Medium/High según complejidad.

### Recommendation

Tratar el contrato HTTP del job **principalmente additive**: **no se eliminaron claves** conocidas de `result` de generate/finalize; **sí se agregó** `audit_warnings`. Los **cuerpos de entrada** de queue no cambian. El riesgo mayor para “PA falla” suele ser **Parse JSON con esquema estricto** o **automatizaciones que lean el Excel por posición o por el texto `VALIDAR`**, no el JSON de `GET /jobs`. Conviene revisar flujos que pasen `historical_file_path` a Notify/Merge (sigue igual) y validar que ningún schema JSON tenga `additionalProperties: false` sobre `result`.

### Risks

- Flujos con validación rígida del payload de `completed` fallan al aparecer `audit_warnings`.
- Nuevos `ValueError` (`invalid_estado_pago`, etc.) generan `error_code` no mapeado → mensaje genérico al usuario; el flujo puede depender de códigos concretos.
- Procesos que no son HTTP pero manipulan el libro de revisión/histórico pueden romperse por **estructura de columnas**.

### Ready for Proposal

**Yes** — para documentación/operación: añadir a propuesta o runbook que `GET .../jobs/{id}` en finalize completado puede incluir `audit_warnings`, y que compatibilidad con PA asume esquema flexible o actualización explícita.

---

## Exploration: HTTP 404 al leer `PAGOS_PENDIENTES.xlsx` durante Finalize

### Current State

- Tras generar histórico/auditoría, `payment_validation_finalize` llama siempre a `client.get_bytes` contra `GRAPH_PENDING_PAYMENTS_FILE_PATH` (`payment_validation_finalize.py` ~l.1198–1201).
- La lógica siguiente usa `if p_bytes:` para **solo entonces** abrir el Excel de pendientes, mover filas del histórico interno del libro, etc.; es decir, el diseño asume que “sin bytes” no hay trabajo de pendientes.
- **No** se captura `httpx.HTTPStatusError` con código 404: si el archivo **no existe** en la ruta del drive, Graph devuelve 404 y la excepción **aborta todo Finalize** antes de llegar al `if p_bytes`.
- Ese tipo de error **no** está en `_FINALIZE_MESSAGES`; `_error_code_from_generate_or_finalize_message` genera un `error_code` largo/truncado a partir del mensaje, y el usuario ve **«Ocurrió un error durante el proceso»** (`_UNKNOWN_USER`).

### Affected Areas

- `app/application/use_cases/payment_validation_finalize.py` — descarga obligatoria del libro de pendientes; falta rama 404 → bytes vacíos (patrón ya usado en `audit_tracking_finalize.py`, `setup_merge_control_workbook.py`, etc.).
- `app/application/job_status_enrichment.py` — excepciones `HTTPStatusError` genéricas en jobs `finalize` no se mapean a mensajes accionables.

### Approaches

1. **Operación / datos** — Crear en SharePoint el archivo `PAGOS_PENDIENTES.xlsx` en la ruta exacta de `GRAPH_PENDING_PAYMENTS_FILE_PATH` (coherente con `.env.example`: `…/00 CONTROL/PAGOS_PENDIENTES.xlsx`), con hoja `Pendientes` mínima si aún no existe.
   - Pros: sin despliegue de código.
   - Cons: cada entorno debe tener el archivo; sigue frágil ante errores de tipeo de ruta.
   - Effort: Low.

2. **Código** — Tras `resolve_sharepoint_path`, envolver `get_bytes` en `try/except HTTPStatusError` y si `status_code == 404`: usar `p_bytes = b""` y continuar (opcional: log / `audit_warnings`).
   - Pros: Finalize puede completar histórico y secretaría aunque el libro de pendientes aún no exista; alinea el comportamiento con el `if p_bytes`.
   - Cons: no actualizará/moverá pendientes hasta que el archivo exista (comportamiento esperado si no hay libro).
   - Effort: Low/Medium.

3. **Enriquecimiento de errores** — Añadir mapeo para `HTTPStatusError` 404 en lecturas de pendientes (o código dedicado `pending_payments_file_not_found`) con mensaje claro y `next_action` (crear archivo o revisar `GRAPH_PENDING_PAYMENTS_FILE_PATH`).
   - Pros: mejor UX aunque no se implemente 2).
   - Effort: Low.

### Recommendation

Combinar **2) + 3)** para robustez y mensajes claros, y en paralelo **validar en Render** que `GRAPH_PENDING_PAYMENTS_FILE_PATH` coincide con la ubicación real y que el archivo existe (o se crea una plantilla una vez).

### Risks

- Tratar 404 como “sin pendientes” puede ocultar un error de path (typo); compensar con log y/o advertencia en respuesta.
- Si más adelante es obligatorio tener el libro para cumplimiento operativo, documentarlo en runbook.

### Ready for Proposal

**Yes** — documentar criterio: «¿Finalize debe poder completar sin `PAGOS_PENDIENTES.xlsx`?» Si sí → implementar 404 → vacío; si no → solo mensajes y validación previa.

---

## Exploration: Merge composite PDFs vs lógica de extractos por crédito (Generate / Ruta / Notify)

### Current State

**Origen del dato “extracto” en Merge**

- `merge_composite_validado_pdfs` lee el **histórico** (`cartera_validada_*.xlsx` vía `control_merge_pdfs.xlsx` fila 2), hoja **Distribución**.
- Los PDFs de extracto no se “descubren” otra vez con la lógica de carpetas de Generate: se toman de la columna **Ruta** (y filtros por estado / Validar Pago vía `distrib_row_included_for_validar_extractos`), usando **`_collect_pdf_paths_from_ruta_cell`** definida en `send_validar_extractos_notification.py`.
- Esa función resuelve **fragmentos** de celda (coma/salto/`;`), prueba `.pdf` directo, carpeta → PDFs hijos, etc. Es la **misma** resolución semántica que **Notify** para adjuntos.

**Lógica rica en Generate (solo al generar el libro de revisión)**

- `payment_validation_generate._resolve_extract_pdf_pool` prioriza subcarpeta **EXTRACTOS** si existe y hay PDFs “extracto” válidos; si no, hace fallback a PDFs en la **raíz de la carpeta del crédito**.
- `_select_extract_by_max_fecha_limite_v2` elige un PDF entre candidatos por **fecha límite** en el contenido; hay códigos de error por empates, fechas ilegibles, etc.
- Lo que termina en **Link extracto** / **Ruta** en Distribución es el resultado de ese pipeline (ruta interna al drive del PDF elegido).

**Lógica añadida solo en Merge (asientos contables)**

- Tras obtener cada ruta de extracto `ep` concreta (archivo `.pdf`), Merge calcula la carpeta de asientos así:  
  `asientos_dir = parent(ep) + "/" + GRAPH_ASIENTOS_CONTABLES_FOLDER_NAME`  
  (`merge_composite_validado_pdfs._prevalidate_id_pago_group`, líneas ~285–300 del módulo).
- Es decir: el padre **inmediato** del archivo PDF de extracto es el contenedor; se **asume** que **ASIENTOS CONTABLES** cuelga de **ese** mismo nivel.

**Consecuencia de desalineación**

- Si el extracto está en `…/CREDITO # n/EXTRACTOS/Extracto.pdf`, `parent(ep)` = `…/EXTRACTOS` y Merge busca `…/EXTRACTOS/ASIENTOS CONTABLES`.
- En muchos arreglos reales (incl. capturas de usuario), **ASIENTOS CONTABLES** está en `…/CREDITO # n/ASIENTOS CONTABLES` (hermano de **EXTRACTOS**, no hijo). Generate ya pudo escribir en **Ruta** la ruta correcta del PDF bajo **EXTRACTOS**, pero Merge **no** sube un nivel para buscar asientos → **404** al listar carpeta o **skip** (`asiento_folder_list_failed`).
- Si el extracto está directamente bajo la carpeta del crédito (`…/CREDITO # n/Extracto.pdf`), `parent(ep)` = carpeta del crédito y `…/CREDITO # n/ASIENTOS CONTABLES` **coincide** con el layout típico → suele funcionar (caso GEOEXCON en logs).

**Filas incluidas en Merge**

- Merge filtra con `distrib_row_included_for_validar_extractos` (Validar Pago + Estado Pago / env `GRAPH_VALIDAR_EXTRACTO_ESTADO_CONTAINS` u orientación a tokens legacy). Debe estar **alineado** con la política de negocio de “qué líneas entran en correo/merge”; no replica la **búsqueda** de extractos de Generate, solo lee lo ya materializado en **Ruta**.

### Affected Areas

- `app/application/use_cases/merge_composite_validado_pdfs.py` — construcción de `asientos_dir_ep` solo desde `parent(ep)`; punto débil frente a layouts con **EXTRACTOS** y **ASIENTOS** al mismo nivel.
- `app/application/use_cases/send_validar_extractos_notification.py` — `_collect_pdf_paths_from_ruta_cell` / `_resolve_pdf_paths_for_ruta_fragment`: **alineado** entre Notify y Merge para **qué PDFs** son extractos.
- `app/application/use_cases/payment_validation_generate.py` — `_resolve_extract_pdf_pool` y selección por fecha: **no** reutilizado en Merge; solo afecta indirectamente vía lo que Finalize deja en **Ruta**.

### Approaches

1. **Mantener SharePoint** — Forzar carpeta `EXTRACTOS/ASIENTOS CONTABLES` donde Merge la espera.  
   - Pros: sin cambio de código.  
   - Cons: duplica estructura, fricción operativa.  
   - Effort: Low (operación).

2. **Candidatos de carpeta asientos en código** — Tras conocer `ep`, probar en orden:  
   `parent(ep)/ASIENTOS`, `parent(parent(ep))/ASIENTOS` si el segmento es `EXTRACTOS`, y opcionalmente rutas desde env.  
   - Pros: alinea Merge con layouts “EXTRACTOS hermano de ASIENTOS”.  
   - Cons: hay que definir orden y reglas para no tomar carpeta equivocada.  
   - Effort: Medium.

3. **Columna explícita o metadata** — Guardar en histórico la ruta de carpeta asientos (nueva columna o convención en Ruta).  
   - Pros: máxima flexibilidad para “muchos casos”.  
   - Cons: cambio de libro, Generate/Finalize, tests.  
   - Effort: High.

### Recommendation

Tratar **2)** como primer paso: amplia el descubrimiento de **ASIENTOS CONTABLES** sin reimplementar todo el árbol de Generate, porque **el extracto ya está bien resuelto** en **Ruta**; el fallo observado es casi siempre la **heurística única** `parent(extracto)/ASIENTOS`. Documentar que Merge **no** vuelve a ejecutar la misma máquina de estados de carpetas que Generate para elegir extracto — solo **consume** el PDF que Finalize ya fijó.

### Risks

- Varias carpetas candidatas pueden existir; hace falta prioridad clara y logs.
- Cambiar heurística puede afectar créditos que hoy dependían de `EXTRACTOS/ASIENTOS` si existiera duplicada.

### Ready for Proposal

**Yes** — proponer cambio acotado “fallback carpeta asientos” + pruebas con rutas `…/EXTRACTOS/foo.pdf` y `…/CREDITO/foo.pdf`; opcionalmente ticket para **3)** si el negocio requiere casos que no cubra ninguna heurística.
