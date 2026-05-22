# Design: Estado Pago + Validar Pago + Excels de auditoría

## Technical Approach

**Una sola columna nueva:** `Validar Pago` (`SI` / `NO`), interruptor para Finalize / Notify / Merge.

La columna que hoy se llama **`Estado`** pasa a **`Estado Pago`** (mismo slot en la tabla o desplazar una posición si `Validar Pago` se inserta “a la derecha” según spec). Su contenido es **solo** el conjunto cerrado:

`ADELANTADO | ATRASADO | INCOMPLETO | NORMAL | REVISION_MANUAL`.

Los valores actuales **`VALIDAR`, `REPROGRAMAR`, `VALIDAR_PARCIAL`, `NO_VALIDAR`, `PENDIENTE_MORA`** y cualquier otro legacy **no** pueden permanecer en esa celda: en libros nuevos Generate escribe únicamente los cinco; en Finalize se **rechaza** fila con Estado Pago inválido (`invalid_estado_pago`). Libros antiguos: **tabla de migración** determinista a par `(Estado Pago, Validar Pago)` (spec) antes de procesar.

**No** se introduce columna “Acción validación”: la decisión operativa anterior (¿procesar?, ¿reprogramar?) se expresa con **Validar Pago** + **Estado Pago** + **Observación** según reglas de negocio documentadas.

**Excels de auditoría** (`pagos_adelantados.xlsx` / `pagos_incompletos.xlsx`): la intención sigue siendo append deduplicado por **ID Pago** y rutas conocidas; el **contrato físico** del libro queda alineado al módulo `setup_audit_tracking_workbooks.py` (implementado antes del cierre de esta change en Finalize):

- **Creación one-shot:** `POST /graph/sharepoint/payment-validation/setup/audit-tracking-workbooks` (idempotente; no pisa archivos existentes).
- **Carpeta base:** `GRAPH_PAYMENT_VALIDATION_CONTROL_PATH` o, si falta, la misma ruta por defecto que el control Merge (`…/00 CONTROL`). Override opcional: `GRAPH_AUDIT_PAGOS_ADELANTADOS_PATH` / `GRAPH_AUDIT_PAGOS_INCOMPLETOS_PATH`.
- **Hoja:** `Registros`. **Tablas Excel:** `tblPagosAdelantados` / `tblPagosIncompletos`. Estilo cabecera = mismo criterio visual que `control_merge_pdfs` (azul `#002060`); **sin** protección de hoja para permitir append programático.
- **Encabezados fila 1 (orden fijo):** Fecha proceso, Fecha hora registro, ID Pago, Cliente, Crédito, Fecha banco, Fecha límite, Monto banco, Valor extracto, Aplicar a extracto, Mora a aplicar, Otros valores, Total aplicado, Saldo por asignar, Estado Pago, Validar Pago, Observación, Ruta, Link extracto, Link carpeta crédito, Archivo histórico.

Finalize debe **añadir filas bajo esa tabla** (o insertar en la tabla) sin renombrar hoja/columnas.

## Architecture Decisions

| Decision | Options | Choice | Rationale |
|----------|---------|--------|-----------|
| ¿Cuántas columnas nuevas? | Una vs dos | **Una**: solo **Validar Pago** | Alcance acordado |
| Columna `Estado` | Conservar nombre vs renombrar | **Renombrar** a **Estado Pago** | Misma función “ubicación” en Distribución, semántica nueva |
| Ubicación `Validar Pago` | Izquierda vs derecha de Estado Pago | **A la derecha de Estado Pago** (requisito usuario previo) | Orden fijado en spec |
| Valores Estado Pago | Mixtos vs cerrado | **Solo 5**; validación en Generate + Finalize + lecturas | APIs limpias |
| Legacy en celda | Coexistir vs prohibir | **Prohibido** en Estado Pago; migración o error | Coherencia auditoría |
| Rol de `VALIDAR` antiguo | Perdido vs reemplazado | Sustituido por **`Validar Pago = SI`** + Estado Pago situacional (p. ej. `NORMAL`) | Pipeline = SI + reglas montos/fechas |
| Fallo migración | Silencioso vs explícito | **Código error** + mensaje usuario | Evita Finalizar con datos incoherentes |
| Notify/Merge | Token `VALIDAR` | **Validar Pago = SI** ∧ filtro `GRAPH_…` sobre **Estado Pago** o columna técnica si se mantiene env (ajuste spec) | Reemplazar dependencia del texto `VALIDAR` en la columna renombrada |
| Bootstrap auditoría | Manual vs API | **API** `setup/audit-tracking-workbooks` + env | Una sola fuente de verdad de columnas y estilo; evita libros vacíos incompatibles con append |
| Carpeta asientos en Merge | Inferir `parent(extracto)` vs ruta en Excel | **Solo `RutaAsientosContables`**; **sin fallback** a heurística legacy | Evita usar carpetas incorrectas por layout antiguo |
| Fallo crear carpeta asientos (Finalize) | Warning vs error | **Error** (falla el job / no publica histórico usable) + código y mensaje enriquecido | No dejar histórico con celdas vacías que Merge interpretaría mal |
| Unidad de crédito en Excel | Solo `webUrl` vs ruta relativa | **RutaUnidadCredito** en Generate | Misma noción que `credit_path` interno; no depende de EXTRACTOS |
| Nombre carpeta asientos | Genérico `ASIENTOS CONTABLES` vs por crédito | **`ASIENTOS CONTABLES CRED {n}`** bajo unidad de crédito | Evita colisiones entre créditos en el mismo cliente |
| Contrato HTTP Generate/Finalize | Ampliar JSON vs solo Excel | **Solo Excel** (columnas ocultas) | Power Automate sin cambios de payload |

## Data Flow

```
Generate
  → escribe Estado Pago ∈ {5} y Validar Pago (default según spec)

Finalize
  → allowlist Estado Pago
  → procesa líneas con Validar Pago = SI según reglas (montos, etc.)
  → append por Estado Pago ADELANTADO / INCOMPLETO a auditoría

Notify/Merge
  → filas: Validar Pago = SI ∧ criterios Estado Pago / rutas / env
```

## Rutas documentales: unidad de crédito y asientos (Merge)

**Problema:** Merge hoy infiere la carpeta de asientos con `parent(ruta_extracto_pdf) + "/ASIENTOS CONTABLES"`, lo que falla cuando el extracto está bajo **EXTRACTOS** y la carpeta de asientos es hermana bajo la unidad de crédito (véase `exploration.md`).

**Solución acordada:** materializar rutas en el Excel (contrato en datos). **No** se alteran los JSON de respuesta de Generate/Finalize/Merge (solo el `.xlsx` de revisión e histórico).

### Columnas (hoja Distribución)

| Columna visible | Constante sugerida | Quién la escribe | Contenido |
|-----------------|-------------------|------------------|-----------|
| **Ruta** (oculta) | `RUTA` | Generate (Finalize puede refrescar extracto) | Ruta relativa al drive del **PDF extracto** elegido |
| **RutaUnidadCredito** (oculta) | `RUTA_UNIDAD_CREDITO` | Generate | Ruta relativa de la **unidad de crédito** (`credit_path`: `…/Cliente/CREDITO # n` o unidad operativa detectada en Generate). Independiente de si existe subcarpeta **EXTRACTOS** |
| **RutaAsientosContables** (oculta) | `RUTA_ASIENTOS_CONTABLES` | **Solo Finalize** en `cartera_validada_*.xlsx` | Ruta relativa de la carpeta de asientos creada o reutilizada |

**Orden de columnas** (a la derecha de las existentes, tras **Ruta**):

1. En `validacion_pagos_*.xlsx` (Generate): `… | Ruta | RutaUnidadCredito`
2. En `cartera_validada_*.xlsx` (Finalize): `… | Ruta | RutaUnidadCredito | RutaAsientosContables`

**Link carpeta crédito** sigue siendo hipervínculo `webUrl`; no sustituye a **RutaUnidadCredito**.

### Carpeta de asientos (Finalize)

- **Anclaje:** dentro de la ruta de **RutaUnidadCredito** (unidad de crédito), no bajo **EXTRACTOS** ni bajo el padre inmediato del PDF.
- **Nombre:** `ASIENTOS CONTABLES CRED {n}` donde `{n}` es el identificador numérico del crédito de la fila (mismo criterio que Merge usa para `_resolve_credit_digits_for_extract` / columna **Crédito**).
- **Creación:** idempotente vía Graph (si la carpeta ya existe, reutilizar y escribir su ruta en **RutaAsientosContables**).
- **Ámbito:** filas que Finalize incluye en el histórico (típicamente **Validar Pago = SI** y reglas de cierre).
- **Errores (sin warning silencioso):**
  - **RutaUnidadCredito** vacía o ausente en fila que requiere asientos → **error** Finalize (no escribir histórico completo / no marcar job como éxito).
  - Fallo Graph al crear o resolver la carpeta (`403`, `404` en padre, timeout, etc.) → **error** con código dedicado (p. ej. `asientos_folder_create_failed`) y mensaje en `job_status_enrichment` indicando cliente, crédito e **RutaUnidadCredito**.
  - Solo tras crear o reutilizar la carpeta con éxito se rellena **RutaAsientosContables** en `cartera_validada_*.xlsx`.

### Merge

- Leer **únicamente** **RutaAsientosContables** por fila (histórico vía `control_merge_pdfs.xlsx`) y listar PDFs hijos (mismo patrón que **Ruta** para extractos).
- **Sin fallback:** no usar `parent(ruta_extracto)` ni `GRAPH_ASIENTOS_CONTABLES_FOLDER_NAME` para adivinar carpetas. Históricos sin columna o con celda vacía en filas incluidas en el merge → **error** (p. ej. `missing_ruta_asientos_contables`) con mensaje que indique re-ejecutar Finalize tras despliegue.
- Eliminar o dejar de invocar la heurística actual en `merge_composite_validado_pdfs` una vez implementado el contrato de columna.

### Data flow (rutas)

```
Generate
  → Ruta = ruta PDF extracto
  → RutaUnidadCredito = credit_path

Finalize (lee validacion_pagos)
  → copia Ruta + RutaUnidadCredito al histórico
  → por fila elegible: crea …/ASIENTOS CONTABLES CRED {n} bajo RutaUnidadCredito
  → escribe RutaAsientosContables en cartera_validada

Merge (lee cartera_validada)
  → extracto desde Ruta
  → asientos desde RutaAsientosContables
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `review_schema.py` | Modify | `DistribucionCols.ESTADO_LINEA` → nombre visible **Estado Pago**; constante `VALIDAR_PAGO`; `HEADERS` orden; `EstadoPago` (5); eliminar/retirar `EstadoLinea.OPTIONS` legacy de la columna Estado Pago; alias legacy “Estado” → canónico; añadir `RUTA_UNIDAD_CREDITO`, `RUTA_ASIENTOS_CONTABLES` y orden en `HEADERS` |
| `payment_validation_generate.py` | Modify | Renombrar header; DV solo 5 + SI/NO; reemplazar `_build_distribution_rows`; persistir `credit_path` → **RutaUnidadCredito**; ocultar columna técnica (como **Ruta**) |
| `payment_validation_finalize.py` | Modify | Sustituir ramas `EstadoLinea` por combinación Estado Pago + Validar Pago; validación estricta; auditoría; crear carpeta asientos bajo **RutaUnidadCredito**; escribir **RutaAsientosContables** en histórico |
| `merge_composite_validado_pdfs.py` | Modify | Leer **RutaAsientosContables** obligatorio; quitar heurística `parent(extracto)`; errores explícitos si falta columna o ruta |
| `job_status_enrichment.py` | Modify | Mensajes para `asientos_folder_create_failed`, `missing_ruta_asientos_contables`, etc. |
| Notify | Modify | Columna `Validar Pago` + `Estado Pago` en lecturas (sin cambio de ruta asientos salvo spec) |
| `setup_audit_tracking_workbooks.py` | Exists (done) | Crea libros de auditoría en SharePoint; constantes `AUDIT_TRACKING_COLUMNS`, tablas y rutas |
| `payment_validation.py` (router) | Exists (done) | `POST …/setup/audit-tracking-workbooks` |
| Auditoría + tests | Modify | Append en Finalize + tests; reutilizar constantes de columnas del setup o import compartido |

## Interfaces / Contracts

```python
EstadoPago.ALLOWED = frozenset(
    {"ADELANTADO", "ATRASADO", "INCOMPLETO", "NORMAL", "REVISION_MANUAL"}
)

# DistribucionCols (ejemplo)
ESTADO_PAGO = "Estado Pago"     # antes ESTADO_LINEA / "Estado"
VALIDAR_PAGO = "Validar Pago"
RUTA = "Ruta"                   # PDF extracto (Generate)
RUTA_UNIDAD_CREDITO = "RutaUnidadCredito"  # credit_path (Generate)
RUTA_ASIENTOS_CONTABLES = "RutaAsientosContables"  # carpeta creada (solo histórico Finalize)

# HEADERS: … LINK_CARPETA_CREDITO, RUTA, RUTA_UNIDAD_CREDITO
# cartera_validada: añadir RUTA_ASIENTOS_CONTABLES al final
```

**HTTP:** respuestas de `POST …/generate` y Finalize sin campos nuevos; el contrato vive en el Excel subido a SharePoint.

**Migración legacy (ejemplo — cerrar en spec):**

| Antiguo `Estado` | Estado Pago | Validar Pago |
|------------------|-------------|--------------|
| VALIDAR | NORMAL (o INCOMPLETO según saldo) | SI |
| REPROGRAMAR | ADELANTADO | SI/NO según negocio |
| … | … | … |

## Testing Strategy

| Layer | Qué |
|-------|-----|
| Unit | allowlist; `normalize_estado_pago`; mapeo migración |
| Integration | Finalize: fila con `Estado Pago=VALIDAR` (mal migrada) → error |
| Regression | Headers solo `Estado Pago` + `Validar Pago`; sin valores ilegales |
| Unit/Integration | Generate escribe **RutaUnidadCredito** = `credit_path` cuando hay candidato |
| Integration | Finalize crea `ASIENTOS CONTABLES CRED {n}` y rellena **RutaAsientosContables** |
| Integration | Finalize: fallo crear carpeta → job error + mensaje |
| Integration | Merge: sin **RutaAsientosContables** en fila incluida → error; sin llamar heurística legacy |

## Migration / Rollout

1. **(Opcional pero recomendado antes de Finalize con auditoría)** Ejecutar `POST …/setup/audit-tracking-workbooks` en el entorno para asegurar los dos Excels con contrato correcto.  
2. Desplegar schema + Generate (solo 5 + SI/NO).  
3. Finalize: `normalize_distrib_row_keys` acepta alias **Estado** → **Estado Pago** una temporada; ejecutar migración única de libros en carpeta revisión o forzar regeneración.  
4. Actualizar PA / documentación de rutas y variables `GRAPH_AUDIT_*` si aplica.  
5. Desplegar rutas documentales: Generate → Finalize (crear carpetas) → Merge; **obligatorio** regenerar Generate + Finalize antes de Merge (históricos sin **RutaAsientosContables** no son válidos para consolidar).

## Open Questions

- [ ] Matriz completa legacy → (Estado Pago, Validar Pago).  
- [ ] `GRAPH_VALIDAR_EXTRACTO_ESTADO_CONTAINS`: reemplazar por **Validar Pago=SI** + condición sobre Estado Pago o nuevo env.  
- [ ] Posición exacta columna `Validar Pago` vs otras columnas renombradas en UX previo.  
- [x] Rutas Merge/asientos: **RutaUnidadCredito** (Generate) + **RutaAsientosContables** (Finalize) — ver sección anterior.  
- [ ] ¿Incluir **RutaUnidadCredito** / **RutaAsientosContables** en append de auditoría (`pagos_adelantados` / `incompletos`)? Por defecto opcional.  
- [x] Fallo al crear carpeta en Graph: **error** Finalize + mensaje enriquecido (no warning).  
- [x] Merge: **sin fallback**; columna **RutaAsientosContables** obligatoria en filas a consolidar.
