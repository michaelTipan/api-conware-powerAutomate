# Proposal: Estado Pago + Validar Pago + Excels de auditoría

## Intent

- **Renombrar** la columna actual `Estado` → **`Estado Pago`**, con solo: ADELANTADO, ATRASADO, INCOMPLETO, NORMAL, REVISION_MANUAL (sin valores legacy).
- **Una** columna nueva: **`Validar Pago`** (SI/NO) a la derecha de Estado Pago.
- Excels de auditoría (`pagos_adelantados.xlsx`, `pagos_incompletos.xlsx`) en **00 CONTROL** (rutas relativas al drive configurables por env; por defecto junto al resto de controles).

## Approach

Sustituir semántica de `EstadoLinea`/`VALIDAR` por par **Estado Pago** + **Validar Pago**; migración determinista de libros viejos; validación estricta en APIs.

**Bootstrap de auditoría (ya implementado en API):** una sola ejecución de `POST /graph/sharepoint/payment-validation/setup/audit-tracking-workbooks` crea ambos libros de forma idempotente (no sobrescribe si ya existen). El append que hará **Finalize** debe respetar el mismo contrato de hoja/columnas que ese setup.

## Specification

Por escribir en `specs/` (matriz migración, reglas INCOMPLETO, Notify/Merge).

## Risks

Regresión Power Automate si depende del texto `VALIDAR` en columna Estado; coordinar env y documentación.
