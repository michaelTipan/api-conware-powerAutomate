# Variables de entorno — Render / producción

Documentación oficial del proyecto (no usar `.env` ni `.env.local` como referencia en git; esos archivos son locales e ignorados).

**No incluye valores reales.** Use placeholders al configurar Render.

## A. Microsoft Graph (obligatorias)

| Variable | Tipo | Notas |
|----------|------|-------|
| `GRAPH_TENANT_ID` | secret/config | `<tenant-id>` |
| `GRAPH_CLIENT_ID` | config | `<client-id>` |
| `GRAPH_CLIENT_SECRET` | secret | `<client-secret>` |
| `GRAPH_SCOPE` | config | Opcional; default `https://graph.microsoft.com/.default` |
| `GRAPH_BASE_URL` | config | Opcional |
| `GRAPH_HTTP_TIMEOUT_SECONDS` | config | Opcional |
| `GRAPH_PUT_MAX_RETRIES` | config | Opcional |
| `GRAPH_PUT_RETRY_BASE_SECONDS` | config | Opcional |

## B. SharePoint (obligatorias para flujo de pagos)

| Variable | Tipo | Notas |
|----------|------|-------|
| `GRAPH_SHAREPOINT_SITE_SEARCH` | config | `<sharepoint-site-search>` |
| `GRAPH_SHAREPOINT_DRIVE_NAME` | config | Ej. `Documentos` |
| `GRAPH_CLIENTS_BASE_PATH` | path | Base carpetas clientes/créditos |
| `GRAPH_BANK_PAYMENTS_FILE_PATH` | path | Excel reporte Banco de Bogotá |
| `GRAPH_BANK_PAYMENTS_FILE_PATH_BANCOLOMBIA` | path | **Recomendada** — reporte Bancolombia |
| `GRAPH_SHAREPOINT_FILE_PATH` | path | Alias legacy (Bogotá) para notify/merge |
| `GRAPH_SHAREPOINT_FILE_PATH_BANCOLOMBIA` | path | Alias legacy Bancolombia |

## C. Carpetas del flujo de pagos

**Compatibilidad (actuales en Render):** rutas completas por carpeta.

| Variable | Obligatoria | Notas |
|----------|-------------|-------|
| `GRAPH_PAYMENT_VALIDATION_CONTROL_PATH` | Recomendada | `…/00 CONTROL` |
| `GRAPH_PAYMENT_VALIDATION_REVIEW_PATH` | Sí (generate/finalize) | `…/01 REVISION` |
| `GRAPH_PAYMENT_VALIDATION_HISTORY_PATH` | Sí (finalize/notify) | `…/02 HISTORICO` |
| `GRAPH_PAYMENT_VALIDATION_LOGS_PATH` | Opcional | `…/04 LOGS` |
| `GRAPH_FOLLOWUP_PAGOS_ADELANTADOS_PATH` | Opcional | Setup followup |
| `GRAPH_FOLLOWUP_PAGOS_INCOMPLETOS_PATH` | Opcional | Setup followup |
| `GRAPH_IBR_DIARIO_PATH` | Opcional | Setup IBR |

**Canónicas (nuevas):** componen rutas desde base + subcarpeta.

| Variable | Notas |
|----------|-------|
| `PAYMENT_VALIDATION_BASE_FOLDER` | Raíz del flujo |
| `PAYMENT_VALIDATION_CONTROL_FOLDER` | Ej. `00 CONTROL` |
| `PAYMENT_VALIDATION_REVIEW_FOLDER` | Ej. `01 REVISION` |
| `PAYMENT_VALIDATION_HISTORICAL_FOLDER` | Ej. `02 HISTORICO` |
| `PAYMENT_VALIDATION_ERRORS_FOLDER` | Ej. `03 ERRORES` |
| `PAYMENT_VALIDATION_LOGS_FOLDER` | Ej. `04 LOGS` |
| `PAYMENT_VALIDATION_EMAIL_FOLDER` | Ej. `05 EMAIL` |
| `PAYMENT_VALIDATION_ASIENTOS_FOLDER` | Ej. `06 ASIENTO CONTABLES GENERADOS` |

Prioridad en código: `GRAPH_PAYMENT_VALIDATION_*_PATH` → `PAYMENT_VALIDATION_BASE_FOLDER` + subcarpeta → defaults HBI en `payment_validation_settings.py`.

## D. Bancos (Opción A — variables por banco)

Por banco (`BOGOTA` / `BANCOLOMBIA`):

| Variable | Ejemplo placeholder |
|----------|---------------------|
| `PAYMENT_BANK_*_CODE` | `banco_bogota` |
| `PAYMENT_BANK_*_NAME` | Nombre display |
| `PAYMENT_BANK_*_INPUT_FILE_PATH` | `<relative-path>/BANCO_….xlsx` |
| `PAYMENT_BANK_*_CONTROL_FILE` | `control_proceso_validacion_pagos_banco_….xlsx` |
| `PAYMENT_BANK_*_EMAIL_LABEL` | Etiqueta visible en cuerpo `{banco}` (ej. `BANCO BOGOTA`) |
| `PAYMENT_BANK_*_EMAIL_SUBJECT` | `ABONOS BANCO …` |
| `PAYMENT_BANK_*_EMAIL_PDF_NAME_TEMPLATE` | `ABONOS … {fecha}.pdf` |

Alias legacy: `GRAPH_BANK_PAYMENTS_FILE_PATH`, `GRAPH_BANK_PAYMENTS_FILE_PATH_BANCOLOMBIA`.

## E. Notify / correo

| Variable | Obligatoria | Notas |
|----------|-------------|-------|
| `GRAPH_VALIDAR_NOTIFY_CORREOS_XLSX_PATH` | Recomendada | `…/CORREOS.xlsx` |
| `GRAPH_VALIDAR_NOTIFY_EMAIL_SUBJECT` | **Obsoleta multi-banco** | Ignorada; use `PAYMENT_BANK_*_EMAIL_SUBJECT` |
| `GRAPH_VALIDAR_NOTIFY_BODY_INTRO_TEMPLATE` | Opcional | HTML intro |
| `GRAPH_VALIDAR_EXTRACTO_ESTADO_CONTAINS` | Opcional | Default `VALIDAR` |
| `GRAPH_VALIDAR_ESTADO_PAGO_CONTAINS` | Opcional | Filtro extra |
| `GRAPH_VALIDAR_NOTIFY_ATTACH_PDFS` | Opcional | Default `true` |
| `GRAPH_VALIDAR_NOTIFY_EXPORT_EMAIL_PDF` | Opcional | Default `true` |
| `GRAPH_VALIDAR_NOTIFY_EXPORT_EMAIL_PDF_FOLDER_PATH` | Opcional | Carpeta `05 EMAIL` |
| `GRAPH_VALIDAR_NOTIFY_EXPORT_EMAIL_PDF_NAME_TEMPLATE` | **Obsoleta multi-banco** | Ignorada; use `PAYMENT_BANK_*_EMAIL_PDF_NAME_TEMPLATE` |
| `GRAPH_MAIL_ATTACH_MAX_BYTES` | Opcional | |

Remitente/destinatarios: **CORREOS.xlsx**, no variables de mail sueltas.

## F. Merge

| Variable | Notas |
|----------|-------|
| `GRAPH_MERGE_COMPOSITE_OUTPUT_FOLDER_PATH` | Salida PDFs consolidados |
| `GRAPH_MERGE_COMPOSITE_CLIENTE_COLUMN` | Opcional |
| `GRAPH_MERGE_COMPOSITE_CREDITO_COLUMN` | Opcional |

## G. Amortización / links

| Variable | Notas |
|----------|-------|
| `GRAPH_LINK_EXTRACTO_PATH_ANCHOR` | Opcional |
| `GENERATE_EXTRACT_SELECTION_V2` | Opcional; default `true` |

## H. Operación / logs

| Variable | Notas |
|----------|-------|
| `LOG_LEVEL` | `INFO` en producción |
| `GRAPH_VALIDATION_FILE_PREFIX` | Prefijo archivos validación |
| `GRAPH_EXTRACT_KEYWORD` | Default `Extracto` |
| `GRAPH_CREDIT_MATCH_TOLERANCE` | Opcional |
| `GRAPH_ASIENTOS_CONTABLES_FOLDER_NAME` | Ensure-asientos |
| `GRAPH_EXTRACTOS_FOLDER_NAME` | Ensure-asientos |

## I. Variables obsoletas — eliminar de Render tras pruebas

| Variable | Motivo |
|----------|--------|
| `GRAPH_MERGE_CONTROL_WORKBOOK_PATH` | `control_merge_pdfs.xlsx` retirado |
| `GRAPH_NOTIFY_BANK_NAME` | No la lee la app |
| `GRAPH_VALIDAR_NOTIFY_EMAIL_SUBJECT` | Ignorada; fuerza asunto incorrecto en multi-banco |
| `GRAPH_VALIDAR_NOTIFY_EXPORT_EMAIL_PDF_NAME_TEMPLATE` | Ignorada; usar `PAYMENT_BANK_*_EMAIL_PDF_NAME_TEMPLATE` |
| `GRAPH_MAIL_SENDER_EMAIL` | CORREOS.xlsx |
| `GRAPH_VALIDAR_NOTIFY_TO` / `GRAPH_VALIDAR_NOTIFY_CC` | CORREOS.xlsx |
| `GRAPH_PAYMENT_VALIDATION_BASE_PATH` | No implementada |
| `GRAPH_PAYMENT_VALIDATION_ERRORS_PATH` | No implementada |
| `GRAPH_VALIDATED_HISTORY_FILE_PREFIX` | No implementada |
| `GRAPH_LINK_EXTRACTO_URL_PREFIX_STRIP` | No implementada |
| `GRAPH_LINK_EXTRACTO_AFTER_SITES_STRIP_LIBRARIES` | No implementada |
| `VERCEL_*`, `JOB_STORE_BACKEND`, `BLOB_*` | Runtime anterior |

## Migrar a otro SharePoint

1. Actualizar `GRAPH_SHAREPOINT_SITE_SEARCH` y rutas `GRAPH_*` o `PAYMENT_VALIDATION_*`.
2. Configurar `PAYMENT_BANK_*_INPUT_FILE_PATH` y controles por banco.
3. Ejecutar setup de controles y smoke del flujo (ver `docs/payment-validation-production-flow.md`).

## Agregar un tercer banco (futuro)

Hoy el código admite dos bancos vía `list_payment_banks()`. Un tercer banco requiere extender el módulo `payment_validation_settings` y los tests — no solo variables.
