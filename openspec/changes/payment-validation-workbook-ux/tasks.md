# Tasks: UX workbook validación de pagos

## Phase 1: Schema y contratos

- [ ] 1.1 En `review_schema.py`: reordenar `DistribucionCols.HEADERS` → ID Pago, Cliente, Crédito, Monto banco, Fecha banco, …, Link carpeta crédito, Ruta (sin mover Ruta del final).
- [ ] 1.2 Renombrar constantes/labels: `APLICAR_A_EXTRACTO`, `MORA_A_APLICAR`, `OTROS_VALORES` con textos visibles nuevos; conservar referencias internas coherentes en todo el repo.
- [ ] 1.3 Añadir `DISTRIB_LEGACY_HEADER_ALIASES` (Valor intereses / Abono a K / Intereses de mora → nombres nuevos) exportable para Finalize.
- [ ] 1.4 Definir `DIST_MONEY_COL_WIDTH = 16` en `payment_validation_generate.py` (o schema si conviene import único).

## Phase 2: Generate — layout y estilos

- [ ] 2.1 Actualizar `_build_distribution_rows` y cualquier dict por columna para usar constantes nuevas; **no** reordenar filas generadas (orden actual intacto).
- [ ] 2.2 `_distrib_freeze_panes_cell`: congelar tras Crédito (3ª columna) → equivalente `D4` con header fila 3; mantener freeze previo de 2 cols sustituido por 3.
- [ ] 2.3 Eliminar asignación `auto_filter.ref` en `_style_distrib_sheet`, `_style_casos_sheet`, `_style_errores_sheet`.
- [ ] 2.4 `_style_distrib_sheet`: ancho 16 en Monto banco, Valor extracto, Aplicar a extracto, Mora a aplicar, Otros valores, Total aplicado, Saldo por asignar; quitar floors distintos por money col.
- [ ] 2.5 Bandas pastel alternas por **cambio de Cliente** (no por ID Pago); sin `sort` de filas.
- [ ] 2.6 Helper `_editable_fill_on_base(group_fill)`; aplicar en cols editables con precedencia: Estado > Obs warning > editable > saldo > banda.
- [ ] 2.7 `_apply_distribution_conditional_formatting`: fórmula Estado con `get_column_letter(HEADERS.index(ESTADO_LINEA))`, no `"N"` fijo.
- [ ] 2.8 Mantener `_configure_distrib_technical_ruta_column`, links y textos hyperlink sin cambiar targets.
- [ ] 2.9 `_write_control_sheet` / post-estilo: fila Procesar (fila 7) altura mayor, banda fila suave; fill Valor SI verde / NO ámbar suave; DV intacto.
- [ ] 2.10 Actualizar `DISTRIB_HELP` y strings que citen nombres viejos de columnas editables.

## Phase 3: Finalize compatibilidad

- [ ] 3.1 Implementar `_normalize_distrib_row_keys(row_dict)` en `payment_validation_finalize.py` usando `DISTRIB_LEGACY_HEADER_ALIASES`.
- [ ] 3.2 Tras `dict(zip(headers, row))` en lectura Distribución, normalizar keys antes de validaciones.
- [ ] 3.3 Reemplazar usos de `VALOR_INTERESES`/`ABONO_K`/`INTERESES_MORA` por constantes nuevas; mantener fallbacks legacy redundantes si útil.
- [ ] 3.4 Verificar `colmap.get(DistribucionCols.RUTA)` y secretary/histórico sin cambios de contrato.

## Phase 4: Mensajes operativos

- [ ] 4.1 `job_status_enrichment.py`: mensajes de error/usuario con Aplicar a extracto, Mora a aplicar, Otros valores.

## Phase 5: Tests

- [ ] 5.1 `test_generate_validation.py`: assert orden headers (Crédito pos 3); freeze `D4`; sin `auto_filter`; money cols width == 16; Link carpeta antes de Ruta oculta.
- [ ] 5.2 Test bandas: dos clientes en orden generado alternan fill sin reordenar filas.
- [ ] 5.3 Test Control: fila 7 altura destacada; fill Procesar según SI/NO.
- [ ] 5.4 `test_finalize_validation.py`: `make_distrib_row` orden 19 cols + nombres nuevos; padding legacy 18→19.
- [ ] 5.5 Test Finalize: libro fixture headers **viejos** pasa validación mora/VALIDAR igual que headers nuevos.
- [ ] 5.6 Ejecutar `pytest tests/` completo; corregir referencias rotas a nombres de columna en tests/helpers.

## Phase 6: Verificación manual (post-implementación)

- [ ] 6.1 Generate local/Render: abrir Distribución fila 3 — orden y 19 columnas; scroll horizontal con ID/Cliente/Crédito fijos.
- [ ] 6.2 Confirmar Notify/Merge sin diff y contratos HTTP sin cambios.
