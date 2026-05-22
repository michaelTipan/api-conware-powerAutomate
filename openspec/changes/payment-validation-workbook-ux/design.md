# Design: UX del workbook de validación de pagos

## Technical Approach

Centralizar layout y estilos en `review_schema.py` + `payment_validation_generate.py`, sin tocar lógica de extractos v2, unidades de crédito, links ni `Ruta`. Finalize sigue leyendo por **nombre de columna** vía `dict(zip(headers, row))`; se añade **normalización de alias legacy** al cargar filas. Fórmulas y CF usan `HEADERS.index`, no letras fijas (hoy `INDIRECT("N"&ROW())` debe corregirse).

## Architecture Decisions

| Decision | Options | Choice | Rationale |
|----------|---------|--------|-----------|
| Orden columnas | Reordenar `HEADERS` vs copiar hoja | Reordenar `DistribucionCols.HEADERS` | Una fuente de verdad; fórmulas ya indexan por nombre |
| Renombre editable | Solo etiqueta vs renombrar constantes | Constantes = texto visible nuevo + mapa legacy | Finalize y libros viejos siguen funcionando |
| Freeze | Fijo `D4` vs calcular tras Crédito | `get_column_letter(index(CREDITO)+1)` + fila datos | Sobrevive futuros cambios de orden |
| Filtros encabezado | Quitar `auto_filter` vs ocultar | No asignar `auto_filter.ref` (Casos/Distrib/Errores) | Quita iconos; mantiene DV en celdas |
| Banda cliente | Por `ID Pago` (actual) vs `Cliente` | Por **Cliente** | Requisito UX; ordenar filas por cliente ayuda lectura |
| Estado línea color | Solo relleno celda vs CF | **CF dinámico** + relleno directo en `estado_c` | CF ya existe; corregir columna; celda Estado prevalece sobre banda |
| Anchos money | Autofit vs constante | `DIST_MONEY_COL_WIDTH = 14` aplicada a todas las cols money | Uniformidad visual |

## Data Flow

```
generate_payment_validation
  → HEADERS (schema) → _apply_table_header_row / append filas
  → _apply_distribution_formulas (refs por índice)
  → _style_distrib_sheet (bandas cliente, money width, freeze D4)
  → _apply_distribution_conditional_formatting (Estado, col dinámica)
  → _apply_distrib_hyperlinks / hide Ruta
  → put_bytes → SharePoint

finalize_payment_validation
  → lee headers fila 3
  → row_dict = zip(headers, values)
  → _normalize_distrib_row_keys(row_dict)  [NUEVO]
  → dist.get(DistribucionCols.APLICAR_A_EXTRACTO) …
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `app/application/services/review_schema.py` | Modify | Nuevo orden `HEADERS`; renombrar 3 columnas; alias legacy dict |
| `app/application/use_cases/payment_validation_generate.py` | Modify | Freeze 3 cols; sin auto_filter; money width const; banda por cliente; editable tint; Control Procesar SI/NO; CF col dinámica; textos ayuda |
| `app/application/use_cases/payment_validation_finalize.py` | Modify | `_normalize_distrib_row_keys`; usar constantes nuevas; ampliar fallbacks existentes |
| `app/application/job_status_enrichment.py` | Modify | Mensajes usuario con nombres nuevos columnas |
| `tests/test_generate_validation.py` | Modify | Orden, freeze `D4`, sin filter, anchos, bandas, Control fila 7 |
| `tests/test_finalize_validation.py` | Modify | Fixtures legacy + nuevos headers; lectura compatible |
| `tests/test_finalize_validation.py` | Modify | `make_distrib_row` orden/nombres alineados |

**Sin cambios:** Notify, Merge, routers HTTP, selección extracto v2, detección crédito, targets hyperlinks.

## Interfaces / Contracts

```python
# review_schema.py (ejemplo)
class DistribucionCols:
    APLICAR_A_EXTRACTO = "Aplicar a extracto"      # era Valor intereses
    MORA_A_APLICAR = "Mora a aplicar"              # era Abono a K
    OTROS_VALORES = "Otros valores"                # era Intereses de mora
    HEADERS = ["ID Pago", "Cliente", "Crédito", "Monto banco", ...]

DISTRIB_LEGACY_HEADER_ALIASES: dict[str, str] = {
    "Valor intereses": DistribucionCols.APLICAR_A_EXTRACTO,
    "Abono a K": DistribucionCols.MORA_A_APLICAR,
    "Intereses de mora": DistribucionCols.OTROS_VALORES,
}

# generate
DIST_MONEY_COL_WIDTH = 14
def _distrib_freeze_panes_cell(dr: int) -> str:
    col = DistribucionCols.HEADERS.index(DistribucionCols.CREDITO) + 1
    return f"{get_column_letter(col + 1)}{dr}"  # → D4 con header fila 3
```

Finalize: tras `row_dict = dict(zip(headers, row))`, aplicar alias para que claves canónicas existan aunque el Excel tenga headers viejos.

## Compatibilidad

| Escenario | Comportamiento |
|-----------|----------------|
| Excel nuevo (post-deploy) | 19 cols, orden nuevo, headers nuevos |
| Excel en revisión (headers viejos) | Finalize normaliza alias → mismas validaciones |
| `Ruta` | Mismo nombre, misma columna final, oculta |
| Links | Sin cambio de URL ni columnas link |
| Secretary/histórico | Finalize escribe hojas propias; no dependen del orden Distribución salvo lectura vía colmap |

**Riesgo alto:** CF con `INDIRECT("N"&ROW())` deja de coincidir si no se parametriza — **bloqueante en implementación**.

**Riesgo medio:** Tests/fixtures con listas posicionales de 18/19 columnas — actualizar `make_distrib_row` y legacy padding.

**Riesgo bajo:** Mensajes Generate (`DISTRIB_HELP`) citan nombres viejos.

## Estilo (precedencia por celda Distribución)

1. **Estado línea** — color estado (`_ESTADO_LINEA_FILLS` / CF)
2. **Observación advertencia** — `_FILL_OBS_WARNING` solo col Observación
3. **Editable** — tinte sobre color banda cliente (`_editable_fill_on_base(group_fill)`)
4. **Saldo por asignar** — `_FILL_SALDO_COL` si no editable/estado
5. **Banda cliente** — pastel A/B alternando por cambio de `Cliente`

Control fila `procesar_row` (hoy fila 7): altura ~28–32, fill fila suave, celda Valor SI=verde (`D4EDDA`), NO=ámbar/rojo suave; DV Procesar intacto.

## Testing Strategy

| Layer | What | Approach |
|-------|------|----------|
| Unit | `_normalize_distrib_row_keys`, freeze cell, money cols set | Tests puros en schema/helpers |
| Integration | `run_generate` workbook | Headers orden; `max_column` 19; col 18 = Link carpeta; Ruta hidden; no `auto_filter.ref`; freeze `D4`; money widths iguales |
| Integration | Finalize fixtures | Libro legacy headers viejos + nuevo → `missing_mora` / VALIDAR igual |
| Visual smoke | Control Procesar | Fila 7 height; fill según SI/NO |

Casos explícitos: Link carpeta presente; Estado línea colores por valor; agrupación visual dos clientes alternan fill.

## Migration / Rollout

- Libros **ya generados** en SharePoint: Finalize compatible vía alias; secretaria puede seguir editando con headers viejos hasta regenerar Generate.
- Regenerar Generate tras deploy para UX completa (orden, freeze, sin filtros, renombres).
- Sin feature flag: cambio solo presentación.
- Verificar deploy ejecuta commit con schema nuevo (evitar desalineación runtime ya diagnosticada).

## Open Questions

- [ ] ¿Ordenar filas de datos por `(Cliente, Crédito)` en Generate además del color, o solo color sin sort?
- [ ] Ancho money `14` vs `16` — validar con secretaria un cliente real.
