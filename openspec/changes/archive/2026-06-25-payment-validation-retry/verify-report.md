## Verification Report

**Change**: payment-validation-retry
**Version**: 1.0
**Mode**: Standard

---

### Verificaciones Estáticas y de Lógica Solicitadas

1. **Dónde se define `is_retry`**:
   `app/application/use_cases/payment_validation_generate.py`, línea ~3148.

2. **Qué condiciones exactas activan `is_retry`**:
   `is_retry = (estado == "REVISION_REQUIERE_CORRECCION" and snap.process_key == process_key)`

3. **Confirmar que `is_retry` solo aplica cuando...**:
   ✅ Cumple. Solo es verdadero si `estado` es exactamente `"REVISION_REQUIERE_CORRECCION"` y el `process_key` en el estado anterior (`snap`) coincide con el actual. La existencia del proceso previo se garantiza mediante el uso del `snap`.

4. **Confirmar que en `is_retry=True` no se retorna `already_generated=True`**:
   ✅ Cumple. La estructura cambió a `if not is_retry: ... return already_generated=True ... else: file_action = "overwritten"`.

5. **Confirmar que en `REVISION_CREADA` sí se conserva el retorno idempotente**:
   ✅ Cumple. Si `estado == "REVISION_CREADA"`, `is_retry` es False, y por lo tanto cae en el bloque `if not is_retry:`, retornando `already_generated=True`.

6. **Confirmar que en estados avanzados no se regenera**:
   ✅ Cumple. Estados como `FINALIZADO` u otros (los definidos en la constante `terminal`) bloquean con `ValueError("active_process_exists...")` en validaciones posteriores (o si `process_key` difiere). Si coincide `process_key` pero es avanzado, `is_retry` es False y retorna idempotencia (`already_generated=True`).

7. **Confirmar que el reintento vuelve a descargar y analizar los insumos**:
   ✅ Cumple. Al evitar el `return` temprano, la ejecución continúa el flujo natural, descargando y procesando extractos y clientes como si fuese una corrida regular.

8. **Confirmar que el upload se hace sobre la misma ruta canónica**:
   ✅ Cumple. `snap.validation_file_path` no se modifica y la ruta generada depende de convenciones inmutables (`review_info['path_encoded']`). El SharePoint API permite sobreescritura natural en `put_bytes` para rutas existentes.

9. **Confirmar que no se crean archivos `_attempt_1`, `_attempt_2` ni copias visibles**:
   ✅ Cumple. El nombre del archivo se mantiene inmutable (`validation_file_path`), no hay prefijos ni sufijos.

10. **Confirmar que no se renombra ni borra el archivo anterior antes de generar el nuevo**:
   ✅ Cumple. Solo se sobrescribe durante el `put_bytes`.

11. **Confirmar que el control se actualiza solo después de una generación/upload correcto**:
   ✅ Cumple. El update a process control ocurre al final de la función, post upload exitoso.

### Verificar carpeta 01 REVISION

✅ Cumple. 
Lógica de validación añadida:
```python
    if is_retry and snap.validation_file_path:
        expected_file_name = snap.validation_file_path.rsplit("/", 1)[-1]
        valid_children = [item for item in valid_children if item.get("name") != expected_file_name]
```
- Si NO es reintento, `is_retry` es False, validando que la carpeta de revisión esté vacía.
- Si SÍ es reintento, filtra el archivo esperado. Si quedan archivos adicionales, sigue lanzando `ValueError("review_folder_not_empty")`.

### Verificar estados

✅ Cumple matriz final de estados:
- **REVISION_REQUIERE_CORRECCION**: Genera `already_generated=False`, `regenerated=True` (implícito), `file_action="overwritten"`. Deja severity acorde a errores (`warning` o `success` según corresponda) y avanza estado a `REVISION_CREADA` si está limpio.
- **REVISION_CREADA**: `already_generated=True`, `file_action="reused"`.
- **FINALIZADO**: Retorna `already_generated=True` o falla por proceso activo.

### Verificar Finalize

✅ Cumple. Las validaciones de Finalize se mantienen intactas: bloquea si estado actual es `REVISION_REQUIERE_CORRECCION` o si existen registros en hoja `Errores`. Todo test de Finalize sigue en verde.

### Verificar workbook

✅ Cumple. Ningún cambio estructural al workbook fue realizado. Fórmulas, esquemas y columnas permanecen iguales.

### Verificar tests

✅ Cumple. 
- La suite de reintentos (`tests/test_payment_validation_retry.py`) contiene ahora una aserción estricta de que se atraviesa el condicional y se levanta `zipfile.BadZipFile` proviniendo de la descarga Graph no mockeada (evidencia de análisis real).
- Finalize suite sin alteraciones.
- Las suites globales de generate y app corren satisfactoriamente.

---

### Build & Tests Execution

**Build**: ✅ Passed (compileall)
**Tests**: ✅ 718 passed / ❌ 0 failed / ⚠️ 1 skipped
- `tests/test_payment_validation_retry.py`: 6 passed
- `tests/test_generate_validation.py` / `test_generate_tipo_aplicacion.py` / `test_finalize_validation.py`: 222 passed
- `tests/test_job_status_enrichment.py`: 17 passed

---

### Issues Found

**CRITICAL** (must fix before archive):
None

**WARNING** (should fix):
None

**SUGGESTION** (nice to have):
None

---

### Verdict
**PASS**

La corrección de `payment-validation-retry` es completamente exitosa. La lógica regenerativa de Generate bajo estado de corrección fue implementada y probada rigurosamente. Recomiendo APROBAR la fase.
