## SDD Archive Report: payment-validation-retry

### 1. Alcance del Cambio
El cambio archivado corresponde única y estrictamente a `payment-validation-retry`. Se solucionó el riesgo detectado en la re-generación temprana omitiendo la barrera de `already_generated = True` exclusivamente en estado `REVISION_REQUIERE_CORRECCION`.

### 2. Power Automate
Se confirma que **no se modificó** ningún flujo de Power Automate.

### 3. Aislamiento de Módulos
Se confirma que **no se tocó** la lógica de:
- Notify
- Merge
- Dry-run
- Apply
- `execution-run-log`

### 4. Pruebas Automatizadas
La suite completa de pruebas quedó en verde, validando un total de **718 passed** de manera exitosa y sin fallos.

### 5. Compilación Estática
El comando `python -m compileall app -q` se ejecutó y **finalizó sin errores**.

### 6. Comportamiento Final
Se verifica el flujo requerido:
- Generate con errores bloqueantes en revisión transiciona el control a `REVISION_REQUIERE_CORRECCION`.
- Finalize bloquea operativamente ese estado.
- Generate puede ser ejecutado nuevamente (reintentado) desde `REVISION_REQUIERE_CORRECCION`.
- En el reintento, Generate **no retorna** de manera prematura con `already_generated`.
- El reintento sobrescribe natural y exactamente la **misma ruta canónica**.
- Si el Excel ya no tiene errores en el reintento, avanza a `REVISION_CREADA`.
- Si el Excel todavía contiene errores, mantiene su estado en `REVISION_REQUIERE_CORRECCION`.
- En cualquier proceso en estado `REVISION_CREADA`, se conserva con rigidez la idempotencia inicial (bloqueo por `already_generated`).

### 7. Integridad Documental
Se confirma que **no se agregaron** atributos auxiliares como eTags, cTags, fingerprints, versionado por intentos ni nuevas columnas al diseño.

### 8. Estructura de Interfaz Excel
Se confirma que **no se ocultaron ni protegieron** hojas adicionales.

### 9. Schemas del Workbook
Se confirma que **no cambiaron** en absoluto los schemas ni contratos JSON internos del workbook de revisión.
