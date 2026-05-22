# Flujo de Power Automate (Excel -> Python -> datos extraidos)

## 1) Trigger en SharePoint
- Accion: `When a file is created (properties only)` o `When a file is created or modified (properties only)`.
- Sitio: tu sitio de SharePoint.
- Biblioteca: donde suben el Excel.

## 2) Obtener contenido del archivo
- Accion: `Get file content`.
- File Identifier: el del trigger.

## 3) Construir payload HTTP
- Accion: `Compose` (opcional) para formar JSON:

```json
{
  "file_name": "@{triggerOutputs()?['body/{FilenameWithExtension}']}",
  "excel_base64": "@{base64(body('Get_file_content'))}",
  "run_id": "@{workflow()?['run']?['name']}",
  "column_letter": "C"
}
```

## 4) Llamar la API
- Accion: `HTTP`.
- Method: `POST`.
- URI: `https://<tu-servicio-en-render>/parse-excel`
- Headers:
  - `Content-Type: application/json`
- Body: salida del `Compose` (o JSON directo).

Nota: `POST /validate-excel` sigue existiendo como alias del mismo comportamiento (obsoleto).

## 5) Procesar respuesta
- Accion: `Parse JSON`.
- Campos utiles:
  - `status`, `run_id`, `file_name`, `column_letter`
  - `total_rows`, `non_empty_count`, `unique_count`, `duplicate_count`
  - `unique_values`, `values_in_row_order`, `duplicates`
  - `elapsed_ms`

## 6) Acciones finales
- Usar `values_in_row_order` o `unique_values` en condiciones, bucles, tablas, notificaciones, etc.
- La comparacion con carpetas de SharePoint, si la necesitas, hazla en Power Automate con los datos devueltos.
