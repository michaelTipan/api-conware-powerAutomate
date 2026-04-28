# Power Automate → API Render

**URL de ejemplo:** `https://api-render-conware-powerautomate.onrender.com`

## HTTP estándar de Power Automate (solo caja “Body” de texto)

El conector **HTTP** que solo muestra **URI, Method, Headers, Queries, Body** (texto) **no arma multipart**. Usa este endpoint:

| Campo | Valor |
|--------|--------|
| **Method** | `POST` |
| **URI** | `https://api-render-conware-powerautomate.onrender.com/parse-excel-json` |
| **Headers** | Clave `Content-Type`, valor `application/json` |
| **Body** | JSON (ver abajo) |

Antes del HTTP: acción **Obtener contenido del archivo** (mismo archivo que disparó el flujo).

**Body** (modo código `</>` si lo tienes, o escribe el JSON y sustituye con contenido dinámico):

```json
{
  "file_name": "<nombre del archivo .xlsx del disparador o dinámico>",
  "excel_base64": "<expresión>",
  "run_id": ""
}
```

- **`excel_base64`:** en el editor de expresiones suele ser  
  `base64(body('NOMBRE_INTERNO_DEL_PASO'))`  
  donde `NOMBRE_INTERNO_DEL_PASO` es el identificador del paso **Obtener contenido del archivo** (en *Ver código* / *Peek code* del paso verás el nombre exacto, p. ej. `Obtener_contenido_del_archivo`).
- **`file_name`:** puedes enlazar el **Nombre de archivo** que da el disparador de SharePoint.
- **`run_id`:** opcional; puedes dejar `""` o usar `workflow()['run']['name']`.

Si la API respondiera error de Base64, algunos conectores ya devuelven el archivo en Base64: prueba **sin** envolver en `base64(...)` y usa solo el campo dinámico que corresponda al contenido.

**Parámetros avanzados:** si solo ves “Autenticación”, en la pestaña **Parámetros** siguen apareciendo **Method**, **URI**, **Body**, etc.; “avanzados” solo añade opciones extra, no sustituye el Body.

---

## ¿Está bien el disparador?

- **Cuando se crea o modifica un archivo (solo propiedades)** en la carpeta/biblioteca: **sí**, es un disparador habitual.
- Ese disparador ya trae **metadatos** del archivo (nombre, identificador, ruta, etc.). **No trae el contenido binario del Excel.**

## ¿Hace falta “Obtener archivos (solo propiedades)”?

- **Casi nunca**, si el disparador ya es “un archivo en esta carpeta”: ya sabes **qué** archivo se subió.
- **“Obtener archivos (solo propiedades)”** solo lista archivos y sus propiedades; **sigue sin darte el contenido** del `.xlsx`.
- Úsalo solo si tu lógica necesita **recorrer varios archivos** de una carpeta (por ejemplo, no tienes disparador por archivo y quieres procesar una lista).

## Paso imprescindible antes del HTTP

- Acción: **Obtener contenido del archivo** (`Get file content`).
  - **SharePoint:** sitio + **Identificador de archivo** del disparador (o el `Identifier` / `ID` que muestre tu conector).
  - **OneDrive for Business:** análogo, “Obtener contenido de archivo” con el id/ruta del disparador.

Sin este paso, **no puedes** enviar el Excel (ni a `/parse-excel` multipart ni a `/parse-excel-json`).

## Condición (recomendado)

- Filtrar por extensión: nombre termina en `.xlsx` (y si quieres, excluir temporales `~$`).
- Así solo llamas a la API cuando el archivo es Excel real.

## Llamada HTTP alternativa: multipart (`/parse-excel`)

Si tu acción HTTP permite **Form-data** / **Multipart**:

- **Método:** `POST`
- **URI:** `.../parse-excel`
- **Autenticación:** `None` (si la API es pública).
- **Cuerpo:** campo **`file`** = contenido binario de **Obtener contenido del archivo**; opcional **`run_id`** en texto.

No pongas `Content-Type: multipart/...` a mano si el conector genera el boundary.

## Respuesta de la API (Parse JSON)

Ejemplo de campos útiles en el cuerpo de la respuesta:

- `status`, `run_id`, `file_name`, `row_count`, `elapsed_ms`
- `columns` — lista de nombres de columna
- `rows` — lista de objetos, una fila por elemento (todas las columnas de la primera hoja)

Usa **Analizar JSON** (`Parse JSON`) con un esquema generado a partir de un ejemplo de respuesta de `/parse-excel` o `/parse-excel-json` (la forma del JSON es la misma).

## Plan gratuito de Render (cold start)

La primera petición tras un rato sin uso puede tardar **varios segundos** mientras el servicio “despierta”. Las siguientes suelen ir más rápido.

## Resumen del flujo recomendado

1. Disparador: archivo creado/modificado (propiedades) en carpeta.  
2. Condición: es `.xlsx`.  
3. **Obtener contenido del archivo** (del mismo archivo del disparador).  
4. **HTTP POST** a `/parse-excel-json` con JSON + Base64 (HTTP estándar), **o** multipart a `/parse-excel` si tu conector lo permite.  
5. **Analizar JSON** y usar `rows` / `columns` en pasos siguientes.
