# Excel Parser API

Servicio Python que recibe un Excel en base64 (desde Power Automate u otro cliente), lee una columna configurable y devuelve los valores extraidos y metricas. Disenado para invocarse por HTTP.

## Stack
- FastAPI
- Pandas + OpenPyXL
- Render (deploy)

## Arquitectura hexagonal
- **Dominio** (`app/domain/`): excepciones, puertos (`GraphApiPort`), servicios puros (p. ej. parseo de Excel).
- **Aplicacion** (`app/application/`): casos de uso y reglas de orquestacion (SharePoint desde `.env`, parseo).
- **Adaptadores secundarios** (`app/adapters/secondary/`): cliente HTTP Microsoft Graph (`MsGraphClient`).
- **Adaptador primario** (`app/adapters/primary/http/`): FastAPI (routers, dependencias, `create_app`).
- **Entrada** (`app/main.py`): carga `.env` y expone `app` para `uvicorn app.main:app`.
- Los modulos `app/excel_parser.py`, `app/graph_client.py` y `app/sharepoint_resolve.py` son **re-exportes** por compatibilidad.

## Ejecucion local
1. Crear entorno virtual e instalar: `pip install -r requirements.txt`
2. Iniciar API: `uvicorn app.main:app --reload`

## Endpoint principal
- `POST /parse-excel`
- Entrada JSON:
  - `file_name`
  - `excel_base64`
  - `run_id` (opcional)
  - `column_letter` (opcional, default `C`)

## Respuesta
- `unique_values`: valores unicos (orden de primera aparicion)
- `values_in_row_order`: todas las celdas no vacias de la columna, en orden de fila
- `duplicates`: valores repetidos (tras normalizar)
- Conteos: `total_rows`, `non_empty_count`, `unique_count`, `duplicate_count`
- `elapsed_ms`

## Alias
- `POST /validate-excel` — mismo comportamiento que `/parse-excel` (obsoleto).

## Integracion Microsoft Graph (base)
- Variables de entorno requeridas:
  - `GRAPH_TENANT_ID`
  - `GRAPH_CLIENT_ID`
  - `GRAPH_CLIENT_SECRET`
  - `GRAPH_SCOPE` (default `https://graph.microsoft.com/.default`)
  - `GRAPH_BASE_URL` (default `https://graph.microsoft.com/v1.0`)
- Endpoints de prueba:
  - `GET /graph/users?top=10` (escenario recomendado con `client_credentials`)
  - `GET /graph/me` (puede fallar en app-only segun permisos/flujo)

## SharePoint con Graph (carpetas/archivos)
- Obtener sitio por path:
  - `GET /graph/sharepoint/site?hostname=contoso.sharepoint.com&site_path=/sites/MiSitio`
- Listar document libraries (drives) del sitio:
  - `GET /graph/sharepoint/sites/{site_id}/drives`
- Listar carpetas/archivos en un drive:
  - `GET /graph/sharepoint/drives/{drive_id}/children?folder_item_id=root`
- Descargar contenido de archivo (base64):
  - `GET /graph/sharepoint/drives/{drive_id}/item-content?item_id={item_id}`
- Actualizar archivo existente por `item_id` (base64):
  - `PUT /graph/sharepoint/drives/{drive_id}/item-content?item_id={item_id}`
- Crear/reemplazar archivo por ruta (base64):
  - `PUT /graph/sharepoint/drives/{drive_id}/path-content?item_path=Carpeta/reporte.xlsx`

Permisos de aplicacion comunes en Azure AD (segun caso): `Sites.Read.All`, `Sites.ReadWrite.All`, `Files.Read.All`, `Files.ReadWrite.All`.

## SharePoint desde variables de entorno (tu flujo)
- `GRAPH_TENANT_ID`: GUID de `https://login.microsoftonline.com/{GUID}/oauth2/v2.0/token`.
- `GRAPH_SHAREPOINT_SITE_SEARCH`: texto de `GET .../sites?search=` (ej. nombre del sitio).
- `GRAPH_SHAREPOINT_FILE_PATH`: ruta bajo el root del drive (ej. `Carpeta/subcarpeta/archivo.xlsx`); se codifica por segmentos para Graph.
- Opcional: `GRAPH_SHAREPOINT_DRIVE_NAME` (ej. `Documentos`); si no se define, se usa el primer drive devuelto.
- Los ids de sitio y biblioteca (`site id`, `drive id`) **no** se configuran: salen de las respuestas de Graph al resolver el flujo.

Endpoints:
- `GET /graph/sharepoint/sites-search?search=...` (o sin query si ya esta `GRAPH_SHAREPOINT_SITE_SEARCH`).
- `GET /graph/sharepoint/resolve-env` — devuelve `site_id`, `drive_id`, ruta codificada y metadatos del archivo.
- `GET /graph/sharepoint/excel-from-env` — mismo flujo y archivo en `content_base64`.
- `PUT /graph/sharepoint/file-from-env` — sube/reemplaza el archivo en `GRAPH_SHAREPOINT_FILE_PATH` (body JSON `content_base64`).
- `POST /graph/sharepoint/parse-excel-from-env?column_letter=C&run_id=` — descarga el Excel configurado y aplica el parser como `POST /parse-excel`.

Al arrancar en local, `python-dotenv` carga `.env` automaticamente.

## Documentacion adicional
- Flujo Power Automate: `docs/power-automate-flow.md`
- Deploy en Render: `docs/deploy-render.md`
