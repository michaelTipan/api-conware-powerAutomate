# Despliegue en Render

## Requisitos previos
- Código en un repositorio Git (**GitHub**, GitLab o Bitbucket) accesible desde Render.
- Cuenta en [Render](https://render.com).

## Opción A — Web Service manual (recomendado la primera vez)

1. En el dashboard de Render: **New** → **Web Service**.
2. Conecta el repositorio y elige la rama (por ejemplo `main`).
3. Configuración sugerida:
   - **Name:** el que prefieras (ej. `conware-excel-api`).
   - **Region:** la más cercana a tus usuarios.
   - **Runtime:** `Python 3`.
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `uvicorn app.main:app --host 0.0.0.0 --port $PORT`
   - **Instance type:** **Free** (pruebas) o **Starter** u otro plan si lo necesitas.
4. **Create Web Service** y espera a que termine el build y el deploy.
5. La URL pública será algo como `https://<nombre-del-servicio>.onrender.com`.

## Opción B — Blueprint (`render.yaml`)

1. **New** → **Blueprint**.
2. Selecciona el mismo repositorio; Render detectará `render.yaml` en la raíz.
3. Revisa el nombre del servicio y el plan, luego confirma la creación.

## Comprobar que funciona
- `GET https://<tu-url>/health` → debe responder `{"status":"ok"}`.
- `POST https://<tu-url>/parse-excel` con **multipart/form-data**:
  - Campo **`file`**: archivo `.xlsx`.
  - Campo opcional **`run_id`**: texto.
- Respuesta JSON: `columns`, `rows`, `row_count`, `file_name`, `elapsed_ms`, etc.

## Power Automate
- Usar **HTTP** con **POST** y cuerpo **form-data** (no JSON con Base64).
- Adjuntar el Excel en la clave **`file`**.

## Notas
- `$PORT` lo inyecta Render; no lo sustituyas por un número fijo en **Start command**.
- No subas `.env` con secretos al repo; usa **Environment** en Render solo si más adelante añades variables.
