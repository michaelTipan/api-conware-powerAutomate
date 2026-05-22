# Despliegue en Render

## 1) Crear servicio web
- Crear nuevo `Web Service` conectado a este repo.
- Runtime: Python.
- Build command: `pip install -r requirements.txt`.
- Start command: `uvicorn app.main:app --host 0.0.0.0 --port $PORT`.

## 2) Verificaciones post despliegue
- `GET /health` debe responder `{"status":"ok"}`.
- Probar `POST /parse-excel` con un Excel de prueba.
- Validar que Power Automate reciba JSON correcto.

## 3) Contrato desde Power Automate (body)
- `file_name` (string)
- `excel_base64` (string)
- `run_id` (string, opcional; puede ir vacio `""`)
- `column_letter` (string, opcional; por defecto en API es `C`)
