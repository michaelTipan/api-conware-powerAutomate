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

## 3) Flujo validacion de pagos (COMWARE)
Ver `docs/payment-validation-production-flow.md`. En produccion use Generate/Finalize/Notify/Merge/Amortization; no use `validate-payment-report`.

## 4) Variables de entorno
Configure en el dashboard de Render según `docs/render-env-variables.md`. No commitear `.env` ni `.env.local` (locales, ignorados por git).

## 4) Contrato desde Power Automate (parse-excel)
- `file_name` (string)
- `excel_base64` (string)
- `run_id` (string, opcional; puede ir vacio `""`)
- `column_letter` (string, opcional; por defecto en API es `C`)
