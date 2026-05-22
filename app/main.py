"""
Punto de entrada ASGI: carga entorno y construye la app (adaptador primario HTTP).
Arquitectura hexagonal: dominio / aplicacion / adaptadores en subpaquetes.
"""

from dotenv import load_dotenv

load_dotenv()

from app.adapters.primary.http.app_factory import create_app

app = create_app()
