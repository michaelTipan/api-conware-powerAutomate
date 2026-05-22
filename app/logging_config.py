"""Configuración de logging para que los mensajes de `app.*` salgan en consola."""

import logging
import os
import sys


def configure_logging() -> None:
    level_name = os.getenv("LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)

    root = logging.getLogger()
    if not root.handlers:
        logging.basicConfig(
            level=level,
            format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
            stream=sys.stderr,
        )
    root.setLevel(level)
    logging.getLogger("app").setLevel(level)
    logging.getLogger("uvicorn.access").setLevel(logging.INFO)
