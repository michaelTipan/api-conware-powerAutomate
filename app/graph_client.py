"""Compatibilidad: usar `app.adapters.secondary.ms_graph_client.MsGraphClient`."""

from app.adapters.secondary.ms_graph_client import MsGraphClient
from app.domain.exceptions import GraphConfigError

GraphClient = MsGraphClient

__all__ = ["GraphClient", "GraphConfigError", "MsGraphClient"]
