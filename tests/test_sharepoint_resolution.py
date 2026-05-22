import asyncio
import pytest
from app.application.sharepoint_resolution import encode_graph_drive_path

try:
    from app.application.sharepoint_resolution import resolve_sharepoint_path
except ImportError:
    resolve_sharepoint_path = None

class DummyGraphClient:
    async def get(self, endpoint, params=None):
        if "/sites" in endpoint and params and "search" in params:
            return {"value": [{"id": "dummy_site_id"}]}
        if "/drives" in endpoint:
            return {"value": [{"id": "dummy_drive_id", "name": "Documentos"}]}
        return {}

def test_encode_graph_drive_path():
    encoded = encode_graph_drive_path("Carpetas/Mi Archivo.xlsx")
    assert encoded == "Carpetas/Mi%20Archivo.xlsx"

def test_resolve_sharepoint_path():
    if resolve_sharepoint_path is None:
        pytest.fail("resolve_sharepoint_path no implementado")
        
    async def run_test():
        client = DummyGraphClient()
        result = await resolve_sharepoint_path(
            client, 
            site_search="SITIO", 
            drive_name="Documentos", 
            path="Mis Pagos/PENDIENTES.xlsx"
        )
        
        assert result["site_id"] == "dummy_site_id"
        assert result["drive_id"] == "dummy_drive_id"
        assert result["file_path"] == "Mis Pagos/PENDIENTES.xlsx"
        assert result["path_encoded"] == "Mis%20Pagos/PENDIENTES.xlsx"
        
    asyncio.run(run_test())
