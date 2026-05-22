import asyncio
import logging
import os
import random
from typing import Any

import httpx

from app.domain.exceptions import GraphConfigError

logger = logging.getLogger(__name__)


def _http_timeout() -> httpx.Timeout:
    """Timeouts largos para Graph (búsquedas, PDFs, subida Excel)."""
    try:
        total = float(os.getenv("GRAPH_HTTP_TIMEOUT_SECONDS", "300"))
    except ValueError:
        total = 300.0
    return httpx.Timeout(total, connect=min(60.0, total))


class MsGraphClient:
    """Adaptador secundario: Microsoft Graph via HTTP (client_credentials)."""

    def __init__(self) -> None:
        self.tenant_id = os.getenv("GRAPH_TENANT_ID", "").strip()
        self.client_id = os.getenv("GRAPH_CLIENT_ID", "").strip()
        self.client_secret = os.getenv("GRAPH_CLIENT_SECRET", "").strip()
        self.scope = os.getenv("GRAPH_SCOPE", "https://graph.microsoft.com/.default").strip()
        self.base_url = os.getenv("GRAPH_BASE_URL", "https://graph.microsoft.com/v1.0").strip()
        self._timeout = _http_timeout()

    def _validate_config(self) -> None:
        missing = []
        if not self.tenant_id:
            missing.append("GRAPH_TENANT_ID")
        if not self.client_id:
            missing.append("GRAPH_CLIENT_ID")
        if not self.client_secret:
            missing.append("GRAPH_CLIENT_SECRET")

        if missing:
            raise GraphConfigError(
                f"Missing environment variables for Graph: {', '.join(missing)}"
            )

    async def _get_access_token(self) -> str:
        self._validate_config()
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        token_payload = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": self.scope,
        }

        async with httpx.AsyncClient(timeout=self._timeout) as client:
            response = await client.post(
                token_url,
                data=token_payload,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
            )
            response.raise_for_status()
            token_data = response.json()

        access_token = token_data.get("access_token", "")
        if not access_token:
            raise GraphConfigError("Unable to obtain Graph access token.")
        return access_token

    async def get(self, endpoint: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        token = await self._get_access_token()
        url = f"{self.base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        headers = {"Authorization": f"Bearer {token}"}

        async with httpx.AsyncClient(timeout=self._timeout) as client:
            response = await client.get(url, headers=headers, params=params)
            response.raise_for_status()
            return response.json()

    async def put_bytes(
        self,
        endpoint: str,
        content: bytes,
        content_type: str = "application/octet-stream",
    ) -> dict[str, Any]:
        """Sube contenido. Reintenta en 423 Locked (archivo abierto en Excel / check-out)."""
        try:
            max_retries = max(1, int(os.getenv("GRAPH_PUT_MAX_RETRIES", "6")))
        except ValueError:
            max_retries = 6
        try:
            base_delay = float(os.getenv("GRAPH_PUT_RETRY_BASE_SECONDS", "3"))
        except ValueError:
            base_delay = 3.0

        token = await self._get_access_token()
        url = f"{self.base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": content_type,
        }

        last_response: httpx.Response | None = None
        async with httpx.AsyncClient(timeout=self._timeout) as client:
            for attempt in range(max_retries):
                response = await client.put(url, headers=headers, content=content)
                last_response = response
                if response.status_code != 423:
                    response.raise_for_status()
                    return response.json()
                if attempt < max_retries - 1:
                    delay = min(90.0, base_delay * (2**attempt) + random.uniform(0, 1.5))
                    await asyncio.sleep(delay)
                    token = await self._get_access_token()
                    headers["Authorization"] = f"Bearer {token}"

        assert last_response is not None
        detail = ""
        try:
            body = last_response.text
            if body:
                detail = f" Cuerpo: {body[:500]}"
        except Exception:
            pass
        raise httpx.HTTPStatusError(
            (
                "423 Locked: el archivo en SharePoint está bloqueado (suele pasar si Excel "
                "lo tiene abierto o hay check-out). Cierra el libro, libera la versión en "
                "SharePoint y reintenta."
                f"{detail}"
            ),
            request=last_response.request,
            response=last_response,
        )

    async def get_bytes(self, endpoint: str, params: dict[str, Any] | None = None) -> bytes:
        token = await self._get_access_token()
        url = f"{self.base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        headers = {"Authorization": f"Bearer {token}"}

        async with httpx.AsyncClient(timeout=self._timeout, follow_redirects=True) as client:
            response = await client.get(url, headers=headers, params=params)
            response.raise_for_status()
            return response.content

    async def delete(self, endpoint: str) -> None:
        """Elimina un ítem del drive por ruta (p. ej. .../root:/ruta/al/archivo.pdf:)."""
        token = await self._get_access_token()
        url = f"{self.base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        headers = {"Authorization": f"Bearer {token}"}

        async with httpx.AsyncClient(timeout=self._timeout) as client:
            response = await client.delete(url, headers=headers)
            response.raise_for_status()

    async def post_json(self, endpoint: str, body: dict[str, Any]) -> tuple[dict[str, Any], int]:
        """
        POST JSON a Graph (p. ej. sendMail). sendMail responde 202 Accepted con cuerpo vacío.
        Devuelve (cuerpo_json_o_vacio, código_http).
        """
        token = await self._get_access_token()
        url = f"{self.base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

        async with httpx.AsyncClient(timeout=self._timeout) as client:
            response = await client.post(url, headers=headers, json=body)
            code = response.status_code
            logger.info("Graph POST %s -> HTTP %s", endpoint.split("?", 1)[0], code)
            try:
                response.raise_for_status()
            except httpx.HTTPStatusError:
                err_txt = (response.text or "").strip()[:4000]
                logger.warning("Graph POST error %s: %s", code, err_txt or "(sin cuerpo)")
                raise
            text = (response.text or "").strip()
            if not text:
                return {}, code
            try:
                return response.json(), code
            except Exception:
                return {}, code
