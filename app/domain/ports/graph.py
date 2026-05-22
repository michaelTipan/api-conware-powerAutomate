from typing import Any, Protocol, runtime_checkable


@runtime_checkable
class GraphApiPort(Protocol):
    """Puerto secundario: acceso a Microsoft Graph (implementado por HTTP/httpx)."""

    async def get(self, endpoint: str, params: dict[str, Any] | None = None) -> dict[str, Any]: ...

    async def get_bytes(self, endpoint: str, params: dict[str, Any] | None = None) -> bytes: ...

    async def put_bytes(
        self,
        endpoint: str,
        content: bytes,
        content_type: str = "application/octet-stream",
    ) -> dict[str, Any]: ...

    async def post_json(self, endpoint: str, body: dict[str, Any]) -> tuple[dict[str, Any], int]: ...

    async def delete(self, endpoint: str) -> None: ...
