"""Thin async client for the Microsoft Graph To Do endpoints."""

from __future__ import annotations

from typing import Any, Optional

import httpx

from .auth import TokenProvider

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Safety cap so a limit-based list never follows nextLink forever.
_MAX_PAGES = 20


class GraphError(RuntimeError):
    """An error returned by Microsoft Graph (or a transport failure)."""

    def __init__(self, status: Optional[int], message: str) -> None:
        self.status = status
        super().__init__(message)


def _extract_error(response: httpx.Response) -> str:
    try:
        payload = response.json()
        err = payload.get("error", {})
        msg = err.get("message") or err.get("code") or response.text
        return str(msg)
    except Exception:  # noqa: BLE001 - fall back to raw text
        return response.text or f"HTTP {response.status_code}"


class GraphClient:
    """Authenticated async wrapper around the Graph REST API."""

    def __init__(
        self,
        token_provider: TokenProvider,
        base_url: str = GRAPH_BASE_URL,
        timeout: float = 30.0,
    ) -> None:
        self._tp = token_provider
        self._base = base_url.rstrip("/")
        self._timeout = timeout

    async def request(
        self,
        method: str,
        path: str,
        *,
        params: Optional[dict[str, Any]] = None,
        json_body: Optional[dict[str, Any]] = None,
    ) -> dict[str, Any]:
        """Perform a single Graph request and return the parsed JSON (or {})."""
        token = await self._tp.get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        }
        url = path if path.startswith("http") else f"{self._base}{path}"
        try:
            async with httpx.AsyncClient(timeout=self._timeout) as client:
                response = await client.request(
                    method, url, params=params, json=json_body, headers=headers
                )
        except httpx.TimeoutException as exc:
            raise GraphError(None, "Request to Microsoft Graph timed out.") from exc
        except httpx.HTTPError as exc:
            raise GraphError(None, f"Network error talking to Microsoft Graph: {exc}") from exc

        if response.status_code == 204 or not response.content:
            return {}
        if response.status_code >= 400:
            raise GraphError(response.status_code, _extract_error(response))
        return response.json()

    async def get_collection(
        self,
        path: str,
        *,
        params: Optional[dict[str, Any]] = None,
        limit: int = 50,
    ) -> tuple[list[dict[str, Any]], bool]:
        """Fetch up to `limit` items from a paged collection.

        Follows @odata.nextLink until `limit` items are collected or there are
        no more pages. Returns (items, has_more).
        """
        params = dict(params or {})
        params.setdefault("$top", min(limit, 100))
        items: list[dict[str, Any]] = []
        next_url: Optional[str] = None
        pages = 0

        while True:
            if next_url:
                data = await self.request("GET", next_url)
            else:
                data = await self.request("GET", path, params=params)
            items.extend(data.get("value", []))
            next_url = data.get("@odata.nextLink")
            pages += 1
            if len(items) >= limit or not next_url or pages >= _MAX_PAGES:
                break

        has_more = bool(next_url) or len(items) > limit
        return items[:limit], has_more
