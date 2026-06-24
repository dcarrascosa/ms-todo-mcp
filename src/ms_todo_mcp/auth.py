"""Authentication for the Microsoft To Do MCP server.

Uses MSAL with the OAuth2 device-code flow (delegated permissions). Microsoft
Graph does NOT allow creating/updating To Do tasks with application-only
permissions, so a delegated (signed-in user) token is mandatory.

Tokens are cached on disk (access + refresh) so the interactive sign-in only
happens once; afterwards tokens are refreshed silently.

NOTE: this module must never write to stdout. The MCP server talks JSON-RPC
over stdout when using the stdio transport, so all human-facing messages
(including the device-code prompt) go to stderr.
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import sys
from pathlib import Path
from typing import Optional

import msal

logger = logging.getLogger("ms_todo_mcp.auth")

# Reserved scopes (openid/profile/offline_access) are added automatically by
# MSAL — do NOT list them here.
GRAPH_SCOPES = ["Tasks.ReadWrite"]

DEFAULT_CACHE_PATH = Path.home() / ".ms_todo_mcp" / "token_cache.json"


class AuthError(RuntimeError):
    """Raised when authentication cannot be completed."""


class TokenProvider:
    """Acquires and caches Microsoft Graph access tokens (delegated)."""

    def __init__(
        self,
        client_id: str,
        tenant_id: str = "organizations",
        cache_path: Optional[str | os.PathLike[str]] = None,
    ) -> None:
        if not client_id:
            raise AuthError(
                "MS_TODO_CLIENT_ID is not set. Create an Entra ID app registration "
                "and set its Application (client) ID in MS_TODO_CLIENT_ID. "
                "See the README for the full setup."
            )
        self.client_id = client_id
        self.tenant_id = tenant_id or "organizations"
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.cache_path = Path(cache_path) if cache_path else DEFAULT_CACHE_PATH

        self._cache = msal.SerializableTokenCache()
        self._load_cache()
        self._app = msal.PublicClientApplication(
            self.client_id,
            authority=self.authority,
            token_cache=self._cache,
        )
        self._lock = asyncio.Lock()

    # -- cache helpers -----------------------------------------------------
    def _load_cache(self) -> None:
        if self.cache_path.exists():
            try:
                self._cache.deserialize(self.cache_path.read_text(encoding="utf-8"))
            except Exception as exc:  # noqa: BLE001 - cache corruption is non-fatal
                logger.warning("Ignoring unreadable token cache at %s: %s", self.cache_path, exc)

    def _save_cache(self) -> None:
        if not self._cache.has_state_changed:
            return
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        self.cache_path.write_text(self._cache.serialize(), encoding="utf-8")
        try:
            os.chmod(self.cache_path, 0o600)
        except OSError:  # e.g. on Windows / unsupported filesystems
            pass

    # -- token acquisition -------------------------------------------------
    def _acquire_sync(self, *, allow_interactive: bool = True) -> str:
        """Blocking token acquisition. Run inside a thread from async code."""
        result = None
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])

        if not result:
            if not allow_interactive:
                raise AuthError(
                    "No cached credentials. Run 'ms-todo-mcp login' once to sign in."
                )
            flow = self._app.initiate_device_flow(scopes=GRAPH_SCOPES)
            if "user_code" not in flow:
                raise AuthError(
                    "Failed to start device-code flow: "
                    f"{flow.get('error_description', json.dumps(flow))}"
                )
            # Human-facing prompt -> stderr only (stdout is the JSON-RPC channel).
            print("\n" + flow["message"] + "\n", file=sys.stderr, flush=True)
            logger.info("Waiting for device-code sign-in to complete...")
            result = self._app.acquire_token_by_device_flow(flow)

        self._save_cache()

        if "access_token" not in result:
            raise AuthError(
                "Could not obtain an access token: "
                f"{result.get('error_description', result.get('error', 'unknown error'))}"
            )
        return result["access_token"]

    async def get_access_token(self) -> str:
        """Return a valid access token, refreshing or signing in as needed."""
        async with self._lock:
            return await asyncio.to_thread(self._acquire_sync, allow_interactive=True)

    def login(self) -> None:
        """Run the interactive device-code flow once and persist the token cache."""
        self._acquire_sync(allow_interactive=True)
        logger.info("Sign-in complete. Token cache stored at %s", self.cache_path)


def token_provider_from_env() -> TokenProvider:
    """Build a TokenProvider from MS_TODO_* environment variables."""
    return TokenProvider(
        client_id=os.environ.get("MS_TODO_CLIENT_ID", ""),
        tenant_id=os.environ.get("MS_TODO_TENANT_ID", "organizations"),
        cache_path=os.environ.get("MS_TODO_TOKEN_CACHE"),
    )
