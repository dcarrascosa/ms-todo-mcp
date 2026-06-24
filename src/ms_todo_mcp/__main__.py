"""Entry point for the Microsoft To Do MCP server.

Usage:
    ms-todo-mcp                 # run the MCP server over stdio (default)
    ms-todo-mcp --http [--port 8000]   # run over streamable HTTP
    ms-todo-mcp login           # interactive one-time sign-in (device code)
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from pathlib import Path


def _load_dotenv() -> None:
    """Minimal .env loader (no external dependency).

    Loads KEY=VALUE lines from a .env file in the current directory without
    overriding variables already present in the environment.
    """
    path = Path(os.environ.get("MS_TODO_DOTENV", ".env"))
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def main(argv: list[str] | None = None) -> int:
    # Logs go to stderr — stdout is reserved for the JSON-RPC stdio channel.
    logging.basicConfig(
        level=os.environ.get("MS_TODO_LOG_LEVEL", "INFO"),
        stream=sys.stderr,
        format="%(asctime)s %(name)s %(levelname)s %(message)s",
    )
    _load_dotenv()

    parser = argparse.ArgumentParser(prog="ms-todo-mcp", description="Microsoft To Do MCP server")
    sub = parser.add_subparsers(dest="command")
    sub.add_parser("login", help="Interactive one-time sign-in (device code flow)")

    parser.add_argument("--http", action="store_true", help="Use streamable HTTP transport")
    parser.add_argument("--port", type=int, default=8000, help="Port for --http (default 8000)")
    args = parser.parse_args(argv)

    if args.command == "login":
        # Imported lazily so other commands don't require msal config to parse.
        from .auth import token_provider_from_env

        token_provider_from_env().login()
        print("Sign-in complete. You can now run the server.", file=sys.stderr)
        return 0

    from .server import mcp

    if args.http:
        mcp.run(transport="streamable-http", port=args.port)
    else:
        mcp.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
