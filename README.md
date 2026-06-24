# Microsoft To Do MCP server

An [MCP](https://modelcontextprotocol.io) server that lets an LLM client (Claude
Desktop, Cowork, etc.) read and manage your **Microsoft To Do** tasks through the
**Microsoft Graph** To Do API.

It exposes seven tools:

| Tool | What it does | Write? |
|------|--------------|:------:|
| `todo_list_lists` | List your task lists | – |
| `todo_list_tasks` | List tasks in a list (by id or name), optional status filter | – |
| `todo_create_task` | Create a task (title, note, due date, importance) | ✓ |
| `todo_update_task` | Update title / note / due date / importance / status | ✓ |
| `todo_complete_task` | Mark a task completed | ✓ |
| `todo_delete_task` | Delete a task (permanent) | ✓ |
| `todo_create_list` | Create a new task list | ✓ |

## How auth works

Microsoft Graph **does not support application-only (daemon) permissions for
creating or updating To Do tasks** — only **delegated** permissions work. This
server therefore signs in as *you* using the OAuth2 **device-code flow** and
caches the resulting access + refresh tokens on disk, so you only sign in once.

> The token cache contains live credentials. It is written to
> `~/.ms_todo_mcp/token_cache.json` by default and is **git-ignored**. Never
> commit it.

## 1. Register an app in Microsoft Entra ID

1. [Entra admin center](https://entra.microsoft.com) → **App registrations** → **New registration**.
2. Name it e.g. `ms-todo-mcp`. For *Supported account types* pick
   **Accounts in this organizational directory only** (single tenant) unless you
   need broader access.
3. No redirect URI is required for the device-code flow. Click **Register**.
4. **Authentication** → **Advanced settings** → set
   **Allow public client flows** = **Yes** (required for device-code).
5. **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions** → add **`Tasks.ReadWrite`**. If your tenant requires
   it, click **Grant admin consent**.
6. Copy the **Application (client) ID** and **Directory (tenant) ID** from the
   app's *Overview* page.

## 2. Configure

```bash
cp .env.example .env
# edit .env and set MS_TODO_CLIENT_ID and MS_TODO_TENANT_ID
```

| Variable | Required | Notes |
|----------|:--------:|-------|
| `MS_TODO_CLIENT_ID` | ✓ | Application (client) ID |
| `MS_TODO_TENANT_ID` | ✓ for work/school | Tenant ID/domain, or `organizations` / `common` |
| `MS_TODO_TOKEN_CACHE` | – | Override token-cache path |
| `MS_TODO_TIMEZONE` | – | IANA tz for due dates (default `UTC`) |

## 3. Install

Install the package into the **Python interpreter your MCP client will launch**.
By default that is plain `python` on your `PATH`. Because the client starts the
server from its own working directory, the simplest, path-free setup is to
install into that default interpreter — **not** a virtualenv:

```bash
pip install -e .          # from the repo root (use `pip install .` for non-editable)
```

On Windows this also places an `ms-todo-mcp` launcher in your Python `Scripts\`
folder.

> **Want isolation instead of a global install?** Don't reach for a bare venv
> with the default config — `"command": "python"` won't see a venv package.
> Use **pipx** (`pipx install .`), which isolates the package *and* puts the
> `ms-todo-mcp` command on your `PATH`; then set `"command": "ms-todo-mcp"` in the
> config. (A manual venv also works, but only if the config points at that venv's
> `python.exe` by absolute path — see [`examples/README.md`](./examples/README.md).)

## 4. Sign in once

```bash
ms-todo-mcp login
```

Follow the printed instructions: open the URL, enter the code, sign in. The token
cache is then stored and reused.

## 5. Run

```bash
ms-todo-mcp            # stdio transport (for local MCP clients)
ms-todo-mcp --http     # streamable HTTP on :8000 (for remote use)
```

## 6. Add it to your MCP client

For a local **stdio** client (e.g. Claude Desktop / Cowork custom connector), add
an entry like the following. A ready-to-edit copy lives at
[`examples/claude_desktop_config.json`](./examples/claude_desktop_config.json) and
uses your **default Python** on `PATH` (no virtualenv path hardcoded):

```json
{
  "mcpServers": {
    "ms-todo": {
      "command": "python",
      "args": ["-m", "ms_todo_mcp"],
      "env": {
        "MS_TODO_CLIENT_ID": "<your-client-id>",
        "MS_TODO_TENANT_ID": "<your-tenant-id>",
        "MS_TODO_TIMEZONE": "Europe/Madrid"
      }
    }
  }
}
```

Notes:

- The config file lives at `%APPDATA%\Claude\claude_desktop_config.json` (Windows)
  or `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS).
  Restart the app after editing it.
- For `python -m ms_todo_mcp` to resolve, install the package into that **default**
  Python — run `pip install -e .` (or `pip install .`) **without** an active
  virtualenv. If `python` isn't your launcher, use `python3` or the absolute
  interpreter path. On Windows you can instead set `"command": "ms-todo-mcp"`
  (the console script in your Python `Scripts` folder).
- **Run `ms-todo-mcp login` once first** so the token cache exists before the
  client launches the server (the device-code prompt goes to stderr, not chat).
- Credentials live in `env` here because the repo `.env` is **not** read when the
  client starts the server (its working directory isn't the project).

## Usage examples (what to ask your assistant)

- "List my To Do lists."
- "Show the unfinished tasks in my 'Tasks' list."
- "Add 'Send Microsoft invoices to Diego' to my Work list, due Friday, high importance."
- "Mark the task about UAT as completed."

## Development

```bash
pip install -e ".[dev]"
ruff check .
python -m py_compile src/ms_todo_mcp/*.py
```

## Project layout

```
ms-todo-mcp/
├── pyproject.toml
├── README.md
├── CONTRIBUTING.md
├── LICENSE                 # MIT
├── .env.example
├── .gitignore
├── examples/
│   ├── README.md
│   └── claude_desktop_config.json
└── src/
    └── ms_todo_mcp/
        ├── __init__.py
        ├── __main__.py
        ├── auth.py
        ├── graph.py
        └── server.py
```

## Releases

Releases are automated with
[release-please](https://github.com/googleapis/release-please-action). You do
**not** create tags or GitHub releases by hand.

The version is derived from [Conventional Commits](https://www.conventionalcommits.org/).
Because `main` is squash-merged, **the PR title is what counts** — prefix it
accordingly:

- `fix: ...` → patch bump (`0.1.0` → `0.1.1`)
- `feat: ...` → minor bump (`0.1.0` → `0.2.0`)
- `feat!: ...` or a `BREAKING CHANGE:` footer → major bump (`0.1.0` → `1.0.0`)
- `docs: ...`, `chore: ...`, `ci: ...`, `refactor: ...` → no release

Flow:

1. Merge PRs into `main` using Conventional Commit titles.
2. release-please opens (and keeps updating) a **release PR** titled
   `chore: release X.Y.Z`, bumping the version in `pyproject.toml` and
   `src/ms_todo_mcp/__init__.py` and updating `CHANGELOG.md`.
3. When you merge that release PR, release-please creates the `vX.Y.Z` git tag
   and the GitHub Release with auto-generated notes.
