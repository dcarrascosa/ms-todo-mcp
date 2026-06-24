# Examples

## `claude_desktop_config.json`

A complete, ready-to-edit MCP client configuration that runs this server over
stdio using your **default Python** interpreter on `PATH` (no virtualenv path
hardcoded).

### Use it

1. Install the package into your default Python (no active virtualenv):

   ```bash
   pip install -e .        # from the repo root, or: pip install .
   ```

2. Sign in once so the token cache exists:

   ```bash
   ms-todo-mcp login
   ```

3. Copy the contents of `claude_desktop_config.json` into your client's config
   file and replace the two `REPLACE_WITH_*` placeholders with the
   **Application (client) ID** and **Directory (tenant) ID** from your Entra app
   registration (see the main README, step 1).

   Config file location:
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
   - **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

   If the file already has other servers, merge the `"ms-todo"` entry into the
   existing `"mcpServers"` object instead of overwriting it.

4. Restart the app.

### Variants

- **`python` not found / multiple Pythons**: set `"command": "python3"` or the
  absolute interpreter path (e.g. `C:\\Python312\\python.exe`).
- **Use the console script instead** (Windows `Scripts\` folder must be on
  `PATH`):

  ```json
  { "command": "ms-todo-mcp", "args": [] }
  ```

- **Isolation without absolute paths — pipx** (recommended if you don't want a
  global install): `pipx install .` keeps the package in its own environment and
  puts the command on your `PATH`. Then use `{ "command": "ms-todo-mcp", "args": [] }`.
- **Manual virtualenv** (also isolated, but fiddlier): point `command` at the venv
  interpreter by absolute path, e.g.
  `C:\\path\\to\\ms-todo-mcp\\.venv\\Scripts\\python.exe` with
  `"args": ["-m", "ms_todo_mcp"]`. Plain `python` won't see a venv package.

### Notes

- Credentials are set in `env` here on purpose: when the client launches the
  server, the repo `.env` is **n