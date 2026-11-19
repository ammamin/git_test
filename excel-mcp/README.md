# Excel MCP server

A small remote MCP server that lets Claude (web, desktop and the Android/iOS apps)
read and edit one Excel workbook that lives on your server - for example a todo
list. The row tools work with whatever column headers your sheet has.

## Tools

| Tool | What it does |
| --- | --- |
| `workbook_info` | Sheets, used ranges and header rows |
| `list_rows` | List rows as records by column header, with filters (e.g. `Status` = `Open`) |
| `add_row` | Add a row by column header, e.g. `{"Task": "Buy milk", "Status": "Open"}` |
| `update_row` | Change columns of one row, e.g. mark a todo done |
| `read_range` | Read a range (values, or formulas with `show_formulas`) |
| `search` | Find cells containing a text |
| `write_cells` | Set individual cells, e.g. `{"B2": 42, "C2": "=B2*2"}` |
| `write_range` | Write a block of rows from a top-left cell |
| `append_rows` | Add rows after the last used row |
| `insert_rows` / `delete_rows` | Insert or remove whole rows |
| `add_sheet` | Create a new sheet |

Every edit first copies the workbook into `.<filename>.backups/` next to it
(last 10 kept) and then replaces the file atomically.

## 1. Install on the server

```bash
cd /opt && git clone <this repo> && cd git_test/excel-mcp
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## 2. Run it

```bash
export MCP_PUBLIC_URL=https://excel.example.com   # the public HTTPS address of this server
export MCP_PASSWORD='a long password only you know'   # 12+ characters
export EXCEL_FILE=/project/todo.xlsx              # this is also the default
.venv/bin/python server.py
```

The server listens on `127.0.0.1:8000` (change with `MCP_HOST` / `MCP_PORT`).
The user running it needs write access to the workbook **and its folder**.

To keep it running, use a systemd unit (`/etc/systemd/system/excel-mcp.service`):

```ini
[Unit]
Description=Excel MCP server
After=network.target

[Service]
User=excel
WorkingDirectory=/opt/git_test/excel-mcp
Environment=EXCEL_FILE=/project/todo.xlsx
Environment=MCP_PUBLIC_URL=https://excel.example.com
Environment=MCP_PASSWORD=a-long-password-only-you-know
ExecStart=/opt/git_test/excel-mcp/.venv/bin/python server.py
Restart=on-failure

[Install]
WantedBy=multi-user.target
```

## 3. Put HTTPS in front

Claude connects from Anthropic's cloud, not from your phone, so the server must be
reachable from the public internet over HTTPS with a valid certificate. With
[Caddy](https://caddyserver.com) that is two lines (`/etc/caddy/Caddyfile`):

```
excel.example.com {
    reverse_proxy 127.0.0.1:8000
}
```

No domain or open ports? A tunnel such as Cloudflare Tunnel or Tailscale Funnel
pointing at `http://127.0.0.1:8000` works too.

## 4. Add it to Claude

1. Open **claude.ai** in a browser → **Settings → Connectors → Add custom connector**.
2. Name it (e.g. "Todo") and paste `https://excel.example.com/mcp`. Leave the
   optional OAuth client ID/secret fields empty - Claude registers itself.
3. Click **Connect**. Your server's sign-in page opens; type `MCP_PASSWORD` and
   press **Allow access**.
4. In the Android app, open a chat, enable the connector from the tools menu and ask
   e.g. "What's still open on my todo list?" or "Add 'buy milk' and mark the rent one done".

Connectors are tied to your account, so once added on the web it shows up in the
mobile app.

## Security (OAuth)

The server is its own OAuth 2.1 authorization server (`oauth.py`), for one user:

- `/mcp` rejects every request without a valid access token.
- Claude gets a token only after you enter `MCP_PASSWORD` on the sign-in page
  (authorization code flow with PKCE, dynamic client registration).
- Access tokens last 1 hour; refresh tokens rotate on every use and expire after
  90 days without use, after which you sign in again.
- Only token hashes are stored, in `oauth-state.json` (mode 0600, path changeable
  with `MCP_STATE_FILE`). Delete that file and restart to disconnect everything.
- Five wrong passwords lock the sign-in page for 15 minutes. The lock is global,
  so someone hammering the page can also delay *your* sign-in; already-connected
  Claude sessions are not affected.
- Changing `MCP_PASSWORD` does not revoke existing tokens - delete the state file too.

## Limitations

- Editing goes through `openpyxl`. Cell values, formulas, styles and most formatting
  survive a save, but **charts, images and some pivot-table details may be dropped**.
  Try it on a copy first if the workbook has those.
- After an edit, formula results are not recalculated until the file is opened in
  Excel/LibreOffice; until then `read_range` shows the formula text for those cells.
- `insert_rows` / `delete_rows` do not re-point formulas that reference moved rows.
- Do not keep the file open in desktop Excel on the server while Claude edits it.
- `.xls` (the old binary format) is not supported; `.xlsx` and `.xlsm` are.
