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
export EXCEL_FILE=/project/todo.xlsx              # this is also the default
export MCP_SECRET=$(python3 -c "import secrets; print(secrets.token_urlsafe(32))")
export MCP_ALLOWED_HOSTS=excel.example.com     # your public hostname
echo "Connector URL: https://excel.example.com/$MCP_SECRET/mcp"
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
Environment=MCP_SECRET=put-your-long-random-secret-here
Environment=MCP_ALLOWED_HOSTS=excel.example.com
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
2. Name it (e.g. "Excel") and paste `https://excel.example.com/<MCP_SECRET>/mcp`.
3. Leave the OAuth fields empty and save.
4. In the Android app, open a chat, enable the connector from the tools menu and ask
   e.g. "What's still open on my todo list?" or "Add 'buy milk' and mark the rent one done".

Connectors are tied to your account, so once added on the web it shows up in the
mobile app.

## Security

The only thing protecting the file is the secret in the URL plus HTTPS. Treat the
full URL like a password: anyone who has it can read and change the workbook.
Rotate it by changing `MCP_SECRET` and re-adding the connector. For something
stronger, put OAuth in front (the MCP SDK supports it via `token_verifier` /
`auth_server_provider`).

## Limitations

- Editing goes through `openpyxl`. Cell values, formulas, styles and most formatting
  survive a save, but **charts, images and some pivot-table details may be dropped**.
  Try it on a copy first if the workbook has those.
- After an edit, formula results are not recalculated until the file is opened in
  Excel/LibreOffice; until then `read_range` shows the formula text for those cells.
- `insert_rows` / `delete_rows` do not re-point formulas that reference moved rows.
- Do not keep the file open in desktop Excel on the server while Claude edits it.
- `.xls` (the old binary format) is not supported; `.xlsx` and `.xlsm` are.
