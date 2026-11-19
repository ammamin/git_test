#!/usr/bin/env bash
# One-shot installer: runs the Excel MCP server as a systemd service and publishes it
# through a Cloudflare Tunnel. Safe to re-run. Debian/Ubuntu-style Linux with systemd.
#
#   sudo ./setup.sh
#
# Optional overrides (environment variables):
#   DOMAIN       public hostname            (default todo.basehomemedia.stream)
#   EXCEL_FILE   workbook to serve          (default /project/todo.xlsx)
#   MCP_PORT     local port for the server  (default 8377)
#   RUN_AS       user the server runs as    (default: owner of the workbook)
#   TUNNEL_NAME  Cloudflare tunnel name     (default todo-mcp)
set -euo pipefail

DOMAIN=${DOMAIN:-todo.basehomemedia.stream}
EXCEL_FILE=${EXCEL_FILE:-/project/todo.xlsx}
MCP_PORT=${MCP_PORT:-8377}
TUNNEL_NAME=${TUNNEL_NAME:-todo-mcp}
APP_DIR=$(cd "$(dirname "$0")" && pwd)
ENV_FILE=/etc/excel-mcp.env
STATE_DIR=/var/lib/excel-mcp
CF_DIR=/etc/cloudflared
CF_CONFIG=$CF_DIR/$TUNNEL_NAME.yml

say() { printf '\n\033[1m== %s\033[0m\n' "$*"; }
die() { printf '\nERROR: %s\n' "$*" >&2; exit 1; }

[ "$(id -u)" = 0 ] || die "Run with sudo: sudo $0"
command -v systemctl >/dev/null || die "systemd is required"
[ -f "$EXCEL_FILE" ] || die "Workbook not found: $EXCEL_FILE (set EXCEL_FILE=... to override)"
RUN_AS=${RUN_AS:-$(stat -c %U "$EXCEL_FILE")}
id "$RUN_AS" >/dev/null 2>&1 || die "User '$RUN_AS' does not exist (set RUN_AS=...)"
sudo -u "$RUN_AS" test -w "$EXCEL_FILE" -a -w "$(dirname "$EXCEL_FILE")" ||
    die "User '$RUN_AS' needs write access to $EXCEL_FILE and its folder (backups are stored beside it)"
sudo -u "$RUN_AS" test -r "$APP_DIR/server.py" ||
    die "User '$RUN_AS' cannot read $APP_DIR - move this folder somewhere readable, e.g. /opt/excel-mcp"

# ---- 1. Python environment -------------------------------------------------
say "Installing Python dependencies"
if ! python3 -m venv --help >/dev/null 2>&1 || ! python3 -c 'import ensurepip' 2>/dev/null; then
    apt-get update -qq && apt-get install -y -qq python3-venv
fi
python3 -c 'import sys; sys.exit(sys.version_info < (3, 10))' || die "Python 3.10 or newer is required"
[ -x "$APP_DIR/.venv/bin/python" ] || python3 -m venv "$APP_DIR/.venv"
"$APP_DIR/.venv/bin/pip" install -q --upgrade pip
"$APP_DIR/.venv/bin/pip" install -q -r "$APP_DIR/requirements.txt"

# ---- 2. Password and service -----------------------------------------------
say "Configuring the excel-mcp service (runs as '$RUN_AS', port $MCP_PORT)"
if [ -f "$ENV_FILE" ] && grep -q '^MCP_PASSWORD=' "$ENV_FILE"; then
    PASSWORD=$(sed -n 's/^MCP_PASSWORD=//p' "$ENV_FILE")
    echo "Keeping the existing sign-in password from $ENV_FILE"
else
    while :; do
        read -rsp "Choose the sign-in password for Claude (12+ characters): " PASSWORD; echo
        [ "${#PASSWORD}" -ge 12 ] || { echo "Too short."; continue; }
        read -rsp "Repeat it: " AGAIN; echo
        [ "$PASSWORD" = "$AGAIN" ] && break
        echo "They do not match."
    done
fi
if systemctl is-active -q excel-mcp; then systemctl stop excel-mcp; fi
if ss -ltn 2>/dev/null | grep -q ":$MCP_PORT "; then
    die "Port $MCP_PORT is already in use. Re-run with another one: sudo MCP_PORT=8378 $0"
fi

install -d -o "$RUN_AS" -m 700 "$STATE_DIR"
umask 077
cat >"$ENV_FILE" <<EOF
EXCEL_FILE=$EXCEL_FILE
MCP_PUBLIC_URL=https://$DOMAIN
MCP_PORT=$MCP_PORT
MCP_STATE_FILE=$STATE_DIR/oauth-state.json
MCP_PASSWORD=$PASSWORD
EOF
umask 022

cat >/etc/systemd/system/excel-mcp.service <<EOF
[Unit]
Description=Excel MCP server ($EXCEL_FILE)
After=network.target

[Service]
User=$RUN_AS
WorkingDirectory=$APP_DIR
EnvironmentFile=$ENV_FILE
ExecStart=$APP_DIR/.venv/bin/python $APP_DIR/server.py
Restart=on-failure
RestartSec=3

[Install]
WantedBy=multi-user.target
EOF
systemctl daemon-reload
systemctl enable -q --now excel-mcp
sleep 3
systemctl is-active -q excel-mcp || { journalctl -u excel-mcp -n 20 --no-pager; die "excel-mcp failed to start"; }
curl -fsS -o /dev/null -H "Host: $DOMAIN" "http://127.0.0.1:$MCP_PORT/.well-known/oauth-authorization-server" ||
    die "excel-mcp is running but not answering on port $MCP_PORT"
echo "excel-mcp is running."

# ---- 3. Cloudflare Tunnel --------------------------------------------------
say "Setting up the Cloudflare Tunnel '$TUNNEL_NAME' for $DOMAIN"
if ! command -v cloudflared >/dev/null; then
    command -v dpkg >/dev/null || die "Install cloudflared manually (https://pkg.cloudflare.com), then re-run"
    deb=$(mktemp --suffix=.deb)
    curl -fsSL -o "$deb" \
        "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-$(dpkg --print-architecture).deb"
    dpkg -i "$deb" >/dev/null && rm -f "$deb"
fi
CLOUDFLARED=$(command -v cloudflared)

# Tunnel credentials live in root's ~/.cloudflared. This never touches an existing
# cloudflared service or /etc/cloudflared/config.yml - it uses its own config and unit.
export HOME=/root
if [ ! -f /root/.cloudflared/cert.pem ]; then
    echo "Open the link below in a browser, log in to Cloudflare and pick ${DOMAIN#*.}:"
    "$CLOUDFLARED" tunnel login
fi
tunnel_id() {
    "$CLOUDFLARED" tunnel list --output json 2>/dev/null | python3 -c '
import json, sys
print(next((t["id"] for t in json.load(sys.stdin) or [] if t["name"] == sys.argv[1]), ""))' "$TUNNEL_NAME"
}
TUNNEL_ID=$(tunnel_id)
if [ -z "$TUNNEL_ID" ]; then
    "$CLOUDFLARED" tunnel create "$TUNNEL_NAME"
    TUNNEL_ID=$(tunnel_id)
fi
[ -n "$TUNNEL_ID" ] || die "Could not create or find the tunnel '$TUNNEL_NAME'"
CREDS=/root/.cloudflared/$TUNNEL_ID.json
[ -f "$CREDS" ] || die "Tunnel '$TUNNEL_NAME' exists but its credentials are not on this machine.
Delete it (cloudflared tunnel delete $TUNNEL_NAME) or re-run with TUNNEL_NAME=another-name."

"$CLOUDFLARED" tunnel route dns --overwrite-dns "$TUNNEL_NAME" "$DOMAIN"

install -d -m 755 "$CF_DIR"
cat >"$CF_CONFIG" <<EOF
tunnel: $TUNNEL_ID
credentials-file: $CREDS
ingress:
  - hostname: $DOMAIN
    service: http://127.0.0.1:$MCP_PORT
  - service: http_status:404
EOF
cat >"/etc/systemd/system/cloudflared-$TUNNEL_NAME.service" <<EOF
[Unit]
Description=Cloudflare Tunnel $TUNNEL_NAME ($DOMAIN)
After=network-online.target
Wants=network-online.target

[Service]
ExecStart=$CLOUDFLARED --no-autoupdate tunnel --config $CF_CONFIG run
Restart=on-failure
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF
systemctl daemon-reload
systemctl enable -q "cloudflared-$TUNNEL_NAME"
systemctl restart "cloudflared-$TUNNEL_NAME"

# ---- 4. Check from the outside ---------------------------------------------
say "Checking https://$DOMAIN"
ok=
for _ in $(seq 1 12); do
    if curl -fsS -m 10 "https://$DOMAIN/.well-known/oauth-protected-resource/mcp" 2>/dev/null | grep -q "$DOMAIN"; then
        ok=1; break
    fi
    sleep 5
done
if [ -n "$ok" ]; then
    cat <<EOF

Done. The server is public at https://$DOMAIN

Next, in a browser on claude.ai:
  Settings -> Connectors -> Add custom connector
  URL: https://$DOMAIN/mcp
  Click Connect and type your sign-in password.

Logs:  journalctl -u excel-mcp -f      journalctl -u cloudflared-$TUNNEL_NAME -f
EOF
else
    die "The services are installed but https://$DOMAIN is not answering yet.
DNS can take a few minutes; then check:  journalctl -u cloudflared-$TUNNEL_NAME -n 30"
fi
