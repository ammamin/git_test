"""Self-contained single-user OAuth 2.1 authorization server for the Excel MCP server.

Claude registers itself (dynamic client registration), sends you to /login, and you
approve the connection by typing the password from MCP_PASSWORD. Claude then receives
an access token (1 hour) and a rotating refresh token (90 days, renewed on each use).

Registered clients and token hashes are kept in a small JSON state file so a restart
does not disconnect Claude. Tokens themselves are never written to disk.
"""

from __future__ import annotations

import hashlib
import hmac
import html
import json
import os
import secrets
import tempfile
import time
from pathlib import Path
from typing import Any

from mcp.server.auth.provider import (
    AccessToken,
    AuthorizationCode,
    AuthorizationParams,
    AuthorizeError,
    RefreshToken,
    RegistrationError,
    construct_redirect_uri,
)
from mcp.shared.auth import OAuthClientInformationFull, OAuthToken
from starlette.requests import Request
from starlette.responses import HTMLResponse, RedirectResponse, Response

ACCESS_TTL = 3600
REFRESH_TTL = 90 * 24 * 3600
CODE_TTL = 300
LOGIN_TTL = 600
MAX_CLIENTS = 100
MAX_FAILURES = 5
FAILURE_WINDOW = 15 * 60
SUBJECT = "owner"


def _hash(token: str) -> str:
    return hashlib.sha256(token.encode()).hexdigest()


class SingleUserOAuthProvider:
    def __init__(self, *, public_url: str, resource_url: str, password: str, state_file: Path):
        self.public_url = public_url.rstrip("/")
        self.resource_url = resource_url
        self._password = password.encode()
        self._state_file = state_file
        # Persistent: clients by id; grants by id -> {client_id, scopes, access, access_exp, refresh, refresh_exp}
        self._clients: dict[str, dict[str, Any]] = {}
        self._grants: dict[str, dict[str, Any]] = {}
        # In memory only (short-lived)
        self._pending: dict[str, dict[str, Any]] = {}
        self._codes: dict[str, AuthorizationCode] = {}
        self._failures: list[float] = []
        self._load()

    # ---- persistence -------------------------------------------------------

    def _load(self) -> None:
        if self._state_file.exists():
            data = json.loads(self._state_file.read_text())
            self._clients = data.get("clients", {})
            self._grants = data.get("grants", {})

    def _persist(self) -> None:
        now = time.time()
        self._grants = {k: g for k, g in self._grants.items() if g["refresh_exp"] > now}
        fd, tmp = tempfile.mkstemp(dir=self._state_file.parent, prefix=".oauth-")
        with os.fdopen(fd, "w") as f:  # mkstemp creates the file with mode 0600
            json.dump({"clients": self._clients, "grants": self._grants}, f)
        os.replace(tmp, self._state_file)

    # ---- clients -----------------------------------------------------------

    async def get_client(self, client_id: str) -> OAuthClientInformationFull | None:
        data = self._clients.get(client_id)
        return OAuthClientInformationFull.model_validate(data["info"]) if data else None

    async def register_client(self, client_info: OAuthClientInformationFull) -> None:
        if not client_info.client_id:
            raise RegistrationError("invalid_client_metadata", "client_id missing")
        if len(self._clients) >= MAX_CLIENTS:
            # Registration is open to the internet; drop the oldest client nobody logged in with.
            in_use = {g["client_id"] for g in self._grants.values()}
            idle = sorted((c["created"], cid) for cid, c in self._clients.items() if cid not in in_use)
            if not idle:
                raise RegistrationError("invalid_client_metadata", "Too many registered clients")
            del self._clients[idle[0][1]]
        self._clients[client_info.client_id] = {
            "created": time.time(),
            "info": client_info.model_dump(mode="json", exclude_none=True),
        }
        self._persist()

    # ---- authorization (browser) ------------------------------------------

    async def authorize(self, client: OAuthClientInformationFull, params: AuthorizationParams) -> str:
        now = time.time()
        self._pending = {k: p for k, p in self._pending.items() if p["expires"] > now}
        if len(self._pending) > 1000:
            raise AuthorizeError("temporarily_unavailable", "Too many logins in progress")
        txn = secrets.token_urlsafe(32)
        self._pending[txn] = {"client": client, "params": params, "expires": now + LOGIN_TTL}
        return f"{self.public_url}/login?txn={txn}"

    def _locked_out(self) -> bool:
        cutoff = time.time() - FAILURE_WINDOW
        self._failures = [t for t in self._failures if t > cutoff]
        return len(self._failures) >= MAX_FAILURES

    async def login(self, request: Request) -> Response:
        """GET shows the password form; POST checks it and sends the browser back to Claude."""
        if request.method == "GET":
            txn = request.query_params.get("txn", "")
            password = None
        else:
            form = await request.form()
            txn = str(form.get("txn", ""))
            password = str(form.get("password", ""))

        pending = self._pending.get(txn)
        if not pending or pending["expires"] < time.time():
            return _page("This sign-in link has expired. Go back to Claude and connect again.", status=400)
        client: OAuthClientInformationFull = pending["client"]
        params: AuthorizationParams = pending["params"]
        who = html.escape(client.client_name or "An application")
        where = html.escape(params.redirect_uri.host or str(params.redirect_uri))

        def form_page(error: str = "", status: int = 200) -> Response:
            return _page(
                f"<p><b>{who}</b> (returning to <code>{where}</code>) wants to read and edit your "
                f"Excel workbook.</p>{error}"
                f'<form method="post" action="/login"><input type="hidden" name="txn" value="{html.escape(txn)}">'
                '<input type="password" name="password" placeholder="Password" autofocus required>'
                "<button>Allow access</button></form>",
                status=status,
            )

        if password is None:
            return form_page()
        if self._locked_out():
            return form_page('<p class="err">Too many wrong passwords. Try again in 15 minutes.</p>', 429)
        if not hmac.compare_digest(password.encode(), self._password):
            self._failures.append(time.time())
            return form_page('<p class="err">Wrong password.</p>', 401)

        del self._pending[txn]
        code = secrets.token_urlsafe(32)
        self._codes[code] = AuthorizationCode(
            code=code,
            scopes=params.scopes or [],
            expires_at=time.time() + CODE_TTL,
            client_id=client.client_id or "",
            code_challenge=params.code_challenge,
            redirect_uri=params.redirect_uri,
            redirect_uri_provided_explicitly=params.redirect_uri_provided_explicitly,
            resource=self.resource_url,  # this server is the only resource tokens are good for
            subject=SUBJECT,
        )
        return RedirectResponse(
            construct_redirect_uri(str(params.redirect_uri), code=code, state=params.state),
            status_code=302,
            headers={"Cache-Control": "no-store"},
        )

    # ---- tokens ------------------------------------------------------------

    def _issue(self, grant_id: str, client_id: str, scopes: list[str]) -> OAuthToken:
        access, refresh = secrets.token_urlsafe(32), secrets.token_urlsafe(32)
        now = int(time.time())
        self._grants[grant_id] = {
            "client_id": client_id,
            "scopes": scopes,
            "access": _hash(access),
            "access_exp": now + ACCESS_TTL,
            "refresh": _hash(refresh),
            "refresh_exp": now + REFRESH_TTL,
        }
        self._persist()
        return OAuthToken(
            access_token=access,
            expires_in=ACCESS_TTL,
            refresh_token=refresh,
            scope=" ".join(scopes) or None,
        )

    def _find(self, kind: str, token: str) -> tuple[str, dict[str, Any]] | None:
        digest = _hash(token)
        for grant_id, grant in self._grants.items():
            if hmac.compare_digest(grant[kind], digest):
                return (grant_id, grant) if grant[f"{kind}_exp"] > time.time() else None
        return None

    async def load_authorization_code(
        self, client: OAuthClientInformationFull, authorization_code: str
    ) -> AuthorizationCode | None:
        code = self._codes.get(authorization_code)
        return code if code and code.client_id == client.client_id else None

    async def exchange_authorization_code(
        self, client: OAuthClientInformationFull, authorization_code: AuthorizationCode
    ) -> OAuthToken:
        self._codes.pop(authorization_code.code, None)  # single use
        return self._issue(secrets.token_hex(8), client.client_id or "", authorization_code.scopes)

    async def load_refresh_token(self, client: OAuthClientInformationFull, refresh_token: str) -> RefreshToken | None:
        found = self._find("refresh", refresh_token)
        if not found or found[1]["client_id"] != client.client_id:
            return None
        grant = found[1]
        return RefreshToken(
            token=refresh_token,
            client_id=grant["client_id"],
            scopes=grant["scopes"],
            expires_at=grant["refresh_exp"],
            resource=self.resource_url,
            subject=SUBJECT,
        )

    async def exchange_refresh_token(
        self, client: OAuthClientInformationFull, refresh_token: RefreshToken, scopes: list[str]
    ) -> OAuthToken:
        found = self._find("refresh", refresh_token.token)
        grant_id = found[0] if found else secrets.token_hex(8)
        # Re-issuing under the same grant id replaces (rotates) both old tokens.
        return self._issue(grant_id, client.client_id or "", scopes or refresh_token.scopes)

    async def load_access_token(self, token: str) -> AccessToken | None:
        found = self._find("access", token)
        if not found:
            return None
        grant = found[1]
        return AccessToken(
            token=token,
            client_id=grant["client_id"],
            scopes=grant["scopes"],
            expires_at=grant["access_exp"],
            resource=self.resource_url,
            subject=SUBJECT,
        )

    async def revoke_token(self, token: AccessToken | RefreshToken) -> None:
        found = self._find("access", token.token) or self._find("refresh", token.token)
        if found:
            del self._grants[found[0]]
            self._persist()


def _page(body: str, status: int = 200) -> HTMLResponse:
    return HTMLResponse(
        "<!doctype html><html><head><meta charset=utf-8>"
        '<meta name=viewport content="width=device-width,initial-scale=1"><title>Excel MCP sign-in</title>'
        "<style>body{font-family:system-ui,sans-serif;max-width:26rem;margin:12vh auto;padding:0 1rem;"
        "line-height:1.5}input,button{font:inherit;padding:.6rem;width:100%;box-sizing:border-box;"
        "margin-top:.6rem}.err{color:#b00020}code{word-break:break-all}</style></head>"
        f"<body><h2>Excel MCP</h2>{body}</body></html>",
        status_code=status,
        headers={"Cache-Control": "no-store", "X-Frame-Options": "DENY", "Referrer-Policy": "no-referrer"},
    )
