"""Microsoft identity token storage and refresh.

Tokens live in macOS Keychain under service ``outlook-mcp``. The bootstrap
flow seeds them once after a Mac Mail OAuth capture; this module owns the
runtime load / refresh / persist cycle and exposes ``get_access_token`` to
the rest of the package.

Default scope is ``https://graph.microsoft.com/.default offline_access`` —
Microsoft Graph access. Apple's Mac Mail client_id is pre-consented for both
EWS and Graph at most institutional tenants, so a single refresh_token can
mint either.
"""

from __future__ import annotations

import time
from dataclasses import dataclass

import httpx
import keyring

from . import config as _cfg

SERVICE = "outlook-mcp"
DEFAULT_SCOPE = "https://graph.microsoft.com/.default offline_access"
EXPIRY_SKEW_SECONDS = 60


class AuthError(RuntimeError):
    """Base class for auth failures."""


class AuthExpiredError(AuthError):
    """Refresh token rejected. User must re-run bootstrap."""


class AuthMissingError(AuthError):
    """No tokens in Keychain. User has not bootstrapped yet."""


@dataclass
class TokenBundle:
    client_id: str
    tenant: str
    access_token: str
    refresh_token: str
    expires_at: int


def _get(account: str) -> str | None:
    return keyring.get_password(SERVICE, account)


def _set(account: str, value: str) -> None:
    keyring.set_password(SERVICE, account, value)


def load_tokens() -> TokenBundle:
    client_id = _get("client_id")
    tenant = _get("tenant") or _cfg.get_tenant()
    access_token = _get("access_token")
    refresh_token = _get("refresh_token")
    expires_at_raw = _get("access_token_expires_at")

    missing = [
        name
        for name, value in [
            ("client_id", client_id),
            ("access_token", access_token),
            ("refresh_token", refresh_token),
            ("access_token_expires_at", expires_at_raw),
        ]
        if not value
    ]
    if missing:
        raise AuthMissingError(
            f"Missing Keychain entries: {', '.join(missing)}. "
            "Run `python -m outlook_mcp.bootstrap` first."
        )

    return TokenBundle(
        client_id=client_id,  # type: ignore[arg-type]
        tenant=tenant,
        access_token=access_token,  # type: ignore[arg-type]
        refresh_token=refresh_token,  # type: ignore[arg-type]
        expires_at=int(expires_at_raw),  # type: ignore[arg-type]
    )


def save_tokens(
    *,
    client_id: str,
    tenant: str,
    access_token: str,
    refresh_token: str,
    expires_in: int,
) -> None:
    _set("client_id", client_id)
    _set("tenant", tenant)
    _set("access_token", access_token)
    _set("refresh_token", refresh_token)
    _set("access_token_expires_at", str(int(time.time()) + int(expires_in)))


def _refresh(bundle: TokenBundle, scope: str = DEFAULT_SCOPE) -> TokenBundle:
    response = httpx.post(
        f"https://login.microsoftonline.com/{bundle.tenant}/oauth2/v2.0/token",
        data={
            "client_id": bundle.client_id,
            "grant_type": "refresh_token",
            "refresh_token": bundle.refresh_token,
            "scope": scope,
        },
        timeout=20.0,
        # Don't pick up Mac/system proxies: refresh must talk directly to MS.
        trust_env=False,
    )
    if response.status_code >= 400:
        raise AuthExpiredError(
            f"Refresh failed ({response.status_code}): {response.text[:300]}. "
            "Re-run `python -m outlook_mcp.bootstrap`."
        )
    payload = response.json()
    new_access = payload["access_token"]
    new_refresh = payload.get("refresh_token", bundle.refresh_token)
    expires_in = int(payload.get("expires_in", 3600))
    new_expires_at = int(time.time()) + expires_in
    _set("access_token", new_access)
    _set("refresh_token", new_refresh)
    _set("access_token_expires_at", str(new_expires_at))
    return TokenBundle(
        client_id=bundle.client_id,
        tenant=bundle.tenant,
        access_token=new_access,
        refresh_token=new_refresh,
        expires_at=new_expires_at,
    )


def get_access_token() -> str:
    """Return a valid Graph access token (cached), refreshing if needed."""
    bundle = load_tokens()
    if time.time() >= bundle.expires_at - EXPIRY_SKEW_SECONDS:
        bundle = _refresh(bundle)
    return bundle.access_token


def force_refresh() -> str:
    """Refresh Graph token unconditionally. Used after a 401."""
    return _refresh(load_tokens()).access_token


def mint_token_for(scope: str) -> tuple[str, int]:
    """Mint a fresh access token for an arbitrary scope using the cached refresh token.

    Does not update the cached Graph access_token in Keychain — returns
    ``(access_token, expires_in_seconds)`` for the caller to hold. Useful for
    resource-specific scopes like
    ``https://outlook.office.com/.default offline_access`` (EWS).

    The refresh_token IS rotated in Keychain if the response carries a new one,
    so the next call uses the latest refresh credential.
    """
    bundle = load_tokens()
    response = httpx.post(
        f"https://login.microsoftonline.com/{bundle.tenant}/oauth2/v2.0/token",
        data={
            "client_id": bundle.client_id,
            "grant_type": "refresh_token",
            "refresh_token": bundle.refresh_token,
            "scope": scope,
        },
        timeout=20.0,
        trust_env=False,
    )
    if response.status_code >= 400:
        raise AuthExpiredError(
            f"Refresh for scope {scope!r} failed ({response.status_code}): "
            f"{response.text[:300]}. Re-run `python -m outlook_mcp.bootstrap`."
        )
    payload = response.json()
    new_refresh = payload.get("refresh_token")
    if new_refresh and new_refresh != bundle.refresh_token:
        _set("refresh_token", new_refresh)
    return payload["access_token"], int(payload.get("expires_in", 3600))


def get_user_email() -> str:
    """Return the primary SMTP address used to scope the Mac Mail OAuth grant."""
    return _get("user_email") or _cfg.get_user_email()


def set_user_email(email: str) -> None:
    _set("user_email", email)


def get_tenant() -> str:
    return _get("tenant") or _cfg.get_tenant()
