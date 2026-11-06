"""Exchange Web Services client built on top of ``exchangelib``.

Apple Mac Mail's OAuth grant has scopes ``EAS.AccessAsUser.All`` +
``EWS.AccessAsUser.All`` for ``outlook.office.com``, plus ``User.Read``
for Graph. Mail / calendar / contacts are therefore reached via EWS, not
via the Graph mail/calendar endpoints (which would 403 here).

One ``Account`` lives for the life of the process. A new one per call costs a
token request, a version probe and a lookup per folder, and exchangelib keeps
each token's ``Protocol`` in a cache forever. The access token is held in memory
and re-minted shortly before it expires.
"""

from __future__ import annotations

import threading
import time
from typing import Callable, TypeVar

from exchangelib import (
    DELEGATE,
    OAUTH2,
    Account,
    Configuration,
    OAuth2AuthorizationCodeCredentials,
)
from exchangelib.errors import UnauthorizedError
from exchangelib.version import EXCHANGE_O365, Version

from . import auth

T = TypeVar("T")

EWS_SCOPE = "https://outlook.office.com/.default offline_access"
REFRESH_SKEW_SECONDS = 300
# Sockets idle this long may be dead without the OS knowing (laptop sleep, network change).
IDLE_CLOSE_SECONDS = 120
MAX_CONNECTIONS = 4

_lock = threading.Lock()
_account: Account | None = None
_expires_at = 0.0
_last_used = 0.0


def _credentials(token: str) -> OAuth2AuthorizationCodeCredentials:
    # This class, not OAuth2Credentials: exchangelib's mode for "the application
    # refreshes tokens itself". OAuth2Credentials has its token wiped whenever a
    # session closes, and then exchangelib attempts a client-credentials login.
    bundle = auth.load_tokens()
    return OAuth2AuthorizationCodeCredentials(
        client_id=bundle.client_id,
        tenant_id=bundle.tenant,
        client_secret=None,  # public client
        # No expiry here on purpose. If oauthlib knows the expiry, an expired token
        # raises TokenExpiredError, which exchangelib retries in `while True` with
        # the same token because a public client can't refresh. Without it, a dead
        # token gets a 401, which raises UnauthorizedError. Expiry is tracked below.
        access_token={"access_token": token, "token_type": "Bearer"},
    )


def _build_account(creds: OAuth2AuthorizationCodeCredentials) -> Account:
    primary_smtp = auth.get_user_email()
    if not primary_smtp:
        raise RuntimeError(
            "user_email not set in Keychain. Re-run `python -m outlook_mcp.bootstrap` "
            "or set it manually under service 'outlook-mcp', account 'user_email'."
        )
    config = Configuration(
        server="outlook.office365.com",
        credentials=creds,
        auth_type=OAUTH2,
        version=Version(build=EXCHANGE_O365),  # skips the version probe
        max_connections=MAX_CONNECTIONS,
    )
    return Account(
        primary_smtp_address=primary_smtp,
        config=config,
        autodiscover=False,
        access_type=DELEGATE,
    )


def _get_account() -> Account:
    """Return the shared Account, with a token that won't expire mid-call."""
    global _account, _expires_at, _last_used
    with _lock:
        now = time.time()
        if _account is not None and now < _expires_at - REFRESH_SKEW_SECONDS:
            if now - _last_used > IDLE_CLOSE_SECONDS:
                _account.protocol.close()
            _last_used = now
            return _account

        token, expires_in = auth.mint_token_for(EWS_SCOPE)
        creds = _credentials(token)
        if _account is None:
            _account = _build_account(creds)
        else:
            # The setter closes pooled sessions. Cached folders and version are kept.
            _account.protocol.credentials = creds
        _expires_at = now + expires_in
        _last_used = now
        return _account


def _drop_token() -> None:
    global _expires_at
    with _lock:
        _expires_at = 0.0


def with_account(action: Callable[[Account], T], *, retry_auth: bool = True) -> T:
    """Run ``action(account)`` on the shared Account.

    If Exchange rejects the token (revoked, password reset), it is dropped so the
    next call mints a new one. ``retry_auth`` also repeats the action once. Write
    tools pass False, because repeating one could duplicate a send or a create.
    """
    try:
        return action(_get_account())
    except UnauthorizedError:
        _drop_token()
        if not retry_auth:
            raise
    return action(_get_account())
