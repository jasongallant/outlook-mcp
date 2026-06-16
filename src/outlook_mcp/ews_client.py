"""Exchange Web Services client built on top of ``exchangelib``.

Apple Mac Mail's OAuth grant has scopes ``EAS.AccessAsUser.All`` +
``EWS.AccessAsUser.All`` for ``outlook.office.com``, plus ``User.Read``
for Graph. Mail / calendar / contacts are therefore reached via EWS, not
via the Graph mail/calendar endpoints (which would 403 here).

Each call mints a fresh resource-scoped access token from the cached
refresh token, then constructs an ``exchangelib.Account``. We keep things
stateless — Account construction is fast enough for interactive MCP calls.
"""

from __future__ import annotations

from typing import Any, Callable, TypeVar

from exchangelib import (
    DELEGATE,
    OAUTH2,
    Account,
    Configuration,
    OAuth2Credentials,
)

from . import auth

T = TypeVar("T")

EWS_SCOPE = "https://outlook.office.com/.default offline_access"


def _build_account() -> Account:
    token = auth.mint_token_for(EWS_SCOPE)
    bundle = auth.load_tokens()
    primary_smtp = auth.get_user_email()
    if not primary_smtp:
        raise RuntimeError(
            "user_email not set in Keychain. Re-run `python -m outlook_mcp.bootstrap` "
            "or set it manually under service 'outlook-mcp', account 'user_email'."
        )
    creds = OAuth2Credentials(
        client_id=bundle.client_id,
        tenant_id=bundle.tenant,
        client_secret=None,  # public client
        access_token={
            "access_token": token,
            "token_type": "Bearer",
            "expires_in": 3600,
        },
    )
    config = Configuration(
        server="outlook.office365.com",
        credentials=creds,
        auth_type=OAUTH2,
    )
    return Account(
        primary_smtp_address=primary_smtp,
        config=config,
        autodiscover=False,
        access_type=DELEGATE,
    )


def with_account(action: Callable[[Account], T]) -> T:
    """Run ``action(account)`` with a freshly built EWS Account."""
    account = _build_account()
    return action(account)
