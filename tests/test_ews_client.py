"""Long-lived Account, token cache, and auth-retry behaviour. No network."""

from __future__ import annotations

import threading
import time
from types import SimpleNamespace

import pytest
from exchangelib import OAUTH2, Configuration
from exchangelib.errors import UnauthorizedError
from exchangelib.protocol import Protocol
from exchangelib.version import EXCHANGE_O365, Version

from outlook_mcp import ews_client as ec


class FakeProtocol:
    def __init__(self, credentials):
        self.credentials = credentials
        self.closed = 0

    def close(self):
        self.closed += 1


class FakeAccount:
    built = 0

    def __init__(self, primary_smtp_address, config, autodiscover, access_type):
        type(self).built += 1
        self.config = config
        self.protocol = FakeProtocol(config.credentials)


@pytest.fixture
def mint(monkeypatch):
    counter = SimpleNamespace(calls=0)

    def fake_mint(scope):
        counter.calls += 1
        return f"tok-{counter.calls}", 3600

    FakeAccount.built = 0
    monkeypatch.setattr(ec.auth, "mint_token_for", fake_mint)
    monkeypatch.setattr(
        ec.auth, "load_tokens", lambda: SimpleNamespace(client_id="cid", tenant="tid")
    )
    monkeypatch.setattr(ec.auth, "get_user_email", lambda: "me@example.edu")
    monkeypatch.setattr(ec, "Account", FakeAccount)
    monkeypatch.setattr(ec, "_account", None)
    monkeypatch.setattr(ec, "_expires_at", 0.0)
    monkeypatch.setattr(ec, "_last_used", 0.0)
    return counter


def _token(account):
    return account.protocol.credentials.access_token["access_token"]


def test_one_mint_and_one_account_across_calls(mint):
    accounts = [ec.with_account(lambda a: a) for _ in range(5)]
    assert mint.calls == 1
    assert FakeAccount.built == 1
    assert all(a is accounts[0] for a in accounts)


def test_concurrent_first_calls_share_one_mint(mint):
    barrier = threading.Barrier(8)
    seen = []

    def worker():
        barrier.wait()
        seen.append(ec.with_account(lambda a: a))

    threads = [threading.Thread(target=worker) for _ in range(8)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert mint.calls == 1
    assert FakeAccount.built == 1
    assert len({id(a) for a in seen}) == 1


def test_remint_inside_skew_keeps_account_and_swaps_token(mint):
    first = ec.with_account(lambda a: a)
    assert _token(first) == "tok-1"
    ec._expires_at = time.time() + ec.REFRESH_SKEW_SECONDS - 1
    second = ec.with_account(lambda a: a)
    assert second is first
    assert mint.calls == 2
    assert _token(second) == "tok-2"


def test_no_remint_outside_skew(mint):
    ec.with_account(lambda a: a)
    ec._expires_at = time.time() + ec.REFRESH_SKEW_SECONDS + 60
    ec.with_account(lambda a: a)
    assert mint.calls == 1


def test_token_given_to_exchangelib_carries_no_expiry(mint):
    # If oauthlib knows the expiry, an expired token raises TokenExpiredError, and
    # exchangelib's services/common.py retries that in `while True` with the same token.
    token = ec.with_account(lambda a: a).protocol.credentials.access_token
    assert set(token) == {"access_token", "token_type"}


def _real_protocol(credentials):
    config = Configuration(
        server="outlook.office365.com",
        credentials=credentials,
        auth_type=OAUTH2,
        version=Version(build=EXCHANGE_O365),
    )
    return Protocol(config=config)


def test_token_survives_session_close_on_real_protocol(mint):
    # exchangelib wipes the token of client-credentials-style credentials when a
    # session closes, and the next session then attempts a client-credentials login.
    creds = ec._credentials("tok")
    protocol = _real_protocol(creds)
    protocol.close_session(protocol.create_session())
    assert creds.access_token["access_token"] == "tok"
    assert protocol.create_session().token["access_token"] == "tok"


def test_token_swap_reaches_new_sessions_on_real_protocol(mint):
    protocol = _real_protocol(ec._credentials("old"))
    protocol.release_session(protocol.create_session())
    new = ec._credentials("new")
    protocol.credentials = new
    assert new.access_token["access_token"] == "new"
    assert protocol.create_session().token["access_token"] == "new"


def test_config_pins_version_and_allows_parallel_requests(mint):
    config = ec.with_account(lambda a: a).config
    assert config.version.build == ec.EXCHANGE_O365
    assert config.max_connections == ec.MAX_CONNECTIONS


def test_read_retries_once_with_new_token_on_unauthorized(mint):
    tokens = []

    def action(a):
        tokens.append(_token(a))
        if len(tokens) == 1:
            raise UnauthorizedError("Invalid credentials")
        return "ok"

    assert ec.with_account(action) == "ok"
    assert tokens == ["tok-1", "tok-2"]


def test_read_gives_up_after_second_unauthorized(mint):
    def action(a):
        raise UnauthorizedError("Invalid credentials")

    with pytest.raises(UnauthorizedError):
        ec.with_account(action)
    assert mint.calls == 2


def test_write_never_retries_but_still_drops_token(mint):
    runs = []

    def action(a):
        runs.append(_token(a))
        raise UnauthorizedError("Invalid credentials")

    with pytest.raises(UnauthorizedError):
        ec.with_account(action, retry_auth=False)
    assert runs == ["tok-1"]
    assert ec.with_account(lambda a: _token(a)) == "tok-2"


def test_other_errors_keep_token_and_do_not_retry(mint):
    runs = []

    def action(a):
        runs.append(1)
        raise ValueError("nope")

    with pytest.raises(ValueError):
        ec.with_account(action)
    assert len(runs) == 1
    assert mint.calls == 1


def test_idle_closes_pooled_sessions_without_reminting(mint):
    account = ec.with_account(lambda a: a)
    ec.with_account(lambda a: a)
    assert account.protocol.closed == 0
    ec._last_used = time.time() - ec.IDLE_CLOSE_SECONDS - 1
    ec.with_account(lambda a: a)
    assert account.protocol.closed == 1
    assert mint.calls == 1
