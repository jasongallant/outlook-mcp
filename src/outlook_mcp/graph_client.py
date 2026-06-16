"""Microsoft Graph HTTP client with auto-refresh on 401."""

from __future__ import annotations

from typing import Any, Callable, TypeVar

import httpx

from . import auth

T = TypeVar("T")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _client(token: str) -> httpx.Client:
    return httpx.Client(
        base_url=GRAPH_BASE,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        },
        timeout=30.0,
        trust_env=False,  # ignore Mac/system proxies during normal Graph calls
    )


def with_retry(action: Callable[[httpx.Client], T]) -> T:
    """Run ``action`` with a Graph HTTP client; refresh once on 401 and retry."""
    token = auth.get_access_token()
    try:
        with _client(token) as c:
            return action(c)
    except httpx.HTTPStatusError as e:
        if e.response.status_code != 401:
            raise
    # 401 → refresh + retry once
    token = auth.force_refresh()
    with _client(token) as c:
        return action(c)


def get(path: str, params: dict[str, Any] | None = None) -> Any:
    """Convenience GET → JSON, raises on non-2xx."""

    def _do(c: httpx.Client) -> Any:
        r = c.get(path, params=params or {})
        r.raise_for_status()
        return r.json()

    return with_retry(_do)


def get_paged(path: str, params: dict[str, Any] | None = None, limit: int = 100) -> list[dict]:
    """Follow @odata.nextLink up to ``limit`` items."""

    def _do(c: httpx.Client) -> list[dict]:
        out: list[dict] = []
        url: str | None = path
        first = True
        while url and len(out) < limit:
            r = c.get(url, params=params if first else None)
            r.raise_for_status()
            data = r.json()
            out.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            first = False
        return out[:limit]

    return with_retry(_do)


def post(path: str, json: dict[str, Any] | None = None) -> Any:
    def _do(c: httpx.Client) -> Any:
        r = c.post(path, json=json or {})
        r.raise_for_status()
        return r.json() if r.content else {}

    return with_retry(_do)


def patch(path: str, json: dict[str, Any]) -> Any:
    def _do(c: httpx.Client) -> Any:
        r = c.patch(path, json=json)
        r.raise_for_status()
        return r.json() if r.content else {}

    return with_retry(_do)


def delete(path: str) -> None:
    def _do(c: httpx.Client) -> None:
        r = c.delete(path)
        r.raise_for_status()

    with_retry(_do)
