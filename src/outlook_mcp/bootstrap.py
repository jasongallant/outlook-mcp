"""Interactive CLI to seed Microsoft 365 tokens into macOS Keychain.

Two paths:

1. **Mac Mail.app sniff (recommended)** — run mitmweb in regular HTTP-proxy
   mode on Mac, install + trust mitmproxy CA, set Mac system proxy to 8080,
   add Exchange account in Mail.app, capture the OAuth POST. Apple's Mac Mail
   client_id is consented at the tenant level for both EWS and Graph.

2. **Re-seed with already-captured values** — paste client_id / tenant /
   access_token / refresh_token directly. Use after running the sniff once.

Walkthrough printed by ``--show-walkthrough`` (default on).
"""

from __future__ import annotations

import json
import sys

import click

from . import auth, config

WALKTHROUGH = """\
Microsoft 365 / Outlook MCP token bootstrap (Mac Mail.app sniff)
================================================================

Why this path: many institutional M365 tenants block 3rd-party OAuth consent,
including device code flows for Microsoft's own CLIs. Apple's Mac Mail uses a
first-party client_id that's pre-consented at tenant level for both EWS and
Graph, and its OAuth flow is non-broker (clean access_token + refresh_token in
JSON, no JWE/PRT garbage). Sniff once on Mac, refresh forever.

Setup
-----
  1. brew install mitmproxy
  2. mitmweb --mode regular@8080 --set web_port=8081 --set web_password=<yourpassword>
  3. Trust mitmproxy CA in macOS System keychain:
       sudo security add-trusted-cert -d -r trustRoot \\
         -k /Library/Keychains/System.keychain \\
         ~/.mitmproxy/mitmproxy-ca-cert.pem
  4. Set Mac proxy:
       networksetup -setwebproxy "Wi-Fi" 127.0.0.1 8080
       networksetup -setsecurewebproxy "Wi-Fi" 127.0.0.1 8080
  5. Open Mail.app → Mail → Add Account → Microsoft Exchange
       Email: you@yourschool.edu
       Sign in → complete your institution's SSO + MFA
  6. Find the captured POST in mitmweb at:
       https://login.microsoftonline.com/<tenant>/oauth2/v2.0/token
     Copy from the request:
       - client_id (form field)
       - tenant (the UUID in the URL path)
     Copy from the JSON response:
       - access_token
       - refresh_token
       - expires_in
  7. Cleanup after capture:
       networksetup -setwebproxystate "Wi-Fi" off
       networksetup -setsecurewebproxystate "Wi-Fi" off
       pkill -f mitmweb
       (optional) sudo security delete-certificate -c mitmproxy \\
         /Library/Keychains/System.keychain
"""


VALUES_HELP = """\
Read captured values from a JSON FILE instead of interactive prompts.

Terminal input is capped at ~1024 bytes per line (MAX_CANON), so pasting a
multi-KB access_token or refresh_token into a prompt gets silently truncated.
A file has no such limit. FILE shape:

  {
    "client_id": "f8d98a96-...",
    "tenant": "common",
    "user_email": "you@yourschool.edu",
    "access_token": "eyJ0...",
    "refresh_token": "0.AX...",
    "expires_in": 3600
  }
"""


@click.command()
@click.option(
    "--show-walkthrough/--no-walkthrough",
    default=True,
    help="Print the Mac Mail capture instructions before prompting.",
)
@click.option(
    "--values",
    "values_path",
    type=click.Path(exists=True, dir_okay=False),
    default=None,
    help=VALUES_HELP,
)
def main(show_walkthrough: bool, values_path: str | None) -> None:
    if values_path:
        data = json.loads(open(values_path).read())
        client_id = str(data["client_id"]).strip()
        tenant = str(data.get("tenant") or config.get_tenant()).strip()
        user_email = str(data.get("user_email") or config.get_user_email() or "").strip()
        access_token = str(data["access_token"]).strip()
        refresh_token = str(data["refresh_token"]).strip()
        expires_in = int(data.get("expires_in", 3600))
    else:
        if show_walkthrough:
            click.echo(WALKTHROUGH)
            if not click.confirm("Ready to paste the captured values?", default=True):
                click.echo("Aborted.")
                sys.exit(1)
            click.echo(
                "Tip: if a token paste gets truncated, re-run with "
                "`--values FILE` (JSON). Terminal lines cap at ~1024 bytes.\n"
            )

        client_id = click.prompt("client_id", type=str).strip()
        tenant = click.prompt(
            "tenant (UUID, or 'common')",
            type=str,
            default=config.get_tenant(),
            show_default=True,
        ).strip()
        user_email = click.prompt(
            "user_email (primary SMTP, e.g. you@yourschool.edu)", type=str,
            default=config.get_user_email() or "",
        ).strip()
        access_token = click.prompt("access_token", type=str, hide_input=True).strip()
        refresh_token = click.prompt("refresh_token", type=str, hide_input=True).strip()
        expires_in = click.prompt("expires_in (seconds)", type=int, default=3600)

    auth.save_tokens(
        client_id=client_id,
        tenant=tenant,
        access_token=access_token,
        refresh_token=refresh_token,
        expires_in=expires_in,
    )
    auth.set_user_email(user_email)
    config.save({"tenant": tenant, "user_email": user_email})

    click.echo("")
    click.echo(f"Saved tokens to macOS Keychain under service '{auth.SERVICE}'.")
    click.echo("Verifying refresh works...")
    try:
        auth.force_refresh()
    except auth.AuthError as exc:
        click.echo(f"  Refresh test FAILED: {exc}", err=True)
        click.echo(
            "  Tokens were saved, but refresh did not work. Double-check client_id, "
            "tenant, refresh_token capture and re-run.",
            err=True,
        )
        sys.exit(2)
    click.echo("  Refresh OK. You're set.")


if __name__ == "__main__":
    main()
