# outlook-mcp

Local MCP server that gives Claude Code (and any other MCP client) read access
to your Outlook + Calendar + Contacts via Microsoft Graph and Exchange Web
Services (EWS), authenticated through a one-time **Mac Mail.app OAuth sniff**.
Token auto-refreshes forever — no admin consent dance, no device-code flow
(many institutional M365 tenants block both).

> Authorization scope: only ever touches *your own* M365 account, using
> *your own* Mac Mail OAuth tokens, on *your own* devices. Apple's Mac Mail
> uses a first-party `client_id` (`f8d98a96-0999-43f5-8af3-69971c7bb423`)
> that's pre-consented at the tenant level for both EWS and Graph.

## Architecture

```
src/outlook_mcp/
  config.py         Settings from ~/.config/outlook-mcp/settings.json
  auth.py           OAuth token storage in macOS Keychain + MS refresh
  graph_client.py   Microsoft Graph httpx wrapper + retry-once-on-401
  ews_client.py     exchangelib wrapper for EWS (mail, calendar, contacts)
  tools.py          MCP tools (read default; write gated by settings)
  server.py         FastMCP entrypoint, stdio transport
  bootstrap.py      Interactive CLI to seed Keychain after Mac Mail sniff
```

## Configuration

Before bootstrapping, create `~/.config/outlook-mcp/settings.json`:

```bash
mkdir -p ~/.config/outlook-mcp
cp settings.json.example ~/.config/outlook-mcp/settings.json
```

Edit the file:

```json
{
  "tenant": "common",
  "user_email": "you@yourschool.edu",
  "timezone": "America/New_York",
  "writes_enabled": false
}
```

| Key | Default | Description |
|---|---|---|
| `tenant` | `"common"` | Azure AD tenant UUID or `"common"` |
| `user_email` | `""` | Primary SMTP address for EWS |
| `timezone` | `"America/New_York"` | IANA timezone for calendar operations |
| `writes_enabled` | `false` | Enable write tools (or use `OUTLOOK_MCP_WRITES=1` env var) |

## Install

```bash
cd /path/to/outlook-mcp
python3.11 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
```

## One-time bootstrap (Mac Mail sniff)

Detailed walkthrough printed by:

```bash
python -m outlook_mcp.bootstrap
```

Summary of the capture sequence:

1. `brew install mitmproxy`
2. `mitmweb --mode regular@8080 --set web_port=8081 --set web_password=<anything>`
3. Trust mitmproxy CA in macOS System keychain:
   ```bash
   sudo security add-trusted-cert -d -r trustRoot \
     -k /Library/Keychains/System.keychain \
     ~/.mitmproxy/mitmproxy-ca-cert.pem
   ```
4. Set Mac proxy:
   ```bash
   networksetup -setwebproxy "Wi-Fi" 127.0.0.1 8080
   networksetup -setsecurewebproxy "Wi-Fi" 127.0.0.1 8080
   ```
5. Mail.app → Mail → Add Account → Microsoft Exchange → sign in → SSO + MFA.
6. In mitmweb find the `POST login.microsoftonline.com/<tenant>/oauth2/v2.0/token`,
   copy `client_id`, `tenant` (UUID in the URL), `access_token`,
   `refresh_token`, `expires_in`.
7. Run `python -m outlook_mcp.bootstrap`, paste captured values.
8. Cleanup:
   ```bash
   networksetup -setwebproxystate "Wi-Fi" off
   networksetup -setsecurewebproxystate "Wi-Fi" off
   pkill -f mitmweb
   # optionally untrust the CA:
   sudo security delete-certificate -c mitmproxy /Library/Keychains/System.keychain
   ```

The bootstrap script verifies refresh works before exiting and saves
`tenant` and `user_email` to `~/.config/outlook-mcp/settings.json`.

## Run

```bash
python -m outlook_mcp.server
# or via console script:
outlook-mcp
```

## Wire it into Claude Code

Add to `~/.claude/settings.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "/path/to/outlook-mcp/.venv/bin/python",
      "args": ["-m", "outlook_mcp.server"]
    }
  }
}
```

To enable writes, set `writes_enabled: true` in settings.json, or pass the
env var in the MCP client config:

```json
"outlook": {
  "command": "/path/to/outlook-mcp/.venv/bin/python",
  "args": ["-m", "outlook_mcp.server"],
  "env": { "OUTLOOK_MCP_WRITES": "1" }
}
```

## Tools exposed

**Read (always on):**
- `get_user_profile` — Graph `/me`
- `list_mail_folders`
- `list_messages` — folder, unread filter, limit
- `get_message` — full body
- `search_messages` — subject + body substring
- `list_attachments` — metadata only
- `list_events` — date range
- `get_event`
- `list_contacts`

**Write (OFF by default):**
- `send_email`
- `reply_to_message` (`reply_all` flag)
- `mark_message_read`
- `delete_message` — moves to Deleted Items (recoverable)
- `create_event`
- `update_event`
- `respond_to_event` — accept / decline / tentativelyAccept

Each write tool description starts with `DESTRUCTIVE:` so the MCP client
surfaces a clear approval prompt before running.

## Token rotation / re-bootstrap

If the server starts failing with `AuthExpiredError`, the refresh token has
been invalidated (tenant policy change, password reset, MFA reset, etc.).
Re-run the Mac Mail sniff and `python -m outlook_mcp.bootstrap` again.

## Layout

```
src/outlook_mcp/
  config.py          ~55 lines
  auth.py            ~200 lines
  graph_client.py    ~100 lines
  ews_client.py      ~65 lines
  tools.py           ~390 lines — 17 @mcp.tool functions
  server.py          ~25 lines
  bootstrap.py       ~120 lines
tests/
  test_auth.py
  test_tools.py
```
