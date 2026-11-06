# outlook-mcp

Local MCP server that gives Claude Code (and any other MCP client) access
to your Outlook + Calendar + Contacts + Microsoft To Do via Microsoft Graph
and Exchange Web Services (EWS), authenticated through a one-time
**Mac Mail.app OAuth sniff**.
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
  ews_client.py     exchangelib wrapper for EWS (one shared Account, cached token)
  tools.py          MCP tools (read default; write gated by settings)
  attachments.py    Safe local saves for downloads, path checks for sends
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
  "writes_enabled": false,
  "mail_writes_enabled": false,
  "calendar_writes_enabled": false,
  "todo_writes_enabled": false,
  "attachment_dir": "~/.cache/outlook-mcp/attachments"
}
```

| Key | Default | Description |
|---|---|---|
| `tenant` | `"common"` | Azure AD tenant UUID or `"common"` |
| `user_email` | `""` | Primary SMTP address for EWS |
| `timezone` | `"America/New_York"` | IANA timezone for calendar operations |
| `writes_enabled` | `false` | **Master** — enable all write tools (or `OUTLOOK_MCP_WRITES=1`) |
| `mail_writes_enabled` | `false` | Enable only email writes (or `OUTLOOK_MCP_MAIL_WRITES=1`) |
| `calendar_writes_enabled` | `false` | Enable only calendar writes (or `OUTLOOK_MCP_CALENDAR_WRITES=1`) |
| `todo_writes_enabled` | `false` | Enable only To Do writes (or `OUTLOOK_MCP_TODO_WRITES=1`) |
| `attachment_dir` | `"~/.cache/outlook-mcp/attachments"` | Where `get_attachment` saves files. Private to your user, purged after 7 days |

Write tools come in three independently-gated groups — **mail** (email
send/reply/delete), **calendar** (events), and **To Do** (tasks). A group is
enabled when the master flag **or** its own flag is set, so you can, for
example, allow To Do writes while keeping email sending off.

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

To enable writes, set the relevant flag in settings.json, or pass the env var
in the MCP client config. Use the master `writes_enabled` /
`OUTLOOK_MCP_WRITES` for everything, or a per-group flag to enable one group.
For example, To Do writes only (email/calendar stay read-only):

```json
"outlook": {
  "command": "/path/to/outlook-mcp/.venv/bin/python",
  "args": ["-m", "outlook_mcp.server"],
  "env": { "OUTLOOK_MCP_TODO_WRITES": "1" }
}
```

## Tools exposed

**Read (always on):**
- `get_user_profile` — Graph `/me`
- `list_mail_folders`
- `list_messages` — folder, unread filter, limit
- `get_message` — full body + attachment metadata
- `search_messages` — Exchange search index: whole words and word prefixes, `from:` / `subject:` supported, not mid-word substrings. Just-delivered mail may take a moment to appear
- `list_attachments` — metadata only, including `attachment_id`
- `get_attachment` — saves one attachment to `attachment_dir` and returns its path. Attached emails are saved as `.eml`
- `list_events` — date range
- `get_event`
- `list_contacts`
- `list_todo_lists` — Microsoft To Do lists + open/total counts
- `list_tasks` — tasks in a list (`list` name, `include_completed`, `limit`)
- `get_task` — full task incl. notes/body
- `search_tasks` — subject + body substring within a list

**Write (OFF by default):**
- `send_email` — `attachments` (local file paths)
- `reply_to_message` — `reply_all` flag, `attachments` (local file paths)
- `mark_message_read`
- `delete_message` — moves to Deleted Items (recoverable)
- `create_event`
- `update_event`
- `respond_to_event` — accept / decline / tentativelyAccept
- `create_task` — `list`, `due` (YYYY-MM-DD), `reminder` (ISO), `body`, `importance`
- `update_task` — patch title / due / reminder / body / importance / status
- `complete_task` — mark done (reversible via `update_task status=notStarted`)
- `delete_task` — moves to Deleted Items (recoverable)
- `create_todo_list`

Each write tool description starts with `DESTRUCTIVE:` so the MCP client
surfaces a clear approval prompt before running.

### Attachment safety

A malicious email could tell an agent to attach a local file and send it out.
The send tools refuse hidden paths (`.ssh`, `.aws`, `.config`) and `~/Library`,
except OneDrive and iCloud Drive folders, but this does not cover every sensitive
file. **Keep client approval on for `send_email` and `reply_to_message`.** That
prompt is the real protection.

### Microsoft To Do

To Do lists are ordinary `IPF.Task` mailbox folders (the default **Tasks**
list plus one child folder per custom list), so they are reached over EWS like
mail and calendar — **not** via Graph. The Mac Mail OAuth grant carries no
`Tasks.ReadWrite` scope, so the Graph `/me/todo/*` endpoints return 403.

Due dates are whole dates (no time); reminders are datetimes in the
`timezone` from settings.json. Due dates are stored at noon UTC so they land
on the intended calendar day in any timezone (EWS otherwise rolls a UTC-midnight
date back a day in negative-offset mailboxes).

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
