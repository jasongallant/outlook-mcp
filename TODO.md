# TODO

- `list_messages(folder="archive")` silently returns the Inbox: `Account` has no `archive` attribute. Decide between `archive_inbox` and a real Archive folder lookup.
- `online_join_url` on events and `web_link` on messages are always None: neither field exists in exchangelib.
- `search_messages` covers the Inbox only. Add a folder option or a mailbox-wide search.
- README "Layout" section is out of date and lists test files that don't exist.
