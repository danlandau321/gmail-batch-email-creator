# Gmail Batch Email Creator

Turns a **Google Doc** (the message) plus a **Google Sheet** (the recipients) into
**Gmail drafts** — one per row, personalized. It never sends anything; you review
and send from Gmail.

```bash
cd ~/gmail-batch-email-creator
.venv/bin/python batch_emailer.py --doc <doc-url> --sheet <sheet-url> --dry-run
```

**It cannot send.** The OAuth scope is `gmail.compose`, which permits creating
drafts and nothing else. That is a property of the token, not a policy in the
code, so no flag or bug can turn this into a sender.

There is also a natural-language wrapper: in Claude Code, say *"send the email in
&lt;doc&gt; to the people in &lt;sheet&gt;"* and the `batch-email` skill dry-runs it,
shows you the merge preview, and asks before creating anything.

---

## Setup

### 1. Google Cloud project

The script talks to three APIs, so all three must be enabled in whatever Cloud
project issues your credentials:

- Gmail API
- Google Docs API
- Google Sheets API

At <https://console.cloud.google.com> → pick (or create) a project → **APIs &
Services → Library** → search each one → **Enable**.

### 2. OAuth consent screen

**APIs & Services → OAuth consent screen.** Internal user type if the project
lives in a Google Workspace org you control; otherwise External, and add yourself under
**Test users**. Nothing else on that screen matters for a personal script.

### 3. Desktop OAuth client

**APIs & Services → Credentials → Create credentials → OAuth client ID →
Application type: Desktop app.** Download the JSON and save it as:

```
~/gmail-batch-email-creator/credentials.json
```

It **must** be a *Desktop app* client. A Web-application client fails with
`Error 400: redirect_uri_mismatch`, because a web client only accepts the exact
redirect URI registered on it, while this script listens on a random localhost
port. Renaming the JSON's `"web"` key to `"installed"` does not change how
Google treats the client.

If you must reuse a web client, register `http://localhost:8080/` as an
authorized redirect URI on it and run with `--port 8080`.

### 4. First run

```bash
cd ~/gmail-batch-email-creator
.venv/bin/python batch_emailer.py --doc <doc-url> --sheet <sheet-url> --dry-run
```

A browser opens; sign in as the account whose Gmail should hold the drafts and
approve the three permissions:

| Permission asked | What it's for |
|---|---|
| Manage drafts and send email | Creating the drafts. The script only ever calls `drafts.create` — never `send`. |
| See your Google Docs documents | Reading the message Doc. |
| See, edit, create and delete your spreadsheets | Reading the recipient Sheet, and writing the `status` column back. |

That writes `token.json` next to the script and it stops asking. Both
`credentials.json` and `token.json` are gitignored — treat them like passwords.

To switch to a different Gmail account, delete `token.json` and run again. To
revoke access entirely: <https://myaccount.google.com/permissions>.

---

## The Doc

The first line can set the subject; everything after it is the body.
`[field]` works as well as `{{field}}` — but only when it names a real column,
so prose like `[Luma link]` or `[TBD]` is left alone.

If the Doc holds several drafts under headings like `Email 1` / `Email 2`, pick
one with `--section "Email 1"`. The section runs to the next heading of the same
shape, or to a divider line.

```
Subject: Quick intro — {{first_name}} <> AI Fund

Hi {{first_name}},

Saw {{fund_name}} led the round in ...

Best,
Dan
```

Bold, italics, links and bullet lists are preserved — the draft goes out as HTML
with a plain-text alternative. Tables and inline images are skipped (you get a
warning). Share the Doc with the account you authorized, or just own it.

## The Sheet

Works with a native Google Sheet **or** an .xlsx uploaded to Drive. The .xlsx
case is read-only — status can't be written back, so `drafts_log.csv` is the
record of what was drafted.

The header row is found automatically (the first row naming an email column), so
title and legend rows above it are fine; `--header-row N` overrides. Tab names
match loosely — `--tab "CEOs & cofounders"` finds "CEOs & Co-founders".

Narrow the list with `--filter COL=VAL`, e.g. `--filter tier=1`, repeatable, and
`VAL` may be a comma-separated list.

**The header names are the merge fields.** Matching
ignores case, spaces and punctuation, so a column `Fund Name` fills
`{{fund_name}}` or `{{Fund Name}}`.

| Column | Required | Meaning |
|---|---|---|
| `email` | ✅ | Recipient. `email address`, `recipient` and `to` also work. |
| `cc` / `bcc` | | Extra addresses, comma or semicolon separated |
| `subject` | | Per-row subject, overrides the Doc's |
| `attachments` | | Local file path(s), semicolon separated |
| `status` | | Written back as `Draft created <date>`; those rows are skipped next run |
| `draft_id` | | Written back with the Gmail draft id |
| anything else | | Available to the Doc as `{{that_column}}` |

If there's a `name` column but no `first_name`, the first word is used as
`{{first_name}}` and the last as `{{last_name}}`.

Add `status` and `draft_id` columns if you want the sheet to track itself —
a re-run after a partial failure then picks up where it stopped instead of
double-drafting.

## Flags

| Flag | Effect |
|---|---|
| `--dry-run` | Print what would be drafted; create nothing. Always start here. |
| `--tab NAME` | Which sheet tab (default: the first one) |
| `--section NAME` | Use one section of a multi-draft Doc, e.g. `"Email 1"` |
| `--filter COL=VAL` | Only rows matching, e.g. `tier=1`; repeatable |
| `--exclude COL=VAL` | Drop rows matching, e.g. `company="Jivi Health"` |
| `--map PH=COL` | Fill a placeholder from another column, e.g. `name=first_name` |
| `--header-row N` | Point at the header row if auto-detection misses |
| `--limit N` | Only the first N eligible rows |
| `--cc` / `--bcc` | Applied to every draft; repeatable |
| `--attach PATH` | Attach a file to every draft; repeatable |
| `--subject "..."` | Override the Doc's subject line |
| `--force` | Re-draft rows that already have a status |
| `--no-mark` | Don't write anything back to the sheet |
| `--allow-missing` | Draft even when a `{{placeholder}}` has no value |

By default a row whose placeholder has no value is **skipped**, so nobody gets
"Hi ,". Use `--allow-missing` only when you mean it.

Every created draft is appended to `drafts_log.csv` (date, row, email, subject,
draft id).

## Typical run

```bash
cd ~/gmail-batch-email-creator

# 1. See what you'd get, without touching Gmail
.venv/bin/python batch_emailer.py --doc <doc> --sheet <sheet> --dry-run

# 2. Two real drafts, eyeball them in Gmail
.venv/bin/python batch_emailer.py --doc <doc> --sheet <sheet> --limit 2

# 3. The rest (already-drafted rows are skipped via the status column)
.venv/bin/python batch_emailer.py --doc <doc> --sheet <sheet> --cc someone@example.com
```

## Troubleshooting

| Message | Fix |
|---|---|
| `Google Docs API has not been used in project ... before` | Enable the API (step 1); the error text links straight to the right page. Wait ~1 min after enabling. |
| `403 The caller does not have permission` | The authorized account can't open that Doc/Sheet — share it. |
| `invalid_grant` / re-auth loop | Delete `token.json` and run again. |
| `Error 400: redirect_uri_mismatch` | `credentials.json` is a Web-application client, not a Desktop one. Redo step 3. |
| `Access blocked: has not completed verification` | Add yourself as a Test user on the OAuth consent screen. |
| `❌ No subject` | Add a `Subject:` first line to the Doc, a `subject` column, or pass `--subject`. |
