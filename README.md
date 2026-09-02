# Gmail Batch Email Creator

## What it is

A command-line tool that takes a **Google Doc** (the message) and a **Google Sheet**
(the recipients) and creates one personalized **Gmail draft** per row. Placeholders
like `{{first_name}}` fill from the Sheet's columns.

It never sends. The OAuth scope is `gmail.compose`, which only permits creating
drafts — that's a property of the token itself, so no flag or bug can turn this into
a sender. I review everything in Gmail and send by hand.

```bash
cd ~/gmail-batch-email-creator
.venv/bin/python batch_emailer.py --doc <doc-url> --sheet <sheet-url> --dry-run
```

## Why I built it

I kept needing to send the same email to a list of people with a few details changed
per person, and doing it by hand doesn't scale past a handful. Mail-merge services
want to own your list and send on your behalf; I wanted the opposite — drafts sitting
in my own Gmail that I read before anything goes out. Writing the copy in a Doc and
keeping the list in a Sheet means I can edit both in the places I already work.

## What it's built with

Python 3.10+, no framework. Just the Google API client libraries — Gmail, Docs,
Sheets, and Drive — talking to a desktop OAuth client whose token lives on my
machine. The whole thing is one self-contained script.

---

## Setup

### 1. Google Cloud project

The script talks to four APIs, so all four must be enabled in whatever Cloud
project issues your credentials:

- Gmail API
- Google Docs API
- Google Sheets API
- Google Drive API

At <https://console.cloud.google.com> → pick (or create) a project → **APIs &
Services → Library** → search each one → **Enable**.

### 2. OAuth consent screen

**APIs & Services → OAuth consent screen.** Internal user type if the project lives
in a Google Workspace org you control; otherwise External, and add yourself under
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

### 4. Install

```bash
git clone https://github.com/danlandau321/gmail-batch-email-creator.git
cd gmail-batch-email-creator
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

### 5. First run

```bash
.venv/bin/python batch_emailer.py --doc <doc-url> --sheet <sheet-url> --dry-run
```

A browser opens; sign in as the account whose Gmail should hold the drafts and
approve the permissions:

| Permission asked | What it's for |
|---|---|
| Manage drafts and send email | Creating the drafts. The script only ever calls `drafts.create` — never `send`. |
| See your Google Docs documents | Reading the message Doc. |
| See all your Google Sheets spreadsheets | Reading the recipient Sheet. Read-only. |
| See and download your Google Drive files | Detecting whether `--sheet` is a native Sheet or an uploaded .xlsx. Read-only. |

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
Subject: Quick intro — {{first_name}} <> Example Co

Hi {{first_name}},

Saw {{fund_name}} led the round in ...

Best,
Dan
```

Bold, italics, links and bullet lists are preserved — the draft goes out as HTML
with a plain-text alternative. Tables and inline images are skipped (you get a
warning). Share the Doc with the account you authorized, or just own it.

## The Sheet

Works with a native Google Sheet **or** an .xlsx uploaded to Drive. Nothing is
ever written back — the Sheet is read-only input.

The header row is found automatically (the first row naming an email column), so
title and legend rows above it are fine; `--header-row N` overrides. Tab names
match loosely — `--tab "CEOs & cofounders"` finds "CEOs & Co-founders".

Narrow the list with `--filter COL=VAL`, e.g. `--filter tier=1`, repeatable, and
`VAL` may be a comma-separated list.

**The header names are the merge fields.** Matching ignores case, spaces and
punctuation, so a column `Fund Name` fills `{{fund_name}}` or `{{Fund Name}}`.

| Column | Required | Meaning |
|---|---|---|
| `email` | ✅ | Recipient. `email address`, `recipient` and `to` also work. |
| `cc` / `bcc` | | Extra addresses, comma or semicolon separated |
| `subject` | | Per-row subject, overrides the Doc's |
| `attachments` | | Local file path(s), semicolon separated |
| anything else | | Available to the Doc as `{{that_column}}` |

If there's a `name` column but no `first_name`, the first word is used as
`{{first_name}}` and the last as `{{last_name}}`.

### Re-running

The script keeps **no state**. Every eligible row is drafted on every run, so
running the same command twice gives you two sets of drafts. That's deliberate —
it keeps the Sheet clean and the tool predictable.

In practice this doesn't bite, because nothing sends on its own: you see the
duplicates in Gmail before anyone else does. Start with `--dry-run`, and use a
fresh Sheet per send rather than re-running against an old one.

## Flags

| Flag | Effect |
|---|---|
| `--dry-run` | Print what would be drafted; create nothing. Always start here. |
| `--tab NAME` | Which sheet tab (default: the first one) |
| `--section NAME` | Use one section of a multi-draft Doc, e.g. `"Email 1"` |
| `--filter COL=VAL` | Only rows matching, e.g. `tier=1`; repeatable |
| `--exclude COL=VAL` | Drop rows matching, e.g. `company="Example Co"` |
| `--map PH=COL` | Fill a placeholder from another column, e.g. `name=first_name` |
| `--header-row N` | Point at the header row if auto-detection misses |
| `--limit N` | Only the first N eligible rows |
| `--cc` / `--bcc` | Applied to every draft; repeatable |
| `--attach PATH` | Attach a file to every draft; repeatable |
| `--subject "..."` | Override the Doc's subject line |
| `--allow-missing` | Draft even when a `{{placeholder}}` has no value |
| `--port N` | Fixed localhost port for the OAuth callback |

By default a row whose placeholder has no value is **skipped**, so nobody gets
"Hi ,". Use `--allow-missing` only when you mean it.

Duplicate addresses within one run are skipped automatically.

Every created draft is appended to `drafts_log.csv` (date, row, email, subject,
draft id). That file is gitignored — it contains recipient addresses.

## Typical run

```bash
cd ~/gmail-batch-email-creator

# 1. See what you'd get, without touching Gmail
.venv/bin/python batch_emailer.py --doc <doc> --sheet <sheet> --dry-run

# 2. Two real drafts, eyeball them in Gmail
.venv/bin/python batch_emailer.py --doc <doc> --sheet <sheet> --limit 2

# 3. The rest
.venv/bin/python batch_emailer.py --doc <doc> --sheet <sheet> --cc someone@example.com
```

Step 2 creates two drafts that step 3 will create again — delete those two, or
just accept four drafts for the first two people and delete the extras.

## Troubleshooting

| Message | Fix |
|---|---|
| `Google Docs API has not been used in project ... before` | Enable the API (step 1); the error text links straight to the right page. Wait ~1 min after enabling. |
| `403 The caller does not have permission` | The authorized account can't open that Doc/Sheet — share it. |
| `invalid_grant` / re-auth loop | Delete `token.json` and run again. |
| `Error 400: redirect_uri_mismatch` | `credentials.json` is a Web-application client, not a Desktop one. Redo step 3. |
| `Access blocked: has not completed verification` | Add yourself as a Test user on the OAuth consent screen. |
| `❌ No subject` | Add a `Subject:` first line to the Doc, a `subject` column, or pass `--subject`. |

## License

MIT
