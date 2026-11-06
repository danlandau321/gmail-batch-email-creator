---
name: batch-email
description: Create Gmail drafts from a Google Doc (the message) plus a Google Sheet (the recipients). Use whenever the user says something like "send the email in <doc> to the people in <sheet>", "draft the email in this doc to everyone in that spreadsheet", "mail merge", or "batch email". Creates DRAFTS only, never sends.
---

# Batch email: Doc + Sheet -> Gmail drafts

Wraps the working script at `~/gmail-batch-email-creator/batch_emailer.py`. It merges a Google
Doc (the message copy) with a Google Sheet (one row per recipient) and creates one
Gmail **draft** per row. It cannot send — the OAuth scope is `gmail.compose`.

## Invocation

Always use the venv interpreter, from that directory:

```bash
cd ~/gmail-batch-email-creator && .venv/bin/python batch_emailer.py --doc <DOC> --sheet <SHEET> [flags]
```

`--doc` and `--sheet` accept a full URL or a bare file id.

## Procedure — follow in order

1. **Extract the two URLs** from the user's sentence. The Doc is the message;
   the Sheet is the recipients. If either is missing, ask for it — do not guess,
   and do not go hunting through Drive unless the user asks you to.

2. **Map the extra intent onto flags.** Common phrasings:

   | User says | Flag |
   |---|---|
   | "only tier 1", "just the West Coast ones" | `--filter COL=VAL` (repeatable; VAL may be comma-separated) |
   | "skip the ones already contacted" | `--exclude COL=VAL` |
   | "cc Ellen" | `--cc <addr>` (repeatable) |
   | "just the first 5", "test on a couple" | `--limit N` |
   | "use the second tab", "the 2026 tab" | `--tab <name>` |
   | "use Email 2 from the doc" | `--section 'Email 2'` |
   | "attach the deck" | `--attach <path>` |
   | a subject that overrides the Doc | `--subject '<text>'` |

   Placeholders: `{{column_name}}` in the Doc fill from Sheet headers,
   case/space-insensitive. `--map PLACEHOLDER=COL` fixes a mismatch.

3. **Dry run first. Always.** Append `--dry-run` and run it. This creates nothing.
   It prints, per eligible row, the To / Subject / Cc / Bcc and the **first 6 lines**
   of the merged body, then a total count.

   Read that output for two failure signals and surface them before going further:
   - `⚠️ unfilled <name>` — a Doc `{{placeholder}}` has no matching Sheet column.
     The draft would go out with a blank there. Stop and get it fixed.
   - `⏭️ skipped` — the row had no value for a required placeholder, so it was
     dropped. Tell the user which rows and why; a silently smaller send is the
     thing they most need to know about.

   Show the user the row count and paste one row's preview so they can eyeball
   that the merge fields resolved to real values.

4. **Get explicit confirmation** before the real run. Say plainly how many drafts
   will be created. Creating dozens of drafts in someone's mailbox is outward-facing
   and tedious to undo, so never skip this step, even if the user sounded decisive.

5. **Real run**: same command minus `--dry-run`. Report the number of drafts created
   and any rows that errored.

## Things to know

- **Re-runs are NOT safe.** The script has no skip logic: it does not write a
  `status` column and there is no `--force` or `--no-mark`. Every run creates a
  draft for every eligible row, so re-running to add one person duplicates the
  whole batch (this happened on 28 Oct 2026). To add a person, put them in a
  separate tab or filter to them with `--filter email=<addr>`. To re-create a
  batch after a copy change, delete the old drafts first (ids are in
  `drafts_log.csv`; the Gmail API `drafts.delete` works with the script's token).
- **First run on a new machine** opens a browser for Google consent and writes
  `token.json`. If you see an auth error or the command appears to hang, tell the
  user to run it themselves with the `!` prefix so they can complete the browser
  flow — it needs their hands, not yours.
- **The Sheet needs an `email` column** (or `email address` / `recipient`).
  Optional columns: `cc`, `bcc`, `subject`, `attachments`, `status`, `draft_id`.
- The Doc's first line may be `Subject: ...`; everything after is the body. Bold,
  italics, links and bullets survive. Tables and inline images are dropped.
- `--allow-missing` lets rows through that have empty placeholder values, and
  `--header-row N` handles a sheet whose headers aren't on row 1.

## Reporting back

Give the user the count and anything that went wrong. Don't paste all the drafts —
they're in Gmail. Do paste the one sample during the dry run, because that's the
step where a bad merge gets caught.
