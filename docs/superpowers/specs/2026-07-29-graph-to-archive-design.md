# Keeping a PST archive current from live mail

**Status:** design, approved 29 July 2026
**Skills touched:** `outlook-graph`, `pst-to-markdown`

## Problem

`pst-to-markdown` turns a PST export into a markdown archive. That archive is
frozen the moment it is produced: it holds the mail that was in the PST and
nothing after it. Mail that arrives from then on lives only in Outlook.

The two halves of the pack already cover the two halves of the problem - the
mail you were handed in a box, and the mail you are handling now - but nothing
joins them. There is no way to add today's mail to yesterday's archive in the
shape the archive already uses.

## Goal

One continuous archive. The PST backfills history; live mail appends into the
identical structure, deduplicated, with the index and integrity manifest
covering both.

```
archive/
├── emails/
│   ├── Inbox/Cherise/2026-07-21_150500_from-.../   <- from the PST
│   └── Inbox/Cherise/2026-07-29_101200_from-.../   <- from Graph
├── index.csv          both, deduplicated by Message-ID
├── index.md
└── manifest.sha256
```

Non-goal: narrative or topic-level summarisation. That is a separate practice
and stays out of this design.

## Approach

`extract_pst.py` needs no changes. Three behaviours it already has make the
join almost free:

| Behaviour | Where |
|---|---|
| Accepts a directory of `.eml` in place of a PST | `extract_pst.py:386` |
| Mirrors that directory's layout into the archive | `extract_pst.py:502`, `618-621` |
| `--append` skips any Message-ID already in `index.csv` | `extract_pst.py:584` |

So `staging/Inbox/Cherise/anything.eml` lands at
`archive/emails/Inbox/Cherise/<date>_<from>_<to>_<subject>/`. The staging
directory's *layout* is the archive's layout; the `.eml` filename itself is
never read for naming and only has to be unique.

The one missing piece is that `outlook-graph` cannot currently write `.eml`.
Graph serves raw RFC 822 MIME at `GET /me/messages/{id}/$value` - the same
`$value` pattern the attachment downloader already uses
(`outlook-graph-mail.sh:1930`). That MIME carries the real `Message-ID` header,
which is exactly what `--append` deduplicates on.

### Why staging rather than writing the archive directly

`export` writes plain `.eml` and stops. It does not know about `email.md`,
checksums or the index. This keeps the skills independently changeable: the
archive format belongs to `pst-to-markdown` alone, and `outlook-graph` gains a
generally useful export verb rather than a coupling to another skill's
on-disk format.

## Work

### 1. `export` verb in `outlook-graph-mail.sh`

```
outlook-graph-mail.sh export <folder> <output-dir> [--since YYYY-MM-DD] [--count N]
```

- Resolve `<folder>` with the existing `resolve_folder_id`, so folder naming
  behaves exactly as it does for `move`, `folder` and the rest.
- Page the message list via `@odata.nextLink`. `$top` caps at 1000 per page and
  a mailbox folder can hold far more, so paging is required, not optional.
- For each message, write MIME to
  `<output-dir>/<resolved folder path>/<YYYYMMDD>_<HHMMSS>_<short-id>.eml`,
  where the timestamp is `receivedDateTime` in UTC and `<short-id>` is the last
  20 characters of the Graph message id - the same slice the listing commands
  print (`format_message`, `.[-20:]`) and that `resolve_message_id` matches on,
  so an id copied from a listing appears verbatim in the filename. Any `/` is
  stripped before use so a server-supplied id cannot escape the output
  directory. The timestamp keeps a directory listing chronological; the short
  id guarantees uniqueness when two messages share a second.
- `--since YYYY-MM-DD` adds a Graph `$filter` on `receivedDateTime`. Chosen
  over reading a watermark out of `index.csv` deliberately: a watermark would
  couple `outlook-graph` to `pst-to-markdown`'s file format and stop the two
  changing independently. `--append` still backstops any overlap, so a `--since`
  that reaches back too far costs bandwidth and nothing else.
- `--count N` caps total messages, for spot checks.

Failure handling follows the attachment downloader: `curl -sf … -o`, and on
non-zero remove the part-written file and report it before continuing. A Graph
error JSON must never be left on disk named `.eml`, where the next step would
parse it as an email.

### 2. Documented two-step workflow

In both `SKILL.md` files:

```bash
outlook-graph-mail.sh export Inbox/Cherise ./staging/ --since 2026-07-01

${CLAUDE_SKILL_DIR}/../pst-to-markdown/.venv/bin/python \
  ${CLAUDE_SKILL_DIR}/../pst-to-markdown/scripts/extract_pst.py \
  ./staging/ ./archive/ --append
```

The `${CLAUDE_SKILL_DIR}/../<other>` cross-skill form already resolves under
both installers - `install-codex.sh` rewrites it to a sibling Codex path and
says so in comments; `install.sh` symlinks whole directories, so the relative
path resolves naturally.

### 3. Tests

- `helpers_test.sh` - `export` rejects a missing folder or output directory,
  and `--since` rejects a malformed date rather than silently sending it to
  Graph as a filter that matches nothing.
- `test_extract_pst.py` - a directory-mode round trip: run twice with
  `--append`, assert the second run adds no email folders and leaves
  `index.csv` unchanged in row count.

## Known wrinkles

**The `pst_folder` column.** Graph-sourced mail is recorded under a column
named `pst_folder`, which is now a slight misnomer. Keeping the name is the
deliberate choice: it is the same archive and the same index, and renaming the
column would break every existing `index.csv` and anything reading one. The
column means "the folder this message came from", and it always did.

**Attachments.** `$value` MIME carries attachments inline, so
`extract_pst.py` extracts and hashes them exactly as it does for PST-sourced
mail. No separate attachment fetch is needed.

**Sent items.** Exporting a Sent folder works unchanged. The `--owner-email`
flag that fixes MAILER-DAEMON senders is a PST-extraction concern and does not
apply to Graph MIME, which carries a correct `From`.

## Out of scope

- Narrative or topic-level records synthesised across threads
- Any scheduled or automatic sync; this is a command you run
- Changes to the archive format itself
