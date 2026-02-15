---
name: pst-to-markdown
description: Extract emails from Outlook PST files into organised markdown archives. Use when needing to convert PST files to markdown, extract email archives, process Outlook exports, or create searchable email collections. Trigger on phrases like "extract pst", "convert pst", "pst to markdown", "email archive", "extract outlook".
---

# PST Email Extraction

Extract emails from Outlook PST files into an organised, integrity-verified archive of markdown files, raw email backups, and attachments. Supports full extraction and incremental append mode.

## Prerequisites

- Python virtual environment set up (run setup.sh if not done)
- At least one of: `libratom` (Python) or `readpst` (system tool from pst-utils)

### First-Time Setup

```bash
# Set up Python environment (one-time)
~/.claude/skills/pst-to-markdown/setup.sh
```

### System Dependencies (optional fallback)

If libratom installation fails, install readpst as a fallback:

```bash
# Ubuntu/Debian
sudo apt install pst-utils

# macOS
brew install libpst
```

## Extraction Operations

### Full Extraction

Extract all emails from a PST file into markdown:

```bash
# Basic extraction
~/.claude/skills/pst-to-markdown/.venv/bin/python ~/.claude/skills/pst-to-markdown/scripts/extract_pst.py /path/to/file.pst /path/to/output/

# Verbose output with progress
~/.claude/skills/pst-to-markdown/.venv/bin/python ~/.claude/skills/pst-to-markdown/scripts/extract_pst.py /path/to/file.pst /path/to/output/ --verbose

# Include deleted items
~/.claude/skills/pst-to-markdown/.venv/bin/python ~/.claude/skills/pst-to-markdown/scripts/extract_pst.py /path/to/file.pst /path/to/output/ --include-deleted --verbose

# Set timezone for date display
~/.claude/skills/pst-to-markdown/.venv/bin/python ~/.claude/skills/pst-to-markdown/scripts/extract_pst.py /path/to/file.pst /path/to/output/ --timezone "Europe/London"

# Fix MAILER-DAEMON sent items (provide the PST owner's email)
~/.claude/skills/pst-to-markdown/.venv/bin/python ~/.claude/skills/pst-to-markdown/scripts/extract_pst.py /path/to/file.pst /path/to/output/ --owner-email "user@example.com"
```

### Incremental Extraction (Append Mode)

Add only new emails (skips already-extracted messages by Message-ID):

```bash
~/.claude/skills/pst-to-markdown/.venv/bin/python ~/.claude/skills/pst-to-markdown/scripts/extract_pst.py /path/to/file.pst /path/to/output/ --append --verbose
```

### Extract from Pre-Extracted .eml Directory

If emails were already extracted with readpst elsewhere, point at the directory:

```bash
~/.claude/skills/pst-to-markdown/.venv/bin/python ~/.claude/skills/pst-to-markdown/scripts/extract_pst.py /path/to/eml-directory/ /path/to/output/
```

## Output Structure

```
output/
├── emails/
│   ├── FolderName/
│   │   ├── 2023-01-15_093042_from-john.smith_to-jane.doe_RE-Subject/
│   │   │   ├── email.md              # Formatted markdown with YAML frontmatter
│   │   │   ├── email.eml             # Raw original email (RFC 822)
│   │   │   ├── attachment_001_doc.pdf # Extracted attachments
│   │   │   └── checksums.sha256      # Per-email integrity hashes
│   │   └── .../
│   └── .../
├── index.csv                          # Machine-readable master index
├── index.md                           # Human-readable index with timeline
├── extraction_log.txt                 # Processing log with statistics
└── manifest.sha256                    # Master integrity manifest
```

### Email Markdown Format

Each `email.md` contains:
- **YAML frontmatter**: message_id, date, from, to, cc, subject, attachments with SHA256 hashes
- **Formatted body**: HTML converted to markdown, or plain text preserved
- **Attachment links**: Relative links to extracted files with sizes
- **Original headers**: Full RFC 822 headers in code block

### Index Files

- **index.csv**: All emails with date, sender, recipient, subject, folder, attachment count
- **index.md**: Timeline view grouped by year/month with links to each email

## CLI Reference

```
extract_pst.py [-h] [--include-deleted] [--timezone TZ] [--verbose] [--append] [--owner-email EMAIL] pst_file output_dir
```

| Argument | Description |
|----------|-------------|
| `pst_file` | Path to PST file, or directory of pre-extracted .eml files |
| `output_dir` | Output directory (created if needed) |
| `--include-deleted` | Include deleted items from PST |
| `--timezone TZ` | Target timezone for dates (default: UTC) |
| `--verbose`, `-v` | Verbose output with per-email logging |
| `--append` | Skip emails already in archive (by Message-ID) |
| `--owner-email EMAIL` | PST owner's email (fixes MAILER-DAEMON in sent items) |

## Extraction Backends

The tool tries backends in priority order:

1. **libratom** (Python) — preferred, installed via requirements.txt
2. **readpst** (system CLI) — fallback, from pst-utils package
3. **Directory mode** — processes pre-extracted .eml files directly

## Integrity Verification

Every extraction produces a verifiable chain of custody:

1. Each email folder has `checksums.sha256` (SHA256 of all its files)
2. `manifest.sha256` hashes all checksum files plus the index
3. Source PST SHA256 is recorded in the manifest

To verify: `sha256sum -c manifest.sha256`

## Workflow: Extract and Search

For large archives, combine with the repo-search skill:

1. Extract PST to markdown with this skill
2. Index the output with repo-search ingest
3. Semantic search across all emails

```bash
# Step 1: Extract
~/.claude/skills/pst-to-markdown/.venv/bin/python ~/.claude/skills/pst-to-markdown/scripts/extract_pst.py archive.pst ./email-output/ --verbose

# Step 2: Index for search (uses repo-search skill)
~/.claude/skills/repo-search/.venv/bin/python ~/.claude/skills/repo-search/ingest.py ./email-output/ --verbose

# Step 3: Search
~/.claude/skills/repo-search/.venv/bin/python ~/.claude/skills/repo-search/query.py --db-path ./email-output/.vectordb search "settlement agreement"
```

## Error Handling

- **"readpst not found"**: Install pst-utils or ensure libratom is installed via setup.sh
- **Corrupt emails**: Logged to extraction_log.txt, processing continues
- **Encoding issues**: Falls back through UTF-8 → latin-1 → raw bytes
- **Duplicate timestamps**: Appended with -001, -002 suffixes
- **Path too long**: Subject truncated, uniqueness preserved

## Performance

| Scenario | Approximate Speed |
|----------|-------------------|
| Emails without attachments | ~5,000/hour |
| Emails with attachments | ~2,000/hour |

A typical 300MB PST (~1,000-3,000 emails) processes in 5-15 minutes.
