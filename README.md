# sort_requests

Sorts a flat folder of dated files into a `YYYY/MM/DD/` tree, based on the date
found in each filename.

**Before**

```
request/
  apple_20260212.xls
  anthropic_20260213.csv
  microsoft_20260215.xlsx
```

**After**

```
request_by_date/
  2026/
    02/
      12/apple_20260212.xls
      13/anthropic_20260213.csv
      15/microsoft_20260215.xlsx
```

## Design principle

The script only files what it can determine with certainty. Anything ambiguous —
an unparseable name, an impossible date, a broken file — is moved to a
quarantine folder and reported, never guessed at.

A misfiled document does not announce that it was misfiled, so "not sorted, but
reported" always beats "sorted, but wrong."

## Requirements

Python 3.8+. **Standard library only** — no `pandas`, no `openpyxl`, nothing to
install. Developed and tested on Python 3.14 / macOS.

> Console output and `--help` text are in Korean.

## Quick start

```bash
python3 sort_requests.py                              # dry-run (default)
python3 sort_requests.py --apply                      # actually move files
python3 sort_requests.py --apply --copy               # copy instead, keep originals
python3 sort_requests.py --undo sort_manifest.jsonl   # roll back the last run
```

**Nothing is written without `--apply`.** The default run only prints the plan
and writes the CSV report, so you can inspect what would happen first.

## Options

| Option | Default | Meaning |
|---|---|---|
| `src` (positional) | `request` | Source folder, scanned recursively |
| `-o`, `--dest` | `request_by_date` | Output folder |
| `--apply` | off | Perform the moves (without it: dry-run) |
| `--copy` | off | Copy instead of move, preserving originals |
| `--report` | `sort_report.csv` | Per-file report |
| `--manifest` | `sort_manifest.jsonl` | Move log used by `--undo` |
| `--undo MANIFEST` | — | Reverse a previous run |

## How files are classified

### Dates

Only `YYYY MM DD` order is accepted, with or without a separator:

| Accepted | Rejected (quarantined) | Why |
|---|---|---|
| `apple_20260212.xls` | `apple_02122026.xls` | MMDDYYYY or DDMMYYYY? |
| `apple_2026-02-12.xls` | `apple_260212.xls` | two-digit year, ambiguous |
| `apple 2026.02.12.xlsx` | `apple_20260230.xls` | February 30 does not exist |
| `apple_2026_02_12.csv` | `report_20260212_to_20260215.csv` | two different dates |

Years outside 1990–2100 are treated as a coincidental run of digits, not a date.

### File integrity

The first 512 bytes are checked against magic-byte signatures (OLE2, ZIP, PDF,
HTML/XML, text encoding). A file whose real format contradicts its extension —
a `.xls` that is actually HTML, a common Excel "Save as Web Page" artifact — is
**still filed normally** and flagged in the report. Empty and unreadable files
are quarantined instead.

### Outcomes

| Status | Destination |
|---|---|
| `placed` | `DEST/YYYY/MM/DD/` |
| `duplicate` | `DEST/_duplicates/YYYY-MM-DD/` — byte-identical, kept for review, never deleted |
| `unresolved` | `DEST/_unresolved/` — date could not be determined, or file is broken |
| `already_placed` | not moved — an identical file is already at the destination |
| `skipped` | not moved — `.DS_Store`, `._*`, `~$*`, `Thumbs.db`, symlinks |

Name collisions with **different** content get a `__2`, `__3` suffix so nothing
is overwritten.

## Report

`sort_report.csv` is written on every run, including dry-runs, in UTF-8 with BOM
so it opens cleanly in Excel:

```
status, reason, flags, parsed_date, source_path, dest_path, ext, detected_format, size_bytes
```

The console output additionally groups near-identical filenames, to surface
typos and casing differences:

```
[SIMILAR NAMES]
  anthropic    2 files  2026-02-13, 2026-02-20
  Anthropic    1 file   2026-02-14
  anthropicc   1 file   2026-02-18
```

**Filenames are never renamed.** This is an advisory report only — fix the
originals yourself and re-run if you want them merged.

## Undo

Each move is appended to the manifest and fsynced *before* it happens, so an
interrupted run is still reversible:

```bash
python3 sort_requests.py --undo sort_manifest.jsonl
```

Re-running is safe. Files already at their destination are recognised by hash
and skipped rather than duplicated.

## Testing

`make_fixtures.py` builds a `request/` folder covering every edge case, with the
expected verdict written next to each fixture as a comment:

```bash
python3 make_fixtures.py          # create request/ (refuses if it exists)
python3 make_fixtures.py --force  # recreate it
python3 sort_requests.py          # inspect the dry-run against the comments
```

Covered: separator variants, year boundaries, casing collisions, typos, invalid
and ambiguous dates, extension/format mismatches, zero-byte and unreadable
files, duplicates, subfolders, OS junk, Korean filenames, and CP949-encoded CSV.

## Known issues

None that risk data loss. A code review found two such defects in the
case-sensitivity probe — both are fixed, and the broken-symlink case that
triggered one of them is now a permanent fixture in `make_fixtures.py`.

Remaining lower-severity items, in rough priority order:

- The similar-name clustering is O(n²) over distinct name tokens and runs
  unconditionally, including in dry-run (~24s at 3,000 files, minutes at 10k).
- Name collisions among *quarantined* files are not counted in the summary line,
  so a `__2` rename inside `_unresolved/` happens without being announced.
- A copy that fails midway (disk full, I/O error) leaves a truncated file in the
  output tree; it is not cleaned up.
- The third and later byte-identical copies of a file are labelled
  `already_placed` and quietly left in the source folder.
- Files whose `stat()` fails during collection are counted as quarantined but
  never actually moved to `_unresolved/`.

## Development log

`2026-08-31-163101-i-have-to-write-a-script-that-sorts-and-reorders.txt` is the
full transcript of the session that produced this tool — requirements gathering,
the plan, the implementation, the verification runs, and the code review that
caught the two case-sensitivity defects. Kept for coursework reference. One
email address is redacted; everything else is verbatim.

## Out of scope

By design, this script does **not**:

- read file *contents* to extract dates (requires `pandas`/`openpyxl`, and needs
  a policy for when the content disagrees with the filename)
- fall back to file modification time (easily corrupted by copying/downloading)
- rename files or normalise company names
- delete anything, including `_duplicates/`
- parse English month names (`Feb 12 2026`) or two-digit years
