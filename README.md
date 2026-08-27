# docx-toolkit

A CLI (`docxtool`) and Python library over [`python-docx`](https://python-docx.readthedocs.io/) for reading, editing, and creating Microsoft Word (`.docx`) files directly — built for AI agents to drive, but usable from any shell script or Python program.

It exists to avoid the lossy markdown/pandoc round-trip: converting a `.docx` to markdown and back mangles styles, drops formatting, and loses document structure. `docx-toolkit` instead edits the document in place. It never touches `styles.xml` or the document theme (except that `index-add` will create the `index 1`/`index 2`/`index 3` paragraph styles if missing, since a rendered index needs them), and edits preserve run-level formatting (bold/italic/etc.) instead of collapsing it.

## Output format

Every call prints **compact JSON on stdout, one object per invocation**. There is no human-readable mode. On failure, stdout is `{"error": "...", "type": "..."}` and the process exits `1` — check for the `"error"` key rather than relying only on the exit code if you're piping output further.

## Install

```bash
cd scripts
python3 -m venv ../.venv
../.venv/bin/pip install --upgrade pip
../.venv/bin/pip install -e .
```

This installs the `docxtool` command into `.venv/bin/`. Requires Python >= 3.9.

## Quick start

```bash
# Create a new document
docxtool new report.docx
docxtool para-add report.docx "Report Title" --style "Title"

# Inspect an existing document
docxtool read report.docx          # full body content, in document order
docxtool toc report.docx           # just the headings (outline)
docxtool styles report.docx        # every style name + type defined

# Edit a paragraph in place (block_id and hash come from `read`)
docxtool para-edit report.docx 4 "New text" --expect-hash <hash>

# Search
docxtool find report.docx --text "Conventions"
docxtool find report.docx --heading-only
```

## Core concepts

- **Stateless calls**: each invocation opens the file, performs one action (or a `batch` of actions), saves, prints JSON, and exits. There's no persistent server or session.
- **`block_id` is positional**: paragraphs and tables in the document body are numbered in order (0, 1, 2, ...) each time you read. An id is only valid against the state you just read — re-read after any mutation before reusing an id.
- **Containers**: body content is the default; headers/footers are separate addressing spaces, selected with `--container header|footer --section N`.
- **Stale-edit protection**: read/add/list-style commands return a `hash`. Pass it back via `--expect-hash <hash>` on a later mutating call to guarantee nothing else changed the document in between.
- **Batch mode**: pipe a JSON array of `{"op": "<method>", "kwargs": {...}}` to `docxtool batch <path> -` to perform several edits with a single open/save. Supports `"as"`/`$var` result binding to chain dependent operations (e.g. add a table, then fill cells in it, in one call).

## Command overview

| Category | Commands |
|---|---|
| Create / inspect | `new`, `read`, `toc`, `styles`, `section-count`, `find` |
| Paragraphs | `para-get`, `para-add`, `para-edit`, `para-delete` |
| Tables | `table-get`, `table-add`, `table-set-cell`, `table-add-row`, `table-delete-row`, `table-delete`, `table-fill` (batch) |
| Ranges | `delete-range` |
| Comments (body only) | `comment-list`, `comment-add`, `comment-edit` |
| Images (body only) | `image-add`, `image-list`, `image-extract` |
| Headers / footers | `header-get`, `footer-get` |
| Structure | `copy-block`, `move-block` |
| Word fields / index | `field-add`, `xe-add`, `index-add` |
| Integrity | `validate`, `diff` |
| Bulk | `batch` (supports `add_paragraphs`, list items, column add/delete, merge/split paragraphs, and more — see below) |

Some operations (bullet/numbered lists, table column add/delete, paragraph merge/split, copy/move blocks) are only available via `batch`, not as dedicated CLI flags. See **SKILL.md** for the full command reference, JSON shapes, batch examples, and known gotchas.

## Using from Python

```python
from docx_toolkit import DocxDocument

doc = DocxDocument.open("report.docx")
doc.add_paragraph("A new heading", style="Heading 1")
doc.save("report.docx")
```

`scripts/docx_toolkit/core.py` contains the `DocxDocument` class with the full read/write logic; `scripts/docx_toolkit/cli.py` is a thin CLI wrapper around it. Batch op names match `DocxDocument` method names directly.

## Tests

```bash
cd scripts
.venv/bin/python -m pytest
```

## Documentation

See [SKILL.md](SKILL.md) for the complete command reference (all flags, JSON response shapes, batch-mode recipes) and a list of gotchas worth knowing before scripting against this tool.
