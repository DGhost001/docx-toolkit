---
name: docx-toolkit
description: Read, edit, and create Microsoft Word (.docx) files without losing formatting, styles, or structure. Use this whenever a task involves inspecting an existing .docx (outline/TOC, paragraphs, tables, comments, images, headers/footers), editing specific paragraphs/cells/comments in place, or building a new .docx from scratch. Prefer this over converting through pandoc or markdown, which mangles styles and loses formatting fidelity.
---

# docx-toolkit

A CLI over `python-docx` for agents to work with `.docx` files directly,
without going through a lossy markdown/pandoc round-trip. It never touches
`styles.xml` or document theme when editing an existing file — except that
`index-add` will create the `index 1`/`index 2`/`index 3` paragraph styles
if the document doesn't already define them (needed for a rendered dynamic
index); no other command touches styles.xml. Edits otherwise preserve
run-level formatting (bold/italic/etc.) instead of collapsing it.

All output is **compact JSON on stdout, one object per call**. There is no
human-readable mode. On failure, stdout is `{"error": "...", "type": "..."}`
and the process exits 1 — always check for the `"error"` key rather than
relying only on the exit code if you're piping output further.

## Setup

The tool lives in a dedicated venv at `.venv/` inside the skill directory.
The exact path depends on where pi discovers the skill (typically under
`.pi/agent/skills/docx-toolkit/.venv/`). Use `which docxtool` or find it
relative to the skill root:

```bash
docxtool read report.docx   # if .venv/bin is in PATH, or call via full path
```

To locate it explicitly:

```bash
find ~/.pi/agent/skills/docx-toolkit -name docxtool -type f 2>/dev/null
# or on Windows:
dir /s /b %USERPROFILE%\.pi\agent\skills\docx-toolkit\.venv\Scripts\docxtool.exe
```

If the venv is ever missing/corrupted, rebuild it from `scripts/`:

```bash
cd <skill-dir>/scripts
python3 -m venv ../.venv
../.venv/bin/pip install --upgrade pip
../.venv/bin/pip install -e .
```

**Proxy support:** If your network requires a proxy for package installs,
pass it to pip:

```bash
../.venv/bin/pip install --proxy <server:port> -e .
```

The proxy address is environment-specific and should not be committed.

## Core mental model

- **Every call is stateless**: open the file, do one thing (or a `batch` of
  things), save, print JSON, exit. There is no persistent server/session.
- **`block_id` is positional**, assigned fresh each time by walking the
  document body (or a header/footer) in order and numbering paragraphs and
  tables together (0, 1, 2, ...). It is only valid against the state you
  just read — if you mutate the document, **re-read before reusing an id**
  from before that mutation, since ids after the mutation point can shift.
- **Containers**: body content is the default. Headers/footers are separate
  addressing spaces, selected with `--container header|footer --section N`.
- **Stale-edit protection**: every read/add/list-style command returns a
  `hash`. Pass it back via `--expect-hash <hash>` on a later mutating call
  to guarantee nothing else changed the document in between; a mismatch
  returns `{"type": "StaleHashError", ...}` instead of silently overwriting.
- **Comments and images are body-only** in this version (no header/footer
  support) — this is a real python-docx/OOXML limitation, not a shortcut.

## Common workflows

### Inspect a document

```bash
docxtool read report.docx                 # full body content in doc order
docxtool toc report.docx                   # just heading paragraphs (outline)
docxtool styles report.docx                # every style name + type defined
docxtool section-count report.docx
docxtool header-get report.docx --section 0
docxtool footer-get report.docx --section 0
```

`read`/`header-get`/`footer-get` return:
```json
{"hash": "...", "blocks": [
  {"id": 0, "ty": "p", "st": "Heading 1", "tx": "Introduction", "lvl": 1},
  {"id": 1, "ty": "tbl", "st": "Table Grid", "rows": 3, "cols": 2}
]}
```
`ty` is `"p"` (paragraph) or `"tbl"` (table). `lvl` is the heading level
(1, 2, ...) or `null` for non-heading paragraphs.

### Edit an existing paragraph in place

```bash
# Read first to find the block_id and current hash, then:
docxtool para-edit report.docx 4 "New text for this paragraph" \
  --expect-hash <hash-from-read>
```

By default this replaces the *first run's* text and removes any other runs
in that paragraph (simplest case: the paragraph has one style throughout).
If a paragraph has multiple runs with different formatting (e.g. a bolded
word mid-sentence) and you only want to change one of them, target it with
`--run <index>` (0-based, per `paragraph.runs`) to leave the rest alone.

### Add content

```bash
docxtool para-add report.docx "A new heading" --style "Heading 1"
docxtool para-add report.docx "Inserted paragraph" --after 4   # insert after block 4, not at the end
docxtool table-add report.docx 3 2 --style "Table Grid"        # 3 rows x 2 cols
docxtool table-set-cell report.docx 5 0 0 "Field"
docxtool table-add-row report.docx 5 '["Author", "Jane Doe"]'
docxtool image-add report.docx diagram.png --width 4.0         # width/height in inches
```

**Bullet / numbered lists:** List formatting is **only available in batch mode** — there is no CLI `--list` flag (it will fail). Use individual `add_paragraph` calls via batch, each specifying both `"list_type"` and an explicit `"after": <block_id>`:

```bash
cat <<'EOF' | docxtool batch report.docx -
[{"op": "add_paragraph", "kwargs": {"text": "First item", "list_type": "bullet", "after": 4}}]
EOF
```

**⚠️ Pitfalls:**
- **No CLI list flag:** `docxtool para-add ... --list bullet` is not supported. Use batch mode only.
- **No need to always specify `"after"`:** If `"after"` is omitted from kwargs, the batch mode will continue to insert the items where the current 'edit' cursor is positioned. It does **not** default to insert at the end of the document.
- **`add_paragraphs` with `"after"` chains correctly:** When calling `add_paragraphs` with multiple items and an `"after": N`, all items are inserted in sequence right after block N — first item goes after the anchor, each subsequent item chains from the previous one. No need for individual calls.
- **Multiple items with the same list_type share one numbering definition** — this is expected OOXML behavior and not a bug.

**Table columns:** Add or delete columns in existing tables:

```bash
# In Python: doc.add_column(table_id, col_index, ["val1", "val2"])
docxtool table-add-column report.docx <table_id> <col_index> '["A","B"]'
```

Currently `table-add-column` is only available via batch/Python (no dedicated CLI subcommand yet).

### Delete content

```bash
docxtool para-delete report.docx 7                              # delete paragraph at block_id 7
docxtool table-delete report.docx 12                            # delete entire table block
docxtool table-delete-row report.docx 12 3                      # delete row 3 from table 12
docxtool table-delete-column report.docx 12 2                   # delete column 2 from table 12
docxtool delete-range report.docx 5 20                          # delete blocks 5..19 (end exclusive)
```

**Important:** The `<start>` and `<end>` parameters are **block IDs from `read` output**, NOT XML positions.
Block IDs are dense indices into only paragraph (`p`) and table (`tbl`) elements — non-block siblings
like bookmarks, section properties (`sectPr`), and comments are skipped. For example, if the body has
190 XML children but 8 of them are non-block (e.g. `bookmarkStart`, 6×`bookmarkEnd`, `sectPr`), then
there are only 182 block IDs ranging from 0 to 181.

**Range semantics:** `<start>` is **inclusive**, `<end>` is **exclusive**. So `delete-range path 5 20`
deletes blocks with IDs 5, 6, ..., 19 (15 blocks total).

`delete-range` is especially useful for replacing entire sections: find the
start and end block_ids from `read`, then `delete-range` followed by a series of
`para-add --after <anchor>` calls.

### Merge / split paragraphs

```bash
# Merging requires multiple CLI calls or batch mode since merge_paragraphs
# takes a list of block_ids. Use batch:
cat <<'EOF' | docxtool batch report.docx -
[
  {"op": "merge_paragraphs", "kwargs": {"block_ids": [5, 6]}}
]
EOF
```

In Python: `doc.merge_paragraphs([id1, id2])` and `doc.split_paragraph(id, offset)`.

### Search / find

```bash
docxtool find report.docx --text "Conventions"                  # blocks containing this text
docxtool find report.docx --style "Heading 1"                   # all H1 paragraphs
docxtool find report.docx --heading-only                        # all heading blocks (any level)
```

Returns `{"matches": [{"id": ..., "ty": ..., "st": ..., "tx": ..., "lvl": ...}, ...]}`.

`para-add`/`table-add`/`image-add` return `{"id": <new_block_id>, "hash": ...}`.

### Comments (body only)

```bash
docxtool comment-add report.docx 4 0 2 "Please rephrase" --author "Reviewer" --initials "RV"
docxtool comment-list report.docx
docxtool comment-edit report.docx 0 --text "Rephrase this more clearly"
```

`comment-add` anchors to a **run range** `run_start run_end` (end-exclusive)
within paragraph `block_id` — read the paragraph first if you need to know
how many runs it has. There is no delete or reply-threading support
(python-docx doesn't expose it); only add/list/edit.

### Images

```bash
docxtool image-list report.docx
docxtool image-extract report.docx 0 out.png
```

### Batch (multiple edits, one save)

For several edits in one shot (avoids re-opening/re-saving per call), pipe a
JSON array of `{"op": "<DocxDocument method name>", "kwargs": {...}}` to
`batch`. Method names match `docx_toolkit.core.DocxDocument` directly
(`add_paragraph`, `edit_paragraph`, `set_cell`, `add_comment`, etc.):

```bash
cat <<'EOF' | docxtool batch report.docx -
[
  {"op": "add_paragraph", "kwargs": {"text": "Revision History", "style": "Heading 2"}},
  {"op": "add_table", "kwargs": {"rows": 1, "cols": 3, "style": "Table Grid"}}
]
EOF
```

**Variable binding (`as` key):** Use `"as": "<name>"` to bind a result to
a variable, then reference it in later ops with `$<name>`:

```bash
cat <<'EOF' | docxtool batch report.docx -
[
  {"op": "add_table", "kwargs": {"rows": 3, "cols": 2}, "as": "fmt_tbl"},
  {"op": "set_cell", "kwargs": {"block_id": "$fmt_tbl", "row": 0, "col": 0, "text": "Format"}},
  {"op": "set_cell", "kwargs": {"block_id": "$fmt_tbl", "row": 0, "col": 1, "text": "Name"}},
  {"op": "add_row", "kwargs": {"block_id": "$fmt_tbl", "values": ["B", "Boolean"]}}
]
EOF
```
The variable holds the new block's `id` (or `row` index for row additions).
This lets you chain dependent operations in a single batch call.

**Bulk paragraph insertion:** Use the `add_paragraphs` op to insert many
paragraphs at once:

```json
{"op": "add_paragraphs", "kwargs": {
  "after": "$conv_id",
  "items": [
    {"text": "ECSS Reference: ECSS-E-ST-40_0860403"},
    {"text": "Number Notation", "style": "Heading 2"},
    {"text": "Unless stated otherwise..."}
  ]
}}
```

If any op fails, the batch **aborts and does not save** — you get
`{"ok": false, "failed_at": <index>, "error": "...", "results": [...]}`
with results only for ops before the failure. On full success:
`{"ok": true, "results": [<return value of each op, in order>]}`.

**Copy / move blocks:**

```bash
cat <<'EOF' | docxtool batch report.docx -
[
  {"op": "copy_block", "kwargs": {"block_id": 5, "after": 3}},
  {"op": "move_block", "kwargs": {"block_id": 7, "before": "$var0"}}
]
EOF
```

**Table fill (bulk cell population from 2D array or CSV):**

```bash
cat <<'EOF' | docxtool batch report.docx -
[
  {"op": "add_table", "kwargs": {"rows": 4, "cols": 3}, "as": "tbl1"},
  {
    "op": "table_fill",
    "kwargs": {
      "block_id": "$tbl1",
      "data": [["Name","Age","City"],["Alice","30","Berlin"]],
      "header": ["Name","Age","City"]
    }
  }
]
EOF
```

CSV string alternative: `"data": "Name,Age,City\nAlice,30,Berlin"`.

**Validate document integrity:**

```bash
docxtool batch report.docx - <<'EOF'
[{"op": "validate", "kwargs": {}}]
EOF
```

Returns `{"valid": true/false, "errors": [...], "warnings": [...]}`.

**Diff two documents (Python only):**

```python
from docx_toolkit import DocxDocument
doc1 = DocxDocument.open("original.docx")
doc2 = DocxDocument.open("revised.docx")
diff = DocxDocument.diff(doc1, doc2)
# diff["added"]    → blocks in doc2 not in doc1
# diff["removed"]  → blocks in doc1 not in doc2
# diff["modified"] → blocks with changed text
```

### Create a new document from scratch

```bash
docxtool new draft.docx
docxtool para-add draft.docx "Report Title" --style "Title"
```

## Gotchas

- `toc`/`outline()` recognizes English (`Heading N`) and German
  (`Ueberschrift1`, `Abschnitt2`) heading styles, plus any style ending with
  a digit. For completely custom style names (e.g. `Annex1`, `DRD3`), use
  `find --heading-only` or filter `read()`'s blocks by `st` client-side.
- `block_id` is **not stable across mutations** — re-`read` after inserting/
  deleting to get fresh ids for further edits in the same region. In batch
  mode, use `"as": "name"` and `$name` references to chain dependent ops
  without needing fresh reads between steps.
- Accessing a section's header/footer (even just to read it) will
  auto-create an empty one if the document didn't already define it for
  that section — this is python-docx behavior, not something this tool
  adds on top.
- `table-set-cell`/table row edits collapse each cell to a single paragraph/
  run — fine for typical data tables, not appropriate for cells that need
  rich per-run formatting.
- Editing a paragraph's text (without `--run`) discards any runs beyond the
  first — intentional (that's what "replace this paragraph's text" means),
  but don't use it on paragraphs where you need to keep multiple
  differently-formatted runs; use `--run` instead.
- Headlines use autonumbering, if not specified otherwise by the user, do not
  add numbers to headings.
- The tool does not automatically change a "-" paragraph to a item list. Use
  the explicit call to generate lists and numbered lists
- Do not inline comments into the text, use the proper comment function for
  creating comments.

## Source layout

- `scripts/docx_toolkit/core.py` — `DocxDocument` class, the actual
  read/write logic (importable directly from Python if you'd rather not
  shell out to the CLI).
- `scripts/docx_toolkit/cli.py` — the `docxtool` CLI wrapping `core.py`.
- `scripts/tests/` — pytest suite (run via `.venv/bin/python -m pytest`
  from `scripts/`). Includes:
  - `test_extensions.py` — table-delete, delete-range, contextual batch
    with `$var` refs, add_paragraphs bulk op, heading style detection,
    find command, CLI integration tests
  - `test_p2.py` — list support (bullets/numbering), table column ops
    (add/delete), merge-paragraphs, split-paragraph
  - `test_p3.py` — copy/move blocks, table-fill (2D array + CSV),
    validate document integrity, diff two documents
  - `test_p4.py` — CLI subcommands for copy-block, move-block,
    table-fill, validate, diff; advanced diff (character-level);
    template support stubs; auto-numbering continuation
  - `test_p5.py` — Word field support: add_field, add_xe (hidden index
    entries), add_index (dynamic INDEX span + index N styles), validate
    field-structure checks
  - `test_p6.py` — CLI subcommands for field-add, xe-add, index-add;
    batch chaining of the same ops via `$var` refs

## New commands reference

| Command / Op | Description |
|--------------|-------------|
| `table-delete <path> <id>` | Delete entire table block |
| `delete-range <path> <start> <end>` | Delete blocks in range (end exclusive) |
| `find <path> --text "..." [--style ...] [--heading-only]` | Search blocks by text/style/level |
| Batch `add_paragraphs` op | Insert multiple paragraphs at once |
| Batch `$var` refs + `as` key | Chain dependent ops in one batch call |
| Batch `add_column` / `delete_column` | Add/remove columns in existing tables |
| Batch `merge_paragraphs` | Merge adjacent paragraphs into one |
| Batch `split_paragraph` | Split a paragraph at character offset |
| Batch `add_paragraph` w/ `list_type: "bullet"|"number"` | Add paragraph as OOXML list item (batch mode only, no CLI flag) |
| Batch `copy_block` / `move_block` | Copy or move blocks between positions |
| Batch `table_fill` | Bulk-fill table cells from 2D array or CSV |
| `validate <path>` | Check document integrity (block gaps, table consistency) |
| `diff <path1> <path2>` | Compare two documents — returns `{added, removed, modified}` |
| Batch `copy_block` / `move_block` | Copy or move blocks between positions |
| Batch `table_fill` | Bulk-fill table cells from 2D array or CSV |
| `field-add <path> <id> --instr "..."` | Insert a generic Word field (begin/instrText/end) inline at a paragraph (TOC, REF, PAGEREF, XE, ...) |
| `xe-add <path> <id> --term "..." [--see "..."]` | Insert a hidden `XE` index-entry field; `--term "parent:sub"` nests, `--see` adds a cross-reference |
| `index-add <path> --after <id> [--entries JSON] [--xe-pairs JSON]` | Build a dynamic `INDEX` field span (open+cache+close) with explicit, caller-supplied entries/anchors — no heuristic placement |
| Batch `add_field` / `add_xe` / `add_index` | Same three ops available in batch mode, chainable via `$var` refs |
