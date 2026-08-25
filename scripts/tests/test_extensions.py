"""Tests for extended docx-toolkit functionality.

Covers: table-delete, delete-range, contextual batch ($var refs),
add_paragraphs bulk op, heading style detection (German styles), find command.
"""

import json
import pytest

from docx_toolkit import DocxDocument, StaleHashError


# =============================================================================
# 1. table-delete — delete entire table block
# =============================================================================

class TestTableDelete:
    def test_delete_table_removes_block(self):
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Before")
        t1 = doc.add_table(rows=2, cols=2)
        p2 = doc.add_paragraph("After")

        doc.delete_table(t1)

        blocks = doc.read()["blocks"]
        types = [b["ty"] for b in blocks]
        assert types == ["p", "p"]
        texts = [b["tx"] for b in blocks if b["ty"] == "p"]
        assert texts == ["Before", "After"]

    def test_delete_table_unknown_id_raises(self):
        doc = DocxDocument.new()
        from docx_toolkit import DocxError
        with pytest.raises(DocxError):
            doc.delete_table(9999)

    def test_delete_table_with_stale_hash_raises(self):
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=1, cols=1)
        stale = "stale-hash"
        with pytest.raises(StaleHashError):
            doc.delete_table(t1, expect_hash=stale)

    def test_delete_last_table(self):
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=1, cols=1)
        doc.delete_table(t1)
        blocks = doc.read()["blocks"]
        assert len(blocks) == 0


# =============================================================================
# 2. delete-range — delete all blocks between two IDs
# =============================================================================

class TestDeleteRange:
    def test_delete_range_removes_blocks(self):
        doc = DocxDocument.new()
        doc.add_paragraph("Keep1")       # id=0
        doc.add_paragraph("Remove1")     # id=1
        doc.add_paragraph("Remove2")     # id=2
        doc.add_paragraph("Keep2")       # id=3

        doc.delete_range(1, 3)  # delete ids 1 and 2 (end exclusive)

        blocks = doc.read()["blocks"]
        texts = [b["tx"] for b in blocks]
        assert texts == ["Keep1", "Keep2"]

    def test_delete_range_empty_range(self):
        from docx_toolkit import DocxError
        doc = DocxDocument.new()
        doc.add_paragraph("A")  # id=0
        doc.add_paragraph("B")  # id=1
        with pytest.raises(DocxError, match="empty range"):
            doc.delete_range(1, 1)

    def test_delete_range_out_of_bounds_raises(self):
        doc = DocxDocument.new()
        doc.add_paragraph("A")
        from docx_toolkit import DocxError
        with pytest.raises(DocxError):
            doc.delete_range(0, 99)

    def test_delete_range_with_tables_mixed(self):
        doc = DocxDocument.new()
        doc.add_paragraph("P1")          # id=0
        t1 = doc.add_table(rows=1, cols=1)  # id=1
        doc.add_paragraph("P2")          # id=2
        t2 = doc.add_table(rows=1, cols=1)  # id=3
        doc.add_paragraph("P3")          # id=4

        doc.delete_range(1, 4)  # delete table+para+table (ids 1,2,3)

        blocks = doc.read()["blocks"]
        types = [b["ty"] for b in blocks]
        assert types == ["p", "p"]
        texts = [b["tx"] for b in blocks if b["ty"] == "p"]
        assert texts == ["P1", "P3"]

    def test_delete_range_with_hash(self):
        doc = DocxDocument.new()
        doc.add_paragraph("A")
        doc.add_paragraph("B")
        doc.add_paragraph("C")
        h = doc.content_hash
        doc.delete_range(0, 1, expect_hash=h)
        blocks = doc.read()["blocks"]
        assert len(blocks) == 2

    def test_delete_all_blocks(self):
        doc = DocxDocument.new()
        doc.add_paragraph("A")
        doc.add_paragraph("B")
        doc.delete_range(0, 2)
        assert doc.read()["blocks"] == []


# =============================================================================
# 3. Contextual batch — $var references in kwargs
# =============================================================================

class TestContextualBatch:
    def test_batch_with_var_binding(self):
        """Add a table, then fill its cells using the bound variable."""
        from docx_toolkit.cli import _run_batch

        doc = DocxDocument.new()
        ops = [
            {"op": "add_table", "kwargs": {"rows": 2, "cols": 2}, "as": "t1"},
            {"op": "set_cell", "kwargs": {"block_id": "$t1", "row": 0, "col": 0, "text": "Header"}},
            {"op": "add_row", "kwargs": {"block_id": "$t1", "values": ["a", "b"]}},
        ]
        result = _run_batch(doc, ops)
        assert result["ok"] is True

        table = doc.get_table(0)  # only one block now (the table)
        assert table["cells"][0] == ["Header", ""]
        assert table["rows"] == 3  # 2 original + 1 added

    def test_batch_var_not_found_raises(self):
        from docx_toolkit.cli import _run_batch

        doc = DocxDocument.new()
        ops = [
            {"op": "set_cell", "kwargs": {"block_id": "$missing", "row": 0, "col": 0, "text": "x"}},
        ]
        result = _run_batch(doc, ops)
        assert result["ok"] is False
        assert "undefined variable" in result["error"].lower()

    def test_batch_chained_add_then_fill(self):
        """Simulate the conventions-table workflow: add table, fill all cells."""
        from docx_toolkit.cli import _run_batch

        doc = DocxDocument.new()
        ops = [
            {"op": "add_table", "kwargs": {"rows": 3, "cols": 2}, "as": "tbl"},
            {"op": "set_cell", "kwargs": {"block_id": "$tbl", "row": 0, "col": 0, "text": "A"}},
            {"op": "set_cell", "kwargs": {"block_id": "$tbl", "row": 0, "col": 1, "text": "B"}},
            {"op": "add_row", "kwargs": {"block_id": "$tbl", "values": ["C", "D"]}},
        ]
        result = _run_batch(doc, ops)
        assert result["ok"] is True

        table = doc.get_table(0)
        assert table["cells"][0] == ["A", "B"]
        assert table["rows"] == 4  # 3 original + 1 added

    def test_batch_multiple_vars(self):
        """Multiple named variables in one batch."""
        from docx_toolkit.cli import _run_batch

        doc = DocxDocument.new()
        ops = [
            {"op": "add_paragraph", "kwargs": {"text": "First"}, "as": "p1"},
            {"op": "add_paragraph", "kwargs": {"text": "Second"}, "as": "p2"},
        ]
        result = _run_batch(doc, ops)
        assert result["ok"] is True

        blocks = doc.read()["blocks"]
        texts = [b["tx"] for b in blocks]
        assert texts == ["First", "Second"]


# =============================================================================
# 4. add_paragraphs — bulk paragraph insertion
# =============================================================================

class TestAddParagraphs:
    def test_add_paragraphs_basic(self):
        doc = DocxDocument.new()
        ids = doc.add_paragraphs([
            {"text": "Item 1"},
            {"text": "Item 2", "style": "Heading 2"},
            {"text": "Item 3"},
        ])
        assert len(ids) == 3

        blocks = doc.read()["blocks"]
        texts = [b["tx"] for b in blocks]
        assert texts == ["Item 1", "Item 2", "Item 3"]
        styles = [b["st"] for b in blocks]
        assert styles[1] == "Heading 2"

    def test_add_paragraphs_after_target(self):
        doc = DocxDocument.new()
        first = doc.add_paragraph("Keep this")
        ids = doc.add_paragraphs(
            [{"text": "Inserted A"}, {"text": "Inserted B"}],
            after=first,
        )
        assert len(ids) == 2

        blocks = doc.read()["blocks"]
        texts = [b["tx"] for b in blocks]
        # First paragraph stays at position 0, new ones inserted right after it
        assert "Keep this" in texts[0]

    def test_add_paragraphs_empty_list(self):
        doc = DocxDocument.new()
        ids = doc.add_paragraphs([])
        assert ids == []


# =============================================================================
# 5. Heading style detection — German and non-standard styles
# =============================================================================

class TestHeadingDetection:
    def test_english_heading_detected(self):
        """Test _heading_level helper directly with English style names."""
        doc = DocxDocument.new()
        assert doc._heading_level("Heading 1") == 1
        assert doc._heading_level("Heading 2") == 2
        assert doc._heading_level("Heading 9") == 9

    def test_german_heading_detected(self):
        """German style names like 'Überschrift1' should be detected."""
        doc = DocxDocument.new()
        # After extension, these should return proper levels
        # (Currently they return None — tests will fail until P1 is implemented)
        assert doc._heading_level("Überschrift1") == 1
        assert doc._heading_level("Abschnitt2") == 2

    def test_non_heading_returns_none(self):
        doc = DocxDocument.new()
        assert doc._heading_level("Normal") is None
        assert doc._heading_level(None) is None
        assert doc._heading_level("") is None


# =============================================================================
# 6. find command — locate blocks by text/style
# =============================================================================

class TestFind:
    def test_find_by_text(self):
        doc = DocxDocument.new()
        doc.add_paragraph("Hello world")
        doc.add_paragraph("Goodbye world")
        doc.add_paragraph("Hello again")

        matches = doc.find(text_pattern="Hello")
        assert len(matches["matches"]) == 2

    def test_find_by_style(self):
        doc = DocxDocument.new()
        doc.add_paragraph("Title", style="Heading 1")
        doc.add_paragraph("Body", style="Normal")

        matches = doc.find(style="Heading 1")
        assert len(matches["matches"]) == 1
        assert matches["matches"][0]["tx"] == "Title"

    def test_find_heading_only(self):
        doc = DocxDocument.new()
        doc.add_paragraph("H1", style="Heading 1")
        doc.add_paragraph("Normal text")
        doc.add_paragraph("H2", style="Heading 2")

        matches = doc.find(heading_only=True)
        assert len(matches["matches"]) == 2

    def test_find_no_matches(self):
        doc = DocxDocument.new()
        doc.add_paragraph("Hello")
        matches = doc.find(text_pattern="Zzzz")
        assert matches["matches"] == []


# =============================================================================
# 7. CLI integration tests for new commands
# =============================================================================

class TestCLINewCommands:
    """Test the CLI layer for new subcommands."""

    def _run(self, *args):
        from docx_toolkit.cli import run
        return run(list(args))

    def test_table_delete_cli(self, tmp_path):
        path = str(tmp_path / "doc.docx")
        self._run("new", path)
        self._run("para-add", path, "Before")
        added = self._run("table-add", path, "2", "2")
        tbl_id = added["id"]
        self._run("table-set-cell", path, str(tbl_id), "0", "0", "X")

        result = self._run("table-delete", path, str(tbl_id))
        assert "hash" in result

        read = self._run("read", path)
        types = [b["ty"] for b in read["blocks"]]
        assert types == ["p"]

    def test_delete_range_cli(self, tmp_path):
        path = str(tmp_path / "doc.docx")
        self._run("new", path)
        self._run("para-add", path, "Keep1")
        self._run("para-add", path, "Remove")
        self._run("para-add", path, "Keep2")

        result = self._run("delete-range", path, "1", "2")
        assert "hash" in result

        read = self._run("read", path)
        texts = [b["tx"] for b in read["blocks"]]
        assert texts == ["Keep1", "Keep2"]

    def test_find_cli(self, tmp_path):
        path = str(tmp_path / "doc.docx")
        self._run("new", path)
        self._run("para-add", path, "Hello world", "--style", "Heading 1")
        self._run("para-add", path, "Normal text")

        result = self._run("find", path, "--text", "Hello")
        assert len(result["matches"]) == 1
        assert result["matches"][0]["tx"] == "Hello world"

    def test_batch_with_var_refs_cli(self, tmp_path):
        """Test contextual batch via CLI with $var references."""
        path = str(tmp_path / "doc.docx")
        self._run("new", path)

        ops = json.dumps([
            {"op": "add_table", "kwargs": {"rows": 2, "cols": 2}, "as": "t1"},
            {"op": "set_cell", "kwargs": {"block_id": "$t1", "row": 0, "col": 0, "text": "Header"}},
        ])
        result = self._run("batch", path, ops)
        assert result["ok"] is True

        table = self._run("table-get", path, "0")
        assert table["cells"][0][0] == "Header"
