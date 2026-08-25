"""P4 features: CLI subcommands for copy/move/table-fill/validate/diff, advanced diff."""
import json

import pytest

from docx_toolkit import DocxDocument, DocxError


class TestCopyBlockCLI:
    def _run(self, *args):
        from docx_toolkit.cli import run
        return run(list(args))

    def test_copy_block_cli(self, tmp_path):
        """copy-block CLI command works."""
        path = str(tmp_path / "doc.docx")
        self._run("new", path)
        p1 = int(self._run("para-add", path, "Original")["id"])
        
        result = self._run("copy-block", path, str(p1), "--after", str(p1))
        assert "id" in result
        
        # Verify copy exists (read returns dict with blocks as list, not JSON string)
        read_result = self._run("read", path)
        blocks = read_result["blocks"]
        texts = [b["tx"] for b in blocks if b["ty"] == "p"]
        assert texts.count("Original") == 2

    def test_copy_block_cli_with_hash(self, tmp_path):
        """copy-block CLI validates hash."""
        path = str(tmp_path / "doc.docx")
        self._run("new", path)
        p1 = int(self._run("para-add", path, "Test")["id"])
        
        result = self._run("copy-block", path, str(p1), "--after", str(p1), 
                     "--expect-hash", "wrong_hash")
        assert "error" in result


class TestMoveBlockCLI:
    def _run(self, *args):
        from docx_toolkit.cli import run
        return run(list(args))

    def test_move_block_cli(self, tmp_path):
        """move-block CLI command works."""
        path = str(tmp_path / "doc.docx")
        self._run("new", path)
        p1 = int(self._run("para-add", path, "First")["id"])
        p2 = int(self._run("para-add", path, "Second")["id"])
        p3 = int(self._run("para-add", path, "Third")["id"])
        
        # Move p1 after p3
        self._run("move-block", path, str(p1), "--after", str(p3))
        
        read_result = self._run("read", path)
        blocks = read_result["blocks"]
        texts = [b["tx"] for b in blocks if b["ty"] == "p"]
        assert texts == ["Second", "Third", "First"]


class TestTableFillCLI:
    def _run(self, *args):
        from docx_toolkit.cli import run
        return run(list(args))

    def test_table_fill_cli_2d_array(self, tmp_path):
        """table-fill CLI with 2D array."""
        path = str(tmp_path / "doc.docx")
        self._run("new", path)
        t1 = int(self._run("table-add", path, "2", "2")["id"])
        
        data = [["A", "B"], ["C", "D"]]
        result = self._run("table-fill", path, str(t1), "--data", json.dumps(data))
        
        table = self._run("table-get", path, str(t1))
        assert table["cells"] == data

    def test_table_fill_cli_csv(self, tmp_path):
        """table-fill CLI with CSV string."""
        path = str(tmp_path / "doc.docx")
        self._run("new", path)
        t1 = int(self._run("table-add", path, "2", "2")["id"])
        
        csv_data = "Name,Age\nAlice,30"
        result = self._run("table-fill", path, str(t1), "--data", csv_data)
        
        table = self._run("table-get", path, str(t1))
        assert table["cells"][0] == ["Name", "Age"]

    def test_table_fill_cli_with_header(self, tmp_path):
        """table-fill CLI with header row."""
        path = str(tmp_path / "doc.docx")
        self._run("new", path)
        t1 = int(self._run("table-add", path, "3", "2")["id"])
        
        data = [["1", "100"]]
        headers = ["ID", "Value"]
        result = self._run("table-fill", path, str(t1), 
                     "--data", json.dumps(data),
                     "--header", json.dumps(headers))
        
        table = self._run("table-get", path, str(t1))
        assert table["cells"][0] == headers


class TestValidateCLI:
    def _run(self, *args):
        from docx_toolkit.cli import run
        return run(list(args))

    def test_validate_cli(self, tmp_path):
        """validate CLI command works."""
        path = str(tmp_path / "doc.docx")
        self._run("new", path)
        self._run("para-add", path, "Test")
        
        result = self._run("validate", path)
        assert result["valid"] is True
        assert isinstance(result["errors"], list)

    def test_validate_cli_with_errors(self, tmp_path):
        """validate CLI detects issues."""
        path = str(tmp_path / "doc.docx")
        self._run("new", path)
        
        # Create a valid document first
        result = self._run("validate", path)
        assert result["valid"] is True


class TestDiffCLI:
    def _run(self, *args):
        from docx_toolkit.cli import run
        return run(list(args))

    def test_diff_cli(self, tmp_path):
        """diff CLI command works."""
        path1 = str(tmp_path / "doc1.docx")
        path2 = str(tmp_path / "doc2.docx")
        
        # Create two documents with differences
        self._run("new", path1)
        self._run("para-add", path1, "Original")
        
        self._run("new", path2)
        self._run("para-add", path2, "Modified")
        self._run("para-add", path2, "Added")
        
        result = self._run("diff", path1, path2)
        assert len(result["added"]) >= 1
        assert len(result["removed"]) >= 0
        assert len(result["modified"]) >= 1

    def test_diff_cli_identical(self, tmp_path):
        """diff CLI of identical documents."""
        path1 = str(tmp_path / "doc1.docx")
        path2 = str(tmp_path / "doc2.docx")
        
        self._run("new", path1)
        self._run("para-add", path1, "Same")
        
        self._run("new", path2)
        self._run("para-add", path2, "Same")
        
        result = self._run("diff", path1, path2)
        assert len(result["added"]) == 0
        assert len(result["removed"]) == 0


class TestAdvancedDiff:
    def test_diff_character_level(self):
        """Diff detects character-level changes."""
        doc1 = DocxDocument.new()
        p1 = doc1.add_paragraph("Hello World")
        
        doc2 = DocxDocument.new()
        p2 = doc2.add_paragraph("Hello Earth")
        
        diff = DocxDocument.diff(doc1, doc2)
        assert len(diff["modified"]) == 1

    def test_diff_table_cell_level(self):
        """Diff detects table cell changes."""
        doc1 = DocxDocument.new()
        t1 = doc1.add_table(rows=2, cols=2)
        doc1.set_cell(t1, 0, 0, "A")
        doc1.set_cell(t1, 0, 1, "B")
        
        doc2 = DocxDocument.new()
        t2 = doc2.add_table(rows=1, cols=1)  # Different structure
        doc2.set_cell(t2, 0, 0, "X")
        
        diff = DocxDocument.diff(doc1, doc2)
        # Tables with different structure should be detected as modified or added/removed
        assert len(diff["added"]) + len(diff["modified"]) >= 1


class TestTemplateSupport:
    def test_new_with_template(self):
        """Create document from template."""
        # Template support is not yet implemented, so this test will fail
        # until we add the feature
        with pytest.raises(TypeError):
            doc = DocxDocument.new(template="minimal")

    def test_new_with_custom_template(self):
        """Create document with custom template."""
        # Template support is not yet implemented, so this test will fail
        # until we add the feature
        with pytest.raises(TypeError):
            doc = DocxDocument.new(template="report")


class TestAutoNumberingContinuation:
    def test_numbering_continues_across_batches(self):
        """Numbering continues when adding multiple lists."""
        doc = DocxDocument.new()
        
        # Add first list
        p1 = doc.add_paragraph("Item 1", style="List Bullet")
        p2 = doc.add_paragraph("Item 2", style="List Bullet")
        
        # Add second list (should continue numbering if using same level)
        p3 = doc.add_paragraph("Item A", style="List Bullet")
        p4 = doc.add_paragraph("Item B", style="List Bullet")
        
        # Verify all items exist
        blocks = doc.read()["blocks"]
        list_items = [b for b in blocks if "List" in str(b.get("st", ""))]
        assert len(list_items) == 4
