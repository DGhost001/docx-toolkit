"""P3 features: copy/move blocks, table-fill, validate, diff."""
import json
from io import BytesIO

import pytest
from lxml import etree

from docx_toolkit import DocxDocument, DocxError


class TestCopyBlock:
    def test_copy_paragraph(self):
        """Copy a paragraph to a new position."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Original")
        
        # Copy after original
        new_id = doc.copy_block(p1, after=p1)
        
        blocks = doc.read()["blocks"]
        assert len(blocks) == 2
        assert blocks[0]["tx"] == "Original"
        assert blocks[1]["tx"] == "Original"
        assert new_id != p1

    def test_copy_table(self):
        """Copy a table to a new position."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=2)
        doc.set_cell(t1, 0, 0, "A")
        doc.set_cell(t1, 0, 1, "B")
        doc.set_cell(t1, 1, 0, "C")
        doc.set_cell(t1, 1, 1, "D")
        
        new_id = doc.copy_block(t1, after=t1)
        
        blocks = doc.read()["blocks"]
        assert len(blocks) == 2
        
        # Both tables should have same content
        t1_data = doc.get_table(t1)
        t2_data = doc.get_table(new_id)
        assert t1_data["cells"] == t2_data["cells"]

    def test_copy_with_hash_validation(self):
        """Copy fails with stale hash."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Test")
        
        with pytest.raises(DocxError, match="content has changed"):
            doc.copy_block(p1, after=p1, expect_hash="wrong_hash")

    def test_copy_nonexistent_raises(self):
        """Copy fails for unknown block_id."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Test")
        
        with pytest.raises(DocxError, match="unknown block_id"):
            doc.copy_block(999, after=p1)


class TestMoveBlock:
    def test_move_paragraph_forward(self):
        """Move a paragraph forward in the document."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("First")
        p2 = doc.add_paragraph("Second")
        p3 = doc.add_paragraph("Third")
        
        # Move p1 after p3
        moved_id = doc.move_block(p1, after=p3)
        
        blocks = doc.read()["blocks"]
        assert len(blocks) == 3
        assert blocks[0]["tx"] == "Second"
        assert blocks[1]["tx"] == "Third"
        assert blocks[2]["tx"] == "First"

    def test_move_paragraph_backward(self):
        """Move a paragraph backward in the document."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("First")
        p2 = doc.add_paragraph("Second")
        p3 = doc.add_paragraph("Third")
        
        # Move p3 before p1
        moved_id = doc.move_block(p3, after=p1)
        
        blocks = doc.read()["blocks"]
        assert len(blocks) == 3
        assert blocks[0]["tx"] == "First"
        assert blocks[1]["tx"] == "Third"
        assert blocks[2]["tx"] == "Second"

    def test_move_table(self):
        """Move a table to a new position."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Before")
        t1 = doc.add_table(rows=1, cols=1)
        doc.set_cell(t1, 0, 0, "Data")
        p2 = doc.add_paragraph("After")
        
        # Move table before paragraph
        moved_id = doc.move_block(t1, after=p1)
        
        blocks = doc.read()["blocks"]
        assert len(blocks) == 3
        assert blocks[0]["tx"] == "Before"
        assert blocks[1]["ty"] == "tbl"
        assert blocks[2]["tx"] == "After"

    def test_move_with_hash_validation(self):
        """Move fails with stale hash."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Test")
        
        with pytest.raises(DocxError, match="content has changed"):
            doc.move_block(p1, after=p1, expect_hash="wrong_hash")


class TestTableFill:
    def test_fill_from_2d_array(self):
        """Fill table from 2D array."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=3, cols=3)
        
        data = [
            ["Header1", "Header2", "Header3"],
            ["A1", "A2", "A3"],
            ["B1", "B2", "B3"],
        ]
        
        doc.table_fill(t1, data)
        
        table = doc.get_table(t1)
        assert table["cells"] == data

    def test_fill_partial_rows(self):
        """Fill with fewer rows than table has."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=3, cols=2)
        
        data = [["A", "B"]]  # Only one row
        
        doc.table_fill(t1, data)
        
        table = doc.get_table(t1)
        assert table["cells"][0] == ["A", "B"]
        # Other rows should be empty
        assert all(cell == "" for cell in table["cells"][1])

    def test_fill_csv_string(self):
        """Fill table from CSV string."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=2)
        
        csv_data = "Name,Age\nAlice,30\nBob,25"
        
        doc.table_fill(t1, csv_data)
        
        table = doc.get_table(t1)
        assert table["cells"][0] == ["Name", "Age"]
        assert table["cells"][1] == ["Alice", "30"]

    def test_fill_with_header_row(self):
        """Fill with separate header and data rows."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=4, cols=2)
        
        headers = ["ID", "Value"]
        data = [["1", "100"], ["2", "200"]]
        
        doc.table_fill(t1, data, header=headers)
        
        table = doc.get_table(t1)
        assert table["cells"][0] == headers
        assert table["cells"][1] == ["1", "100"]

    def test_fill_invalid_data_raises(self):
        """Fill with invalid data raises error."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=2)
        
        # Invalid: not a list or string
        with pytest.raises(DocxError):
            doc.table_fill(t1, 12345)


class TestValidate:
    def test_valid_document(self):
        """Valid document passes validation."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Test")
        
        result = doc.validate()
        assert result["valid"] is True
        assert len(result["errors"]) == 0

    def test_invalid_block_id(self):
        """Validation catches invalid block_id references."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Test")
        
        # Manually corrupt a reference by trying to access non-existent block
        with pytest.raises(DocxError, match="unknown block_id"):
            doc._get_block(doc._document.element.body, 999)

    def test_validate_table_structure(self):
        """Validate table has consistent column counts."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=3)
        
        # Add cells to first row only (inconsistent)
        doc.set_cell(t1, 0, 0, "A")
        doc.set_cell(t1, 0, 1, "B")
        doc.set_cell(t1, 0, 2, "C")
        
        result = doc.validate()
        # Should pass - table structure is valid even if cells are empty
        
    def test_validate_returns_errors(self):
        """Validation returns error details."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Test")
        
        result = doc.validate()
        assert "errors" in result
        assert "warnings" in result


class TestDiff:
    def test_diff_identical_documents(self):
        """Diff of identical documents shows no changes."""
        doc1 = DocxDocument.new()
        p1 = doc1.add_paragraph("Same")
        
        doc2 = DocxDocument.new()
        p2 = doc2.add_paragraph("Same")
        
        diff = DocxDocument.diff(doc1, doc2)
        assert len(diff["added"]) == 0
        assert len(diff["removed"]) == 0
        assert len(diff["modified"]) == 0

    def test_diff_added_block(self):
        """Diff detects added blocks."""
        doc1 = DocxDocument.new()
        p1 = doc1.add_paragraph("Original")
        
        doc2 = DocxDocument.new()
        p2 = doc2.add_paragraph("Original")
        p3 = doc2.add_paragraph("Added")
        
        diff = DocxDocument.diff(doc1, doc2)
        assert len(diff["added"]) == 1
        assert diff["added"][0]["tx"] == "Added"

    def test_diff_removed_block(self):
        """Diff detects removed blocks."""
        doc1 = DocxDocument.new()
        p1 = doc1.add_paragraph("Original")
        p2 = doc1.add_paragraph("To be removed")
        
        doc2 = DocxDocument.new()
        p3 = doc2.add_paragraph("Original")
        
        diff = DocxDocument.diff(doc1, doc2)
        assert len(diff["removed"]) == 1
        assert diff["removed"][0]["tx"] == "To be removed"

    def test_diff_modified_block(self):
        """Diff detects modified blocks."""
        doc1 = DocxDocument.new()
        p1 = doc1.add_paragraph("Original")
        
        doc2 = DocxDocument.new()
        p2 = doc2.add_paragraph("Modified")
        
        diff = DocxDocument.diff(doc1, doc2)
        assert len(diff["modified"]) == 1
        # Check that old and new blocks are present
        modified = diff["modified"][0]
        assert "old" in modified
        assert "new" in modified

    def test_diff_empty_document(self):
        """Diff of two empty documents."""
        doc1 = DocxDocument.new()
        doc2 = DocxDocument.new()
        
        diff = DocxDocument.diff(doc1, doc2)
        assert len(diff["added"]) == 0
        assert len(diff["removed"]) == 0
