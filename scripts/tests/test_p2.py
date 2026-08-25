"""Tests for P2 docx-toolkit extensions.

Covers: list support (bullets/numbering), table column operations,
merge-paragraphs, split-paragraph.
"""

import json
import pytest

from docx_toolkit import DocxDocument


# =============================================================================
# 1. List support — bullet and numbered lists
# =============================================================================

class TestListSupport:
    def _get_para_xml(self, doc, block_id):
        """Helper to get paragraph XML element for a block_id."""
        from lxml import etree
        body = doc._document.element.body
        tag, el = doc._get_block(body, block_id)
        return el

    def test_add_paragraph_bullet_list(self):
        """Add a paragraph with bullet list formatting."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("First item", list_type="bullet")

        blocks = doc.read()["blocks"]
        assert len(blocks) == 1
        assert blocks[0]["tx"] == "First item"

        # Verify the paragraph has numbering properties in XML
        para_el = self._get_para_xml(doc, p1)
        num_pr = para_el.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr")
        assert num_pr is not None, "Bullet paragraph should have numPr element"

    def test_add_paragraph_numbered_list(self):
        """Add a paragraph with numbered list formatting."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Step 1", list_type="number")

        blocks = doc.read()["blocks"]
        assert len(blocks) == 1

        para_el = self._get_para_xml(doc, p1)
        num_pr = para_el.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr")
        assert num_pr is not None, "Numbered paragraph should have numPr element"

    def test_add_paragraph_no_list(self):
        """Add a regular paragraph without list formatting."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Regular text")

        para_el = self._get_para_xml(doc, p1)
        num_pr = para_el.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr")
        assert num_pr is None, "Regular paragraph should not have numPr element"

    def test_multiple_bullet_items(self):
        """Add multiple bullet items — they should share the same numbering instance."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Item 1", list_type="bullet")
        p2 = doc.add_paragraph("Item 2", list_type="bullet")
        p3 = doc.add_paragraph("Item 3", list_type="bullet")

        blocks = doc.read()["blocks"]
        assert len(blocks) == 3

        # All should have numPr elements
        for pid in [p1, p2, p3]:
            para_el = self._get_para_xml(doc, pid)
            num_pr = para_el.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr")
            assert num_pr is not None

    def test_mixed_list_and_regular(self):
        """Mix list items with regular paragraphs."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Heading", style="Heading 1")
        p2 = doc.add_paragraph("Bullet item", list_type="bullet")
        p3 = doc.add_paragraph("Regular text")
        p4 = doc.add_paragraph("Another bullet", list_type="bullet")

        blocks = doc.read()["blocks"]
        assert len(blocks) == 4

        # Check p2 and p4 have numPr, p1 and p3 don't
        for pid in [p2, p4]:
            para_el = self._get_para_xml(doc, pid)
            num_pr = para_el.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr")
            assert num_pr is not None

        for pid in [p1, p3]:
            para_el = self._get_para_xml(doc, pid)
            num_pr = para_el.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr")
            assert num_pr is None

    def test_bullet_paragraph_has_pStyle(self):
        """Bullet paragraphs MUST have <pStyle val='List Paragraph'> (or localized equivalent)
        in their pPr — without it, Word does not render the bullet marker.
        See: https://learn.microsoft.com/en-us/office/vba/api/word.style.listparagraph
        """
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Bullet item", list_type="bullet")
        para_el = self._get_para_xml(doc, p1)

        # Check that pStyle is present in the paragraph's pPr element
        ppr = para_el.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
        assert ppr is not None, "Paragraph should have a pPr element"

        psty = ppr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
        assert psty is not None, (
            "Bullet paragraph must have <pStyle> in pPr for Word to render bullets. "
            f"Got XML: {para_el.xml[:300]}"
        )

    def test_numbered_paragraph_has_pStyle(self):
        """Numbered paragraphs MUST also have <pStyle> set — same requirement as bullets."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Step 1", list_type="number")
        para_el = self._get_para_xml(doc, p1)

        ppr = para_el.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
        assert ppr is not None

        psty = ppr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
        assert psty is not None, (
            "Numbered paragraph must have <pStyle> in pPr. "
            f"Got XML: {para_el.xml[:300]}"
        )

    def test_explicit_style_not_overridden(self):
        """If the caller provides an explicit style, list_type should NOT override it."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Numbered heading", style="Heading 1", list_type="number")
        para_el = self._get_para_xml(doc, p1)

        psty = para_el.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr"
            "/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle"
        )
        if psty is not None:
            # Word normalises style IDs: "Heading 1" → "Heading1"
            assert psty.get(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
            ) in ("Heading 1", "Heading1"), (
                f"Explicit style should be preserved, got: {psty.attrib}"
            )

    def test_multiple_bullet_items_all_have_pStyle(self):
        """Every bullet item in a list must have <pStyle> set."""
        doc = DocxDocument.new()
        ids = []
        for i, label in enumerate(["Item 1", "Item 2", "Item 3"]):
            pid = doc.add_paragraph(label, list_type="bullet")
            ids.append(pid)

        for pid in ids:
            para_el = self._get_para_xml(doc, pid)
            psty = para_el.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr"
                "/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle"
            )
            assert psty is not None, f"Item '{label}' at block {pid} missing pStyle"

    def test_list_paragraph_has_rPr_lang_en_us(self):
        """Bullet/numbered paragraphs MUST have <rPr><lang w:val='en-US'/> in pPr.
        Without it, Word does not render bullets correctly — the numbering definition
        is present but bullet glyphs don't appear."""
        doc = DocxDocument.new()
        pid = doc.add_paragraph("Bullet item", list_type="bullet")
        para_el = self._get_para_xml(doc, pid)

        ppr = para_el.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr"
        )
        assert ppr is not None
        rpr = ppr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
        assert rpr is not None, "Paragraph missing <rPr> in pPr for list items"
        lang_el = rpr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lang"
        )
        assert lang_el is not None
        val = lang_el.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
        assert val == "en-US", f"Expected 'en-US', got '{val}'"

    def test_batch_paragraphs_have_rPr_lang(self):
        """Batch add_paragraphs with list_type must also set rPr.lang=en-US."""
        doc = DocxDocument.new()
        heading = doc.add_paragraph("Heading")
        ids = doc.add_paragraphs(
            [{"text": f"Item {i}", "list_type": "bullet"} for i in range(1, 4)],
            after=heading,
        )

        for pid in ids:
            para_el = self._get_para_xml(doc, pid)
            ppr = para_el.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr"
            )
            assert ppr is not None
            rpr = ppr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
            assert rpr is not None, f"Batch item {pid} missing <rPr>"
            lang_el = rpr.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lang"
            )
            assert lang_el is not None
            val = lang_el.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
            assert val == "en-US", f"Expected 'en-US' at block {pid}, got '{val}'"

    def test_add_paragraphs_batch_insert_positioning(self):
        """add_paragraphs with 'after' must chain each paragraph after the previous one,
        not append all to end of document.

        The critical invariant: if we add a block AFTER the batch call, the new block
        should come AFTER the batch items — proving they are properly chained at the
        anchor position, NOT appended at the very end of the document.

        Given: [A] [B]
        After add_paragraphs([C, D], after=B) → [A] [B] [C] [D]
        Then doc.add_paragraph("F")            → [A] [B] [C] [D] [F]
        NOT:                                   → [A] [B] [F] [C] [D]  (bug)
                                              → [A] [B] [C] [D] [F]  (correct)
        """
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Block A")
        p2 = doc.add_paragraph("Block B")

        new_ids = doc.add_paragraphs(
            [
                {"text": "Bullet C", "list_type": "bullet"},
                {"text": "Bullet D", "list_type": "bullet"},
            ],
            after=p2,
        )

        # Now add another block AFTER the batch — this is the key test
        p3 = doc.add_paragraph("Block F")

        blocks = doc.read()["blocks"]
        text_order = [b["tx"] for b in blocks]

        assert text_order == ["Block A", "Block B", "Bullet C", "Bullet D", "Block F"], (
            f"Batch items must be chained at anchor position, not appended to end. "
            f"If 'after' was ignored for items 2+, Block F would appear between batch items. "
            f"Got: {text_order}"
        )

    def test_add_paragraphs_batch_insert_positioning_with_non_adjacent_anchor(self):
        """When anchor is NOT the last block, subsequent additions must still chain after anchor.

        Given: [A] [B] [C] (existing)
        After add_paragraphs([D, E], after=B) → [A] [B] [D] [E] [C]
        """
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Block A")
        p2 = doc.add_paragraph("Block B")
        p3_anchor = doc.add_paragraph("Block C")

        new_ids = doc.add_paragraphs(
            [
                {"text": "Bullet D", "list_type": "bullet"},
                {"text": "Bullet E", "list_type": "bullet"},
            ],
            after=p2,
        )

        blocks = doc.read()["blocks"]
        text_order = [b["tx"] for b in blocks]

        # C should still be last (batch items chained between B and C)
        assert text_order[-1] == "Block C", (
            f"Anchor block 'C' must remain at its position. Batch items inserted after B, "
            f"before C. If batch items went to end, order would differ. Got: {text_order}"
        )

        # D and E should be right after B
        b_idx = next(i for i, t in enumerate(text_order) if t == "Block B")
        assert text_order[b_idx + 1] == "Bullet D"
        assert text_order[b_idx + 2] == "Bullet E"

    def test_add_paragraphs_batch_preserves_list_formatting(self):
        """add_paragraphs batch items with list_type should get proper OOXML list formatting."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Heading")
        new_ids = doc.add_paragraphs(
            [
                {"text": "Item 1", "list_type": "bullet"},
                {"text": "Item 2", "list_type": "bullet"},
            ],
            after=p1,
        )

        for pid in new_ids:
            para_el = self._get_para_xml(doc, pid)
            ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

            # Should have numPr
            nump = para_el.find(f".//{ns}numPr")
            assert nump is not None, f"Batch bullet at {pid} missing numPr"

            # Should have pStyle (List Paragraph)
            psty = para_el.find(f".//{ns}pStyle")
            assert psty is not None, (
                f"Batch bullet at {pid} missing pStyle (required for Word rendering). "
                f"XML: {para_el.xml[:300]}"
            )


# =============================================================================
# 2. Table column operations
# =============================================================================

class TestTableColumns:
    def test_add_column_to_table(self):
        """Add a new column to an existing table."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=2)
        
        # Fill original cells
        doc.set_cell(t1, 0, 0, "A")
        doc.set_cell(t1, 0, 1, "B")
        doc.set_cell(t1, 1, 0, "C")
        doc.set_cell(t1, 1, 1, "D")
        
        # Add column at index 1 (between col 0 and col 1)
        new_id = doc.add_column(t1, 1, ["X", "Y"])
        
        table = doc.get_table(t1)
        assert table["cols"] == 3
        assert table["cells"][0] == ["A", "X", "B"]
        assert table["cells"][1] == ["C", "Y", "D"]

    def test_add_column_at_beginning(self):
        """Add a column at the start of a table."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=2)
        doc.set_cell(t1, 0, 0, "A")
        doc.set_cell(t1, 0, 1, "B")
        doc.set_cell(t1, 1, 0, "C")
        doc.set_cell(t1, 1, 1, "D")
        
        doc.add_column(t1, 0, ["NEW_A", "NEW_C"])
        
        table = doc.get_table(t1)
        assert table["cols"] == 3
        assert table["cells"][0] == ["NEW_A", "A", "B"]
        assert table["cells"][1] == ["NEW_C", "C", "D"]

    def test_add_column_at_end(self):
        """Add a column at the end of a table."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=2)
        doc.set_cell(t1, 0, 0, "A")
        doc.set_cell(t1, 0, 1, "B")
        doc.set_cell(t1, 1, 0, "C")
        doc.set_cell(t1, 1, 1, "D")
        
        doc.add_column(t1, 2, ["E", "F"])
        
        table = doc.get_table(t1)
        assert table["cols"] == 3
        assert table["cells"][0] == ["A", "B", "E"]
        assert table["cells"][1] == ["C", "D", "F"]

    def test_delete_column_from_table(self):
        """Delete a column from an existing table."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=3)
        doc.set_cell(t1, 0, 0, "A")
        doc.set_cell(t1, 0, 1, "B")
        doc.set_cell(t1, 0, 2, "C")
        doc.set_cell(t1, 1, 0, "D")
        doc.set_cell(t1, 1, 1, "E")
        doc.set_cell(t1, 1, 2, "F")
        
        # Delete middle column (index 1)
        doc.delete_column(t1, 1)
        
        table = doc.get_table(t1)
        assert table["cols"] == 2
        assert table["cells"][0] == ["A", "C"]
        assert table["cells"][1] == ["D", "F"]

    def test_delete_first_column(self):
        """Delete the first column from a table."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=3)
        doc.set_cell(t1, 0, 0, "A")
        doc.set_cell(t1, 0, 1, "B")
        doc.set_cell(t1, 0, 2, "C")
        doc.set_cell(t1, 1, 0, "D")
        doc.set_cell(t1, 1, 1, "E")
        doc.set_cell(t1, 1, 2, "F")
        
        doc.delete_column(t1, 0)
        
        table = doc.get_table(t1)
        assert table["cols"] == 2
        assert table["cells"][0] == ["B", "C"]
        assert table["cells"][1] == ["E", "F"]

    def test_delete_last_column(self):
        """Delete the last column from a table."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=3)
        doc.set_cell(t1, 0, 0, "A")
        doc.set_cell(t1, 0, 1, "B")
        doc.set_cell(t1, 0, 2, "C")
        doc.set_cell(t1, 1, 0, "D")
        doc.set_cell(t1, 1, 1, "E")
        doc.set_cell(t1, 1, 2, "F")
        
        doc.delete_column(t1, 2)
        
        table = doc.get_table(t1)
        assert table["cols"] == 2
        assert table["cells"][0] == ["A", "B"]
        assert table["cells"][1] == ["D", "E"]

    def test_add_column_invalid_index_raises(self):
        """Adding a column at an invalid index should raise DocxError."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=2)
        
        with pytest.raises(Exception):  # DocxError or similar
            doc.add_column(t1, 5, ["X", "Y"])

    def test_delete_column_invalid_index_raises(self):
        """Deleting a column at an invalid index should raise DocxError."""
        doc = DocxDocument.new()
        t1 = doc.add_table(rows=2, cols=2)
        
        with pytest.raises(Exception):  # DocxError or similar
            doc.delete_column(t1, 5)


# =============================================================================
# 3. Merge and split paragraphs
# =============================================================================

class TestMergeSplit:
    def test_merge_two_paragraphs(self):
        """Merge two adjacent paragraphs into one."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Hello")
        p2 = doc.add_paragraph("World")

        # Merge p2 into p1 (p1 becomes "Hello World", p2 is removed)
        merged_id = doc.merge_paragraphs([p1, p2])

        blocks = doc.read()["blocks"]
        assert len(blocks) == 1
        assert blocks[0]["tx"] == "Hello World"

    def test_merge_three_paragraphs(self):
        """Merge three adjacent paragraphs."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Part 1")
        p2 = doc.add_paragraph("Part 2")
        p3 = doc.add_paragraph("Part 3")

        merged_id = doc.merge_paragraphs([p1, p2, p3])

        blocks = doc.read()["blocks"]
        assert len(blocks) == 1
        assert blocks[0]["tx"] == "Part 1 Part 2 Part 3"

    def test_split_paragraph(self):
        """Split a paragraph at a given character offset."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Hello World")

        # Split after position 5 (between "Hello" and "World")
        new_id = doc.split_paragraph(p1, 5)

        blocks = doc.read()["blocks"]
        assert len(blocks) == 2
        assert blocks[0]["tx"] == "Hello"
        assert blocks[1]["tx"] == " World"

    def test_split_paragraph_at_start(self):
        """Split at position 0 should create empty first paragraph."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Text")

        new_id = doc.split_paragraph(p1, 0)

        blocks = doc.read()["blocks"]
        assert len(blocks) == 2
        assert blocks[0]["tx"] == ""
        assert blocks[1]["tx"] == "Text"

    def test_split_paragraph_at_end(self):
        """Split at the end should create empty second paragraph."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("Text")

        new_id = doc.split_paragraph(p1, 4)

        blocks = doc.read()["blocks"]
        assert len(blocks) == 2
        assert blocks[0]["tx"] == "Text"
        assert blocks[1]["tx"] == ""

    def test_merge_and_split_roundtrip(self):
        """Merge then split should restore original structure."""
        doc = DocxDocument.new()
        p1 = doc.add_paragraph("First")
        p2 = doc.add_paragraph("Second")

        # Merge
        merged_id = doc.merge_paragraphs([p1, p2])
        blocks = doc.read()["blocks"]
        assert len(blocks) == 1
        assert blocks[0]["tx"] == "First Second"

        # Split back at the space (position 5)
        new_id = doc.split_paragraph(merged_id, 5)
        blocks = doc.read()["blocks"]
        assert len(blocks) == 2
        assert blocks[0]["tx"] == "First"
        assert blocks[1]["tx"] == " Second"
