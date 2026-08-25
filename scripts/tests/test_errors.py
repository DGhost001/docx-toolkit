import pytest

from docx_toolkit import DocxDocument, DocxError


def test_unknown_container_raises_on_read(blank_doc: DocxDocument):
    with pytest.raises(DocxError):
        blank_doc.read(container="bogus")


def test_unknown_container_raises_on_paragraph_parent_resolution(blank_doc: DocxDocument):
    with pytest.raises(DocxError):
        blank_doc.add_paragraph("x", container="bogus")


def test_get_block_wrong_type_raises(blank_doc: DocxDocument):
    block_id = blank_doc.add_paragraph("A paragraph, not a table")
    with pytest.raises(DocxError):
        blank_doc.get_table(block_id)


def test_edit_paragraph_run_index_out_of_range_raises(blank_doc: DocxDocument):
    block_id = blank_doc.add_paragraph("only one run")
    with pytest.raises(DocxError):
        blank_doc.edit_paragraph(block_id, "x", run=5)


def test_add_comment_with_empty_run_range_raises(blank_doc: DocxDocument):
    block_id = blank_doc.add_paragraph("Some text")
    with pytest.raises(DocxError):
        blank_doc.add_comment(block_id, run_start=3, run_end=3, text="x")


def test_edit_comment_unknown_id_raises(blank_doc: DocxDocument):
    with pytest.raises(DocxError):
        blank_doc.edit_comment(999, text="x")


def test_edit_paragraph_on_empty_paragraph_adds_run(blank_doc: DocxDocument):
    block_id = blank_doc.add_paragraph("")
    blank_doc.edit_paragraph(block_id, "now has text")
    assert blank_doc.get_paragraph(block_id)["tx"] == "now has text"


def test_edit_comment_updates_initials(blank_doc: DocxDocument):
    block_id = blank_doc.add_paragraph("Some text")
    comment_id = blank_doc.add_comment(block_id, 0, 1, "a comment", author="A", initials="A")
    blank_doc.edit_comment(comment_id, initials="ZZ")
    updated = next(c for c in blank_doc.list_comments() if c["id"] == comment_id)
    assert updated["initials"] == "ZZ"


def test_delete_range_empty_raises(blank_doc: DocxDocument):
    blank_doc.add_paragraph("A", style="Normal")
    blank_doc.add_paragraph("B", style="Normal")
    blank_doc.add_paragraph("C", style="Normal")
    with pytest.raises(DocxError, match="empty range"):
        blank_doc.delete_range(1, 1)


def test_delete_range_reversed_raises(blank_doc: DocxDocument):
    blank_doc.add_paragraph("A", style="Normal")
    blank_doc.add_paragraph("B", style="Normal")
    with pytest.raises(DocxError, match="empty range"):
        blank_doc.delete_range(3, 1)


def test_delete_range_start_past_end_raises(blank_doc: DocxDocument):
    blank_doc.add_paragraph("A", style="Normal")
    with pytest.raises(DocxError, match="out of bounds"):
        blank_doc.delete_range(5, 6)


def test_add_table_in_footer_container_applies_style_and_dims(blank_doc: DocxDocument):
    block_id = blank_doc.add_table(rows=2, cols=2, style="Table Grid", container="footer", section=0)
    footer = blank_doc.get_footer(0)
    table_block = next(b for b in footer["blocks"] if b["id"] == block_id)
    assert table_block["ty"] == "tbl"
    assert table_block["rows"] == 2
    assert table_block["cols"] == 2
    assert table_block["st"] == "Table Grid"
