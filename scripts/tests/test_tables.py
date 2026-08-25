import pytest

from docx_toolkit import DocxDocument, StaleHashError


def test_add_table_returns_block_id_and_correct_dims(blank_doc: DocxDocument):
    block_id = blank_doc.add_table(rows=3, cols=2)
    table = blank_doc.get_table(block_id)
    assert table["rows"] == 3
    assert table["cols"] == 2
    assert len(table["cells"]) == 3
    assert all(len(row) == 2 for row in table["cells"])


def test_set_cell_updates_text(blank_doc: DocxDocument):
    block_id = blank_doc.add_table(rows=2, cols=2)
    blank_doc.set_cell(block_id, 0, 0, "Field")
    blank_doc.set_cell(block_id, 0, 1, "Value")

    table = blank_doc.get_table(block_id)
    assert table["cells"][0] == ["Field", "Value"]


def test_add_row_appends_and_returns_index(blank_doc: DocxDocument):
    block_id = blank_doc.add_table(rows=1, cols=2)
    blank_doc.set_cell(block_id, 0, 0, "Header A")
    blank_doc.set_cell(block_id, 0, 1, "Header B")

    row_idx = blank_doc.add_row(block_id, ["val1", "val2"])
    assert row_idx == 1

    table = blank_doc.get_table(block_id)
    assert table["rows"] == 2
    assert table["cells"][1] == ["val1", "val2"]


def test_delete_row_removes_correct_row(blank_doc: DocxDocument):
    block_id = blank_doc.add_table(rows=1, cols=1)
    blank_doc.set_cell(block_id, 0, 0, "row0")
    blank_doc.add_row(block_id, ["row1"])
    blank_doc.add_row(block_id, ["row2"])

    blank_doc.delete_row(block_id, 1)

    table = blank_doc.get_table(block_id)
    assert table["rows"] == 2
    assert table["cells"] == [["row0"], ["row2"]]


def test_table_in_body_order_alongside_paragraphs(blank_doc: DocxDocument):
    p1 = blank_doc.add_paragraph("Before table")
    t1 = blank_doc.add_table(rows=1, cols=1, after=p1)
    p2 = blank_doc.add_paragraph("After table")

    blocks = blank_doc.read()["blocks"]
    types = [b["ty"] for b in blocks]
    assert types == ["p", "tbl", "p"]


def test_set_cell_with_stale_hash_raises(blank_doc: DocxDocument):
    block_id = blank_doc.add_table(rows=1, cols=1)
    with pytest.raises(StaleHashError):
        blank_doc.set_cell(block_id, 0, 0, "x", expect_hash="stale")
