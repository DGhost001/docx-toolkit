import pytest

from docx_toolkit import DocxDocument


def test_outline_reports_heading_levels(sample_doc: DocxDocument):
    outline = sample_doc.outline()
    levels = [(o["lvl"], o["tx"]) for o in outline]
    assert (1, "Section One") in levels
    assert (1, "Section Two") in levels
    # Title style is not a numbered heading level.
    assert all(tx != "Title Text" for _lvl, tx in levels)


def test_content_hash_changes_after_mutation(blank_doc: DocxDocument):
    h1 = blank_doc.content_hash
    blank_doc.add_paragraph("Something new")
    h2 = blank_doc.content_hash
    assert h1 != h2


def test_content_hash_stable_without_mutation(blank_doc: DocxDocument):
    blank_doc.add_paragraph("Stable content")
    h1 = blank_doc.content_hash
    h2 = blank_doc.content_hash
    assert h1 == h2


def test_list_styles_includes_builtin_styles(blank_doc: DocxDocument):
    styles = blank_doc.list_styles()
    names = {s["name"] for s in styles}
    assert "Normal" in names
    assert "Heading 1" in names


def test_read_empty_document_has_no_blocks():
    doc = DocxDocument.new()
    assert doc.read()["blocks"] == []


def test_block_ids_are_unique_within_a_read(sample_doc: DocxDocument):
    blocks = sample_doc.read()["blocks"]
    ids = [b["id"] for b in blocks]
    assert len(ids) == len(set(ids))
