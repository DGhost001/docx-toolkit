from docx_toolkit import DocxDocument


def test_get_header_returns_block_structure(blank_doc: DocxDocument):
    blank_doc.add_paragraph("Header text", container="header", section=0)
    header = blank_doc.get_header(section=0)
    texts = [b["tx"] for b in header["blocks"] if b["ty"] == "p"]
    assert "Header text" in texts


def test_get_footer_returns_block_structure(blank_doc: DocxDocument):
    blank_doc.add_paragraph("Footer text", container="footer", section=0)
    footer = blank_doc.get_footer(section=0)
    texts = [b["tx"] for b in footer["blocks"] if b["ty"] == "p"]
    assert "Footer text" in texts


def test_header_and_body_are_independent_containers(blank_doc: DocxDocument):
    blank_doc.add_paragraph("Body paragraph", container="body")
    blank_doc.add_paragraph("Header paragraph", container="header", section=0)

    body_texts = [b["tx"] for b in blank_doc.read(container="body")["blocks"] if b["ty"] == "p"]
    header_texts = [b["tx"] for b in blank_doc.get_header(0)["blocks"] if b["ty"] == "p"]

    assert "Body paragraph" in body_texts
    assert "Header paragraph" not in body_texts
    assert "Header paragraph" in header_texts


def test_edit_paragraph_in_footer_container(blank_doc: DocxDocument):
    block_id = blank_doc.add_paragraph("Original footer", container="footer", section=0)
    blank_doc.edit_paragraph(block_id, "Updated footer", container="footer", section=0)

    footer = blank_doc.get_footer(0)
    para = next(b for b in footer["blocks"] if b["id"] == block_id)
    assert para["tx"] == "Updated footer"


def test_section_count_at_least_one(blank_doc: DocxDocument):
    assert blank_doc.section_count() >= 1
