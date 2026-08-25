import pytest

from docx_toolkit import DocxDocument


@pytest.fixture
def tiny_png(png_path):
    return png_path


def test_add_image_returns_block_id(blank_doc: DocxDocument, tiny_png):
    block_id = blank_doc.add_image(tiny_png, width=1.0, height=1.0)
    assert isinstance(block_id, int)

    images = blank_doc.list_images()
    assert len(images) == 1
    assert images[0]["block_id"] == block_id


def test_add_image_reports_dimensions(blank_doc: DocxDocument, tiny_png):
    blank_doc.add_image(tiny_png, width=2.0, height=1.5)
    images = blank_doc.list_images()
    assert images[0]["width"] == pytest.approx(2.0, abs=0.01)
    assert images[0]["height"] == pytest.approx(1.5, abs=0.01)


def test_extract_image_round_trips_bytes(blank_doc: DocxDocument, tiny_png, tmp_path):
    blank_doc.add_image(tiny_png)
    images = blank_doc.list_images()
    image_id = images[0]["image_id"]

    out_path = tmp_path / "extracted.png"
    blank_doc.extract_image(image_id, out_path)

    assert out_path.read_bytes() == tiny_png.read_bytes()


def test_add_image_after_specific_block(blank_doc: DocxDocument, tiny_png):
    p1 = blank_doc.add_paragraph("Before image")
    img_block = blank_doc.add_image(tiny_png, after=p1)
    p2 = blank_doc.add_paragraph("After image")

    blocks = blank_doc.read()["blocks"]
    ids_in_order = [b["id"] for b in blocks]
    assert ids_in_order.index(p1) < ids_in_order.index(img_block) < ids_in_order.index(p2)
