import struct
import zlib

import pytest

from docx_toolkit import DocxDocument


def _make_1x1_png(rgb: tuple[int, int, int] = (255, 0, 0)) -> bytes:
    """Build a minimal valid 1x1 PNG in-process (no external fixture/dep needed)."""

    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
    raw_scanline = b"\x00" + bytes(rgb)
    idat = chunk(b"IDAT", zlib.compress(raw_scanline))
    iend = chunk(b"IEND", b"")
    return sig + ihdr + idat + iend


@pytest.fixture
def blank_doc() -> DocxDocument:
    return DocxDocument.new()


@pytest.fixture
def doc_path(tmp_path):
    return tmp_path / "doc.docx"


@pytest.fixture
def png_path(tmp_path):
    path = tmp_path / "tiny.png"
    path.write_bytes(_make_1x1_png())
    return path


@pytest.fixture
def sample_doc(blank_doc: DocxDocument) -> DocxDocument:
    """A small doc with headings, body paragraphs and a table, for read/toc tests."""
    blank_doc.add_paragraph("Title Text", style="Title")
    blank_doc.add_paragraph("Section One", style="Heading 1")
    blank_doc.add_paragraph("First body paragraph.", style="Normal")
    blank_doc.add_paragraph("Section Two", style="Heading 1")
    blank_doc.add_paragraph("Second body paragraph.", style="Normal")
    blank_doc.add_table(rows=2, cols=2)
    return blank_doc
