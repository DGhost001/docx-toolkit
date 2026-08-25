import json
import subprocess
import sys

import pytest

from docx_toolkit import cli


def _run(*args: str) -> dict:
    return cli.run(list(args))


def test_new_creates_document(tmp_path):
    path = str(tmp_path / "doc.docx")
    result = _run("new", path)
    assert result == {"path": path}
    assert (tmp_path / "doc.docx").exists()


def test_read_on_fresh_document(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    result = _run("read", path)
    assert result["blocks"] == []
    assert "hash" in result


def test_para_add_then_read(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    added = _run("para-add", path, "Hello, world!", "--style", "Heading 1")
    assert added["id"] == 0
    assert "hash" in added

    read = _run("read", path)
    assert read["blocks"] == [{"id": 0, "ty": "p", "st": "Heading 1", "tx": "Hello, world!", "lvl": 1}]


def test_toc_reports_heading_blocks(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    _run("para-add", path, "Title Text", "--style", "Heading 1")
    _run("para-add", path, "Body text", "--style", "Normal")
    toc = _run("toc", path)
    assert toc["blocks"] == [{"id": 0, "lvl": 1, "tx": "Title Text"}]


def test_para_edit_and_delete_roundtrip(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    _run("para-add", path, "Original")
    edited = _run("para-edit", path, "0", "Edited")
    assert "hash" in edited
    assert _run("para-get", path, "0")["tx"] == "Edited"

    _run("para-delete", path, "0")
    assert _run("read", path)["blocks"] == []


def test_stale_hash_is_rejected(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    added = _run("para-add", path, "First")
    stale_hash = added["hash"]
    _run("para-add", path, "Second")

    result = _run("para-edit", path, "0", "Changed", "--expect-hash", stale_hash)
    assert result["type"] == "StaleHashError"


def test_table_workflow(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    added = _run("table-add", path, "2", "2", "--style", "Table Grid")
    block_id = added["id"]

    _run("table-set-cell", path, str(block_id), "0", "0", "Name")
    _run("table-add-row", path, str(block_id), json.dumps(["a", "b"]))
    table = _run("table-get", path, str(block_id))
    assert table["rows"] == 3
    assert table["cells"][0][0] == "Name"
    assert table["cells"][2] == ["a", "b"]

    _run("table-delete-row", path, str(block_id), "2")
    table_after = _run("table-get", path, str(block_id))
    assert table_after["rows"] == 2


def test_comment_workflow(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    _run("para-add", path, "Some text")
    added = _run("comment-add", path, "0", "0", "1", "first comment", "--author", "A", "--initials", "A")
    comment_id = added["id"]

    _run("comment-edit", path, str(comment_id), "--text", "updated comment", "--initials", "ZZ")
    comments = _run("comment-list", path)["comments"]
    comment = next(c for c in comments if c["id"] == comment_id)
    assert comment["text"] == "updated comment"
    assert comment["initials"] == "ZZ"
    assert comment["anchor_block"] == 0


def test_image_workflow(tmp_path, png_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    added = _run("image-add", path, str(png_path), "--width", "1.0", "--height", "1.0")
    assert "id" in added

    images = _run("image-list", path)["images"]
    assert len(images) == 1
    assert images[0]["block_id"] == added["id"]

    out_path = tmp_path / "extracted.png"
    extract_result = _run("image-extract", path, "0", str(out_path))
    assert extract_result["path"] == str(out_path)
    assert out_path.read_bytes() == png_path.read_bytes()


def test_header_and_footer_get(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    header = _run("header-get", path)
    footer = _run("footer-get", path)
    assert "blocks" in header
    assert "blocks" in footer


def test_section_count_and_styles(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    assert _run("section-count", path)["count"] == 1
    styles = _run("styles", path)["styles"]
    assert any(s["name"] == "Normal" for s in styles)


def test_batch_success_saves_once(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    ops = json.dumps(
        [
            {"op": "add_paragraph", "kwargs": {"text": "First", "style": "Heading 1"}},
            {"op": "add_paragraph", "kwargs": {"text": "Second"}},
        ]
    )
    result = _run("batch", path, ops)
    assert result["ok"] is True
    assert result["results"] == [0, 1]

    read = _run("read", path)
    texts = [b["tx"] for b in read["blocks"]]
    assert texts == ["First", "Second"]


def test_batch_failure_does_not_save_partial_state(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    ops = json.dumps(
        [
            {"op": "add_paragraph", "kwargs": {"text": "First"}},
            {"op": "edit_paragraph", "kwargs": {"block_id": 99, "text": "boom"}},
        ]
    )
    result = _run("batch", path, ops)
    assert result["ok"] is False
    assert result["failed_at"] == 1

    read = _run("read", path)
    assert read["blocks"] == []


def test_batch_rejects_unknown_op(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    ops = json.dumps([{"op": "delete_everything", "kwargs": {}}])
    result = _run("batch", path, ops)
    assert result["ok"] is False
    assert "unknown op" in result["error"]


def test_unknown_block_id_returns_structured_error(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    result = _run("para-get", path, "5")
    assert result["type"] == "DocxError"
    assert "error" in result


def test_open_missing_file_returns_structured_error(tmp_path):
    path = str(tmp_path / "does-not-exist.docx")
    result = _run("read", path)
    assert "error" in result


def test_main_prints_json_and_sets_exit_code(tmp_path, capsys):
    path = str(tmp_path / "doc.docx")
    exit_code = cli.main(["new", path])
    captured = capsys.readouterr()
    assert exit_code == 0
    assert json.loads(captured.out) == {"path": path}

    error_exit_code = cli.main(["read", str(tmp_path / "missing.docx")])
    error_output = json.loads(capsys.readouterr().out)
    assert error_exit_code == 1
    assert "error" in error_output


def test_invalid_arguments_return_structured_error(tmp_path, capsys):
    result = _run("para-add", str(tmp_path / "doc.docx"))
    assert result["type"] == "ArgumentError"
    capsys.readouterr()  # discard argparse's usage/error text on stderr


def test_batch_with_malformed_json_returns_structured_error(tmp_path):
    path = str(tmp_path / "doc.docx")
    _run("new", path)
    result = _run("batch", path, "not valid json")
    assert "error" in result
    assert result["type"] == "JSONDecodeError"


def test_cli_entry_point_via_subprocess(tmp_path):
    path = str(tmp_path / "doc.docx")
    proc = subprocess.run(
        [sys.executable, "-m", "docx_toolkit.cli", "new", path],
        capture_output=True,
        text=True,
        check=True,
    )
    assert json.loads(proc.stdout) == {"path": path}
