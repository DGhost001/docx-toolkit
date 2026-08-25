"""Agent-facing CLI for docx-toolkit.

Every invocation is stateless: open the file, perform one operation (or a
batch of them), save if anything mutated, and print a single compact JSON
object to stdout. There is no human-readable output mode -- this tool is
meant to be driven by other programs/agents, not read directly.

On success, stdout is a JSON object specific to the operation (always
including "hash" for mutating operations, so the caller can chain edits with
stale-hash protection). On failure, stdout is ``{"error": "...", "type":
"..."}`` and the process exits with status 1.
"""

from __future__ import annotations

import argparse
import json
import sys
from typing import Any

from docx.opc.exceptions import PackageNotFoundError

from .core import DocxDocument, DocxError


def _json_value(raw: str) -> Any:
    return json.loads(raw)


def _add_container_args(parser: argparse.ArgumentParser, default_container: str = "body") -> None:
    parser.add_argument("--container", choices=["body", "header", "footer"], default=default_container)
    parser.add_argument("--section", type=int, default=0)


def _add_expect_hash_arg(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--expect-hash", dest="expect_hash", default=None)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="docxtool", description="Agent-facing docx read/write tool.")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("new")
    p.add_argument("path")

    p = sub.add_parser("read")
    p.add_argument("path")
    _add_container_args(p)

    p = sub.add_parser("toc")
    p.add_argument("path")

    p = sub.add_parser("styles")
    p.add_argument("path")

    p = sub.add_parser("section-count")
    p.add_argument("path")

    p = sub.add_parser("find")
    p.add_argument("path")
    p.add_argument("--text", default=None)
    p.add_argument("--style", default=None)
    p.add_argument("--heading-only", action="store_true", dest="heading_only")
    _add_container_args(p)

    p = sub.add_parser("para-get")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    _add_container_args(p)

    p = sub.add_parser("para-add")
    p.add_argument("path")
    p.add_argument("text")
    p.add_argument("--style", default=None)
    p.add_argument("--after", type=int, default=None)
    _add_container_args(p)
    _add_expect_hash_arg(p)

    p = sub.add_parser("para-edit")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    p.add_argument("text")
    p.add_argument("--run", type=int, default=None)
    _add_container_args(p)
    _add_expect_hash_arg(p)

    p = sub.add_parser("para-delete")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    _add_container_args(p)
    _add_expect_hash_arg(p)

    p = sub.add_parser("table-get")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    _add_container_args(p)

    p = sub.add_parser("table-add")
    p.add_argument("path")
    p.add_argument("rows", type=int)
    p.add_argument("cols", type=int)
    p.add_argument("--style", default=None)
    p.add_argument("--after", type=int, default=None)
    _add_container_args(p)
    _add_expect_hash_arg(p)

    p = sub.add_parser("table-set-cell")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    p.add_argument("row", type=int)
    p.add_argument("col", type=int)
    p.add_argument("text")
    _add_container_args(p)
    _add_expect_hash_arg(p)

    p = sub.add_parser("table-add-row")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    p.add_argument("values", type=_json_value, help='JSON array, e.g. \'["a","b"]\'')
    _add_container_args(p)
    _add_expect_hash_arg(p)

    p = sub.add_parser("table-delete-row")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    p.add_argument("row", type=int)
    _add_container_args(p)
    _add_expect_hash_arg(p)

    p = sub.add_parser("table-delete")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    _add_container_args(p)
    _add_expect_hash_arg(p)

    p = sub.add_parser("delete-range")
    p.add_argument("path")
    p.add_argument("start_id", type=int)
    p.add_argument("end_id", type=int)
    _add_container_args(p)
    _add_expect_hash_arg(p)

    p = sub.add_parser("comment-list")
    p.add_argument("path")

    p = sub.add_parser("comment-add")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    p.add_argument("run_start", type=int)
    p.add_argument("run_end", type=int)
    p.add_argument("text")
    p.add_argument("--author", default="")
    p.add_argument("--initials", default="")
    _add_expect_hash_arg(p)

    p = sub.add_parser("comment-edit")
    p.add_argument("path")
    p.add_argument("comment_id", type=int)
    p.add_argument("--text", default=None)
    p.add_argument("--author", default=None)
    p.add_argument("--initials", default=None)

    p = sub.add_parser("image-add")
    p.add_argument("path")
    p.add_argument("src")
    p.add_argument("--after", type=int, default=None)
    p.add_argument("--width", type=float, default=None)
    p.add_argument("--height", type=float, default=None)
    _add_container_args(p)
    _add_expect_hash_arg(p)

    p = sub.add_parser("image-list")
    p.add_argument("path")

    p = sub.add_parser("image-extract")
    p.add_argument("path")
    p.add_argument("image_id", type=int)
    p.add_argument("out_path")

    p = sub.add_parser("header-get")
    p.add_argument("path")
    p.add_argument("--section", type=int, default=0)

    p = sub.add_parser("footer-get")
    p.add_argument("path")
    p.add_argument("--section", type=int, default=0)

    # copy-block
    p = sub.add_parser("copy-block")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    p.add_argument("--after", type=int, default=None)
    p.add_argument("--before", type=int, default=None)
    _add_container_args(p)
    _add_expect_hash_arg(p)

    # move-block
    p = sub.add_parser("move-block")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    p.add_argument("--after", type=int, default=None)
    p.add_argument("--before", type=int, default=None)
    _add_container_args(p)
    _add_expect_hash_arg(p)

    # table-fill
    p = sub.add_parser("table-fill")
    p.add_argument("path")
    p.add_argument("block_id", type=int)
    p.add_argument("--data", required=True, help="JSON array or CSV string")
    p.add_argument("--header", default=None, help="JSON array of header values")
    _add_container_args(p)
    _add_expect_hash_arg(p)

    # validate
    p = sub.add_parser("validate")
    p.add_argument("path")

    # diff
    p = sub.add_parser("diff")
    p.add_argument("path1")
    p.add_argument("path2")

    p = sub.add_parser("batch")
    p.add_argument("path")
    p.add_argument(
        "ops",
        nargs="?",
        default="-",
        help='JSON array of {"op": "<method>", "kwargs": {...}}; "-" (default) reads from stdin',
    )

    return parser


# Operations that mutate the document and must be persisted with .save().
_MUTATING_COMMANDS = {
    "para-add",
    "para-edit",
    "para-delete",
    "table-add",
    "table-set-cell",
    "table-add-row",
    "table-delete-row",
    "table-delete",
    "delete-range",
    "comment-add",
    "comment-edit",
    "image-add",
    "copy-block",
    "move-block",
    "table-fill",
}


def _dispatch(doc: DocxDocument, command: str, args: argparse.Namespace) -> dict[str, Any]:
    if command == "read":
        return doc.read(container=args.container, section=args.section)
    if command == "toc":
        return {"blocks": doc.outline()}
    if command == "styles":
        return {"styles": doc.list_styles()}
    if command == "section-count":
        return {"count": doc.section_count()}
    if command == "find":
        return doc.find(
            text_pattern=args.text,
            style=args.style,
            heading_only=args.heading_only,
            container=args.container,
            section=args.section,
        )
    if command == "para-get":
        return doc.get_paragraph(args.block_id, container=args.container, section=args.section)
    if command == "para-add":
        new_id = doc.add_paragraph(
            args.text,
            style=args.style,
            after=args.after,
            container=args.container,
            section=args.section,
            expect_hash=args.expect_hash,
        )
        return {"id": new_id, "hash": doc.content_hash}
    if command == "para-edit":
        doc.edit_paragraph(
            args.block_id,
            args.text,
            run=args.run,
            container=args.container,
            section=args.section,
            expect_hash=args.expect_hash,
        )
        return {"hash": doc.content_hash}
    if command == "para-delete":
        doc.delete_paragraph(
            args.block_id, container=args.container, section=args.section, expect_hash=args.expect_hash
        )
        return {"hash": doc.content_hash}
    if command == "table-get":
        return doc.get_table(args.block_id, container=args.container, section=args.section)
    if command == "table-add":
        new_id = doc.add_table(
            args.rows,
            args.cols,
            style=args.style,
            after=args.after,
            container=args.container,
            section=args.section,
            expect_hash=args.expect_hash,
        )
        return {"id": new_id, "hash": doc.content_hash}
    if command == "table-set-cell":
        doc.set_cell(
            args.block_id,
            args.row,
            args.col,
            args.text,
            container=args.container,
            section=args.section,
            expect_hash=args.expect_hash,
        )
        return {"hash": doc.content_hash}
    if command == "table-add-row":
        row_idx = doc.add_row(
            args.block_id,
            args.values,
            container=args.container,
            section=args.section,
            expect_hash=args.expect_hash,
        )
        return {"row": row_idx, "hash": doc.content_hash}
    if command == "table-delete-row":
        doc.delete_row(
            args.block_id, args.row, container=args.container, section=args.section, expect_hash=args.expect_hash
        )
        return {"hash": doc.content_hash}
    if command == "table-delete":
        doc.delete_table(
            args.block_id, container=args.container, section=args.section, expect_hash=args.expect_hash
        )
        return {"hash": doc.content_hash}
    if command == "delete-range":
        doc.delete_range(
            args.start_id,
            args.end_id,
            container=args.container,
            section=args.section,
            expect_hash=args.expect_hash,
        )
        return {"hash": doc.content_hash}
    if command == "comment-list":
        return {"comments": doc.list_comments()}
    if command == "comment-add":
        comment_id = doc.add_comment(
            args.block_id,
            args.run_start,
            args.run_end,
            args.text,
            author=args.author,
            initials=args.initials,
            expect_hash=args.expect_hash,
        )
        return {"id": comment_id, "hash": doc.content_hash}
    if command == "comment-edit":
        doc.edit_comment(args.comment_id, text=args.text, author=args.author, initials=args.initials)
        return {"hash": doc.content_hash}
    if command == "image-add":
        block_id = doc.add_image(
            args.src,
            after=args.after,
            width=args.width,
            height=args.height,
            container=args.container,
            section=args.section,
            expect_hash=args.expect_hash,
        )
        return {"id": block_id, "hash": doc.content_hash}
    if command == "image-list":
        return {"images": doc.list_images()}
    if command == "image-extract":
        doc.extract_image(args.image_id, args.out_path)
        return {"path": args.out_path}
    if command == "header-get":
        return doc.get_header(section=args.section)
    if command == "footer-get":
        return doc.get_footer(section=args.section)
    if command == "copy-block":
        block_id = doc.copy_block(
            args.block_id,
            after=args.after,
            before=args.before,
            container=args.container,
            section=args.section,
            expect_hash=args.expect_hash,
        )
        return {"id": block_id, "hash": doc.content_hash}
    if command == "move-block":
        block_id = doc.move_block(
            args.block_id,
            after=args.after,
            before=args.before,
            container=args.container,
            section=args.section,
            expect_hash=args.expect_hash,
        )
        return {"id": block_id, "hash": doc.content_hash}
    if command == "table-fill":
        import json as _json
        # Try to parse as JSON, fall back to CSV string
        try:
            data = _json.loads(args.data)
        except (ValueError, TypeError):
            data = args.data  # Keep as CSV string
        doc.table_fill(
            args.block_id,
            data,
            header=_json.loads(args.header) if args.header else None,
            container=args.container,
            section=args.section,
            expect_hash=args.expect_hash,
        )
        return {"hash": doc.content_hash}
    if command == "validate":
        result = doc.validate()
        return result
    if command == "diff":
        # diff requires two documents - handled in run() directly
        raise DocxError("diff is handled specially in run(), not in _dispatch")
    raise DocxError(f"unknown command: {command!r}")


def _resolve_refs(obj: Any, ctx: dict[str, Any]) -> Any:
    """Recursively replace $var references in strings with bound values."""
    if isinstance(obj, str) and obj.startswith("$"):
        var = obj[1:]
        if var not in ctx:
            raise DocxError(f"undefined variable: ${var}")
        return ctx[var]
    elif isinstance(obj, dict):
        return {k: _resolve_refs(v, ctx) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [_resolve_refs(item, ctx) for item in obj]
    return obj


def _run_batch(doc: DocxDocument, ops: list[dict[str, Any]]) -> dict[str, Any]:
    results = []
    ctx: dict[str, Any] = {}  # variable bindings: {"t1": 42}

    for idx, op in enumerate(ops):
        method_name = op.get("op")
        kwargs = op.get("kwargs", {})

        # Resolve $variable references in kwargs before execution
        try:
            kwargs = _resolve_refs(kwargs, ctx)
        except DocxError as exc:
            return {"ok": False, "failed_at": idx, "error": str(exc), "results": results}

        method = getattr(doc, method_name, None)
        if method is None or method_name.startswith("_"):
            return {"ok": False, "failed_at": idx, "error": f"unknown op: {method_name!r}", "results": results}
        try:
            result = method(**kwargs)
            results.append(result)

            # Bind result to variable if "as" key is present
            var_name = op.get("as")
            if var_name:
                if isinstance(result, dict):
                    ctx[var_name] = result.get("id", result.get("row"))
                else:
                    ctx[var_name] = result
        except DocxError as exc:
            return {"ok": False, "failed_at": idx, "error": str(exc), "results": results}
    return {"ok": True, "results": results}


def run(argv: list[str]) -> dict[str, Any]:
    """Parse argv, execute the requested operation, and return a result dict.

    Never raises: all errors (bad args, DocxError, missing files, etc.) are
    reported as ``{"error": ..., "type": ...}`` so callers get structured
    JSON either way.
    """
    parser = build_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return {"error": "invalid arguments", "type": "ArgumentError", "code": exc.code}

    try:
        if args.command == "new":
            doc = DocxDocument.new()
            doc.save(args.path)
            return {"path": args.path}

        if args.command == "batch":
            raw = sys.stdin.read() if args.ops == "-" else args.ops
            ops = json.loads(raw)
            doc = DocxDocument.open(args.path)
            result = _run_batch(doc, ops)
            if result["ok"]:
                doc.save(args.path)
            return result

        if args.command == "diff":
            doc1 = DocxDocument.open(args.path1)
            doc2 = DocxDocument.open(args.path2)
            return DocxDocument.diff(doc1, doc2)

        doc = DocxDocument.open(args.path)
        result = _dispatch(doc, args.command, args)
        if args.command in _MUTATING_COMMANDS:
            doc.save(args.path)
        return result
    except DocxError as exc:
        return {"error": str(exc), "type": type(exc).__name__}
    except PackageNotFoundError as exc:
        return {"error": str(exc), "type": "FileNotFoundError"}
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        return {"error": str(exc), "type": type(exc).__name__}


def main(argv: list[str] | None = None) -> int:
    result = run(sys.argv[1:] if argv is None else argv)
    print(json.dumps(result, separators=(",", ":")))
    return 1 if "error" in result else 0


if __name__ == "__main__":
    sys.exit(main())
