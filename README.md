# docx-toolkit

A CLI over `python-docx` for agents to work with `.docx` files directly,
without going through a lossy markdown/pandoc round-trip. It never touches
`styles.xml` or document theme when editing an existing file, and edits
preserve run-level formatting (bold/italic/etc.) instead of collapsing it.

All output is **compact JSON on stdout, one object per call**. There is no
human-readable mode. On failure, stdout is `{"error": "...", "type": "..."}`
and the process exits 1 — always check for the `"error"` key rather than
relying only on the exit code if you're piping output further.

See SKILL.md for further details.
