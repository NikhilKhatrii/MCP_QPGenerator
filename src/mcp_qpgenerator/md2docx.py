"""Convert Markdown text into a .docx file using the markdown2docx library.

The library reads a .md file and writes a .docx file next to it. We round-trip
through a temp directory so the caller only needs a path back.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

from Markdown2docx import Markdown2docx


def markdown_to_docx(markdown: str, docx_path: str | Path) -> str:
    """Convert Markdown text to a .docx file and return the resolved path.

    The markdown is written to a temp .md file, the library converts it, and
    the resulting .docx is copied to docx_path.
    """
    out = Path(docx_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp:
        project = Path(tmp) / out.stem
        md_path = Path(tmp) / f"{out.stem}.md"
        md_path.write_text(markdown, encoding="utf-8")
        converter = Markdown2docx(project=str(project), markdown=[markdown])
        converter.eat_soup()
        converter.save()
        generated = Path(tmp) / f"{out.stem}.docx"
        out.write_bytes(generated.read_bytes())
    return str(out.resolve())