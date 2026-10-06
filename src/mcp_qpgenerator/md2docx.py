"""Convert Markdown text into a .docx file using the markdown2docx library.

The library itself works on files, so we round-trip through a temp file
and return the resulting bytes.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

from Markdown2docx import Markdown2docx


def markdown_to_docx_bytes(markdown: str, project: str = "question_paper") -> bytes:
    """Convert Markdown text to .docx bytes."""
    with tempfile.TemporaryDirectory() as tmp:
        md_path = Path(tmp) / f"{project}.md"
        md_path.write_text(markdown, encoding="utf-8")
        converter = Markdown2docx(project=str(Path(tmp) / project), markdown=[markdown])
        converter.save()
        out = Path(tmp) / f"{project}.docx"
        return out.read_bytes()


def markdown_to_docx_file(markdown: str, path: str | Path) -> str:
    """Convert Markdown text to a .docx file and return the resolved path."""
    with tempfile.TemporaryDirectory() as tmp:
        project = Path(path).stem or "question_paper"
        converter = Markdown2docx(
            project=str(Path(tmp) / project), markdown=[markdown]
        )
        converter.save()
        out = Path(tmp) / f"{project}.docx"
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        Path(path).write_bytes(out.read_bytes())
    return str(Path(path).resolve())