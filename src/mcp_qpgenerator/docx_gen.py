"""Generate a Word (.docx) question paper from a QuestionPaper."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .paper import QuestionPaper, Question, Section


def _add_heading(doc: Document, text: str, size: int = 18, bold: bool = True) -> None:
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    return p


def _add_meta_line(doc: Document, meta: list[tuple[str, str]]) -> None:
    if not meta:
        return
    p = doc.add_paragraph()
    first = True
    for label, value in meta:
        if not first:
            p.add_run("  |  ")
        first = False
        r1 = p.add_run(f"{label}: ")
        r1.bold = True
        p.add_run(value)


def _add_question(doc: Document, q: Question, number: int) -> None:
    p = doc.add_paragraph()
    r_num = p.add_run(f"{number}. ")
    r_num.bold = True
    p.add_run(q.text)
    if q.marks is not None:
        p.add_run(f"   [{q.marks} Marks]")


def paper_to_docx(paper: QuestionPaper, path: str | Path) -> str:
    """Render a QuestionPaper to a .docx file and return the resolved path."""
    doc = Document()

    if paper.college_name:
        _add_heading(doc, paper.college_name, size=22)
    if paper.subject:
        _add_heading(doc, paper.subject, size=16, bold=False)

    meta = []
    for label, attr in (
        ("Exam", "exam"),
        ("Program", "program"),
        ("Date", "date"),
        ("Time", "time"),
        ("Duration", "duration"),
        ("Max Marks", "max_marks"),
    ):
        val = getattr(paper, attr, "")
        if val:
            meta.append((label, val))
    _add_meta_line(doc, meta)

    doc.add_paragraph()
    doc.add_paragraph("Instructions: Attempt all questions. Marks are indicated in brackets.")

    if paper.sections:
        for section in paper.sections:
            if section.name:
                _add_heading(doc, section.name, size=14)
            for i, q in enumerate(section.questions, start=1):
                _add_question(doc, q, i)
            doc.add_paragraph()
    else:
        for i, q in enumerate(paper.questions, start=1):
            _add_question(doc, q, i)

    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return str(out.resolve())