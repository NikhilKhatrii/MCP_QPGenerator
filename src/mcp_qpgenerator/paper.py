"""Question paper model, Markdown builder and Markdown parser."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class Question:
    number: int
    text: str
    marks: int | None = None


@dataclass
class Section:
    name: str
    questions: list[Question] = field(default_factory=list)


@dataclass
class QuestionPaper:
    college_name: str = ""
    subject: str = ""
    exam: str = ""
    program: str = ""
    date: str = ""
    time: str = ""
    duration: str = ""
    max_marks: str = ""
    sections: list[Section] = field(default_factory=list)
    questions: list[Question] = field(default_factory=list)


_META_KEYS = (
    ("exam", "exam"),
    ("program", "program"),
    ("date", "date"),
    ("time", "time"),
    ("duration", "duration"),
    ("max marks", "max_marks"),
)


def _question_from_dict(item: dict[str, Any]) -> Question:
    number = int(item.get("number", 0))
    text = str(item.get("text", "")).strip()
    marks = item.get("marks", None)
    if marks is not None and str(marks).strip():
        try:
            marks = int(marks)
        except (TypeError, ValueError):
            marks = None
    return Question(number=number, text=text, marks=marks)


def paper_from_json(data: Any) -> QuestionPaper:
    """Build a QuestionPaper from a dict or a JSON string.

    Accepted shapes:
      {"college_name":..., "subject":..., "questions": [ {number,text,marks}, ... ]}
      {"college_name":..., "sections": [ {"section":"A", "questions":[...]} , ... ]}
      or a bare list of the above items.
    """
    if isinstance(data, str):
        data = json.loads(data)

    paper = QuestionPaper()
    if isinstance(data, dict):
        for key in ("college_name", "subject", "exam", "program", "date",
                    "time", "duration", "max_marks"):
            if key in data and data[key] is not None:
                setattr(paper, key, str(data[key]))
        items = data.get("sections") or data.get("questions") or []
    else:
        items = data

    for item in items:
        if not isinstance(item, dict):
            continue
        if "section" in item or "questions" in item:
            sec = Section(name=str(item.get("section", "")))
            for q in item.get("questions", []):
                sec.questions.append(_question_from_dict(q))
            paper.sections.append(sec)
        else:
            paper.questions.append(_question_from_dict(item))

    _autonumber(paper)
    return paper


def _autonumber(paper: QuestionPaper) -> None:
    counter = 1
    for sec in paper.sections:
        for q in sec.questions:
            if not q.number:
                q.number = counter
            counter += 1
    for q in paper.questions:
        if not q.number:
            q.number = counter
        counter += 1


def build_markdown(paper: QuestionPaper) -> str:
    lines: list[str] = []
    if paper.college_name:
        lines.append(f"# {paper.college_name}")
        lines.append("")
    if paper.subject:
        lines.append(f"## {paper.subject}")
        lines.append("")

    meta_parts = []
    for label, attr in _META_KEYS:
        val = getattr(paper, attr, "")
        if val:
            meta_parts.append(f"**{label}:** {val}")
    if meta_parts:
        lines.append(" | ".join(meta_parts))
        lines.append("")

    lines.append("---")
    lines.append("")

    if paper.sections:
        for section in paper.sections:
            if section.name:
                lines.append(f"### {section.name}")
                lines.append("")
            for q in section.questions:
                marks = f" [{q.marks} Marks]" if q.marks is not None else ""
                lines.append(f"{q.number}. {q.text}{marks}")
            lines.append("")
            lines.append("---")
            lines.append("")
    else:
        for q in paper.questions:
            marks = f" [{q.marks} Marks]" if q.marks is not None else ""
            lines.append(f"{q.number}. {q.text}{marks}")
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def paper_to_markdown_file(paper: QuestionPaper, path: str | Path) -> str:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(build_markdown(paper), encoding="utf-8")
    return str(p.resolve())


_META_LINE_RE = re.compile(r"\*\*(.+?):\*\*\s*([^|]+)")
_QUESTION_RE = re.compile(r"^(\d+)\.\s+(.*)$")
_MARKS_RE = re.compile(r"\[(\d+)\s*Marks?\]", re.IGNORECASE)


def parse_markdown(text: str) -> QuestionPaper:
    """Parse the Markdown produced by build_markdown back into a QuestionPaper."""
    paper = QuestionPaper()
    current_section: Section | None = None

    for raw in text.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.startswith("### "):
            current_section = Section(name=line[4:].strip())
            paper.sections.append(current_section)
            continue
        if line.startswith("## "):
            paper.subject = line[3:].strip()
            continue
        if line.startswith("# "):
            paper.college_name = line[2:].strip()
            continue
        if line.startswith("---"):
            current_section = None
            continue
        if line.startswith("**") and "|" in line:
            for part in line.split("|"):
                m = _META_LINE_RE.search(part)
                if m:
                    key = m.group(1).strip().lower()
                    val = m.group(2).strip()
                    for needle, attr in _META_KEYS:
                        if key == needle:
                            setattr(paper, attr, val)
                            break
            continue

        m = _QUESTION_RE.match(line)
        if m:
            number = int(m.group(1))
            rest = m.group(2).strip()
            marks = None
            mm = _MARKS_RE.search(rest)
            if mm:
                marks = int(mm.group(1))
                rest = _MARKS_RE.sub("", rest).strip()
            q = Question(number=number, text=rest, marks=marks)
            if current_section is not None:
                current_section.questions.append(q)
            else:
                paper.questions.append(q)

    return paper