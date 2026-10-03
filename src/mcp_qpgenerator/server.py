"""FastMCP server exposing question-paper tools."""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Any

from fastmcp import FastMCP

from .paper import (
    QuestionPaper,
    build_markdown,
    paper_from_json,
    paper_to_markdown_file,
    parse_markdown,
)
from .docx_gen import paper_to_docx
from .gmail import share_docx
from .secrets import get_gmail_address, get_gmail_app_password

mcp = FastMCP("Question Paper Generator")


@mcp.tool
def generate_question_paper(
    college_name: str = "",
    subject: str = "",
    exam: str = "",
    program: str = "",
    date: str = "",
    time: str = "",
    duration: str = "",
    max_marks: str = "",
    questions: list[dict[str, Any]] | None = None,
    sections: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Build a QuestionPaper object from metadata and questions.

    Provide a flat list of questions OR a list of sections (each with its own
    questions). Each question is {"number": 1, "text": "...", "marks": 5}.
    """
    return _paper_to_dict(
        _build_paper(college_name, subject, exam, program, date, time,
                     duration, max_marks, questions, sections)
    )


def _build_paper(
    college_name: str,
    subject: str,
    exam: str,
    program: str,
    date: str,
    time: str,
    duration: str,
    max_marks: str,
    questions: list[dict[str, Any]] | None,
    sections: list[dict[str, Any]] | None,
) -> QuestionPaper:
    paper = QuestionPaper(
        college_name=college_name,
        subject=subject,
        exam=exam,
        program=program,
        date=date,
        time=time,
        duration=duration,
        max_marks=max_marks,
    )
    for q in questions or []:
        paper.questions.append(
            _question(q.get("number"), q.get("text", ""), q.get("marks"))
        )
    for sec in sections or []:
        paper.sections.append(_section(sec.get("section", ""), sec.get("questions", [])))
    _autonumber(paper)
    return paper


@mcp.tool
def question_paper_to_markdown(
    college_name: str = "",
    subject: str = "",
    exam: str = "",
    program: str = "",
    date: str = "",
    time: str = "",
    duration: str = "",
    max_marks: str = "",
    questions: list[dict[str, Any]] | None = None,
    sections: list[dict[str, Any]] | None = None,
    output_path: str = "question_paper.md",
) -> dict[str, Any]:
    """Create a .md question-paper file and return its path and content."""
    paper = _build_paper(
        college_name, subject, exam, program, date, time,
        duration, max_marks, questions, sections,
    )
    path = paper_to_markdown_file(paper, output_path)
    return {"path": path, "content": build_markdown(paper)}


@mcp.tool
def question_paper_to_docx(
    college_name: str = "",
    subject: str = "",
    exam: str = "",
    program: str = "",
    date: str = "",
    time: str = "",
    duration: str = "",
    max_marks: str = "",
    questions: list[dict[str, Any]] | None = None,
    sections: list[dict[str, Any]] | None = None,
    output_path: str = "question_paper.docx",
) -> dict[str, Any]:
    """Generate a .docx question paper and return its path."""
    paper = _build_paper(
        college_name, subject, exam, program, date, time,
        duration, max_marks, questions, sections,
    )
    path = paper_to_docx(paper, output_path)
    return {"path": path, "docx_b64": base64.b64encode(Path(path).read_bytes()).decode()}


@mcp.tool
def markdown_to_docx(
    markdown: str,
    output_path: str = "question_paper.docx",
) -> dict[str, Any]:
    """Convert a Markdown question paper into a .docx file."""
    paper = parse_markdown(markdown)
    path = paper_to_docx(paper, output_path)
    return {"path": path, "docx_b64": base64.b64encode(Path(path).read_bytes()).decode()}


@mcp.tool
def send_docx_email(
    recipients: list[str],
    docx_path: str | None = None,
    docx_bytes: bytes | None = None,
    attachment_name: str = "question_paper.docx",
    subject: str = "Question Paper",
    body: str = "Please find attached the question paper.",
    sender: str = "",
    smtp_server: str = "smtp.gmail.com",
    smtp_port: int = 587,
    username: str = "",
    password: str = "",
    use_smtp: bool = True,
    token_path: str = "token.json",
    auth: bool = True,
) -> dict[str, Any]:
    """Send a .docx question paper to one or more recipients.

    Use SMTP (default, e.g. Gmail app password) or the Gmail API
    (use_smtp=False with an OAuth token.json file).
    Set auth=False to send without credentials (only works against a local
    SMTP MTA that accepts unauthenticated mail, e.g. Postfix on localhost:25).
    Credentials default to GMAIL_ADDRESS / GMAIL_APP_PASSWORD from the
    environment (set them in the FastMCP Cloud Secrets UI, or in a local
    .env file which is gitignored).
    """
    username = username or get_gmail_address()
    password = password or get_gmail_app_password()
    sender = sender or username
    return share_docx(
        recipients=recipients,
        attachment_path=docx_path,
        attachment_bytes=docx_bytes,
        attachment_name=attachment_name,
        subject=subject,
        body=body,
        sender=sender,
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        username=username,
        password=password,
        use_smtp=use_smtp,
        token_path=token_path,
        auth=auth,
    )


def _question(number: Any, text: str, marks: Any) -> Any:
    from .paper import Question

    num = int(number) if number is not None else 0
    m = int(marks) if marks is not None and str(marks).strip() else None
    return Question(number=num, text=str(text).strip(), marks=m)


def _section(name: str, items: list[dict[str, Any]]) -> Any:
    from .paper import Section

    sec = Section(name=str(name))
    for q in items:
        sec.questions.append(_question(q.get("number"), q.get("text", ""), q.get("marks")))
    return sec


def _autonumber(paper: QuestionPaper) -> None:
    from .paper import _autonumber as _a

    _a(paper)


def _paper_to_dict(paper: QuestionPaper) -> dict[str, Any]:
    return {
        "college_name": paper.college_name,
        "subject": paper.subject,
        "exam": paper.exam,
        "program": paper.program,
        "date": paper.date,
        "time": paper.time,
        "duration": paper.duration,
        "max_marks": paper.max_marks,
        "questions": [
            {"number": q.number, "text": q.text, "marks": q.marks}
            for q in paper.questions
        ],
        "sections": [
            {
                "section": s.name,
                "questions": [
                    {"number": q.number, "text": q.text, "marks": q.marks}
                    for q in s.questions
                ],
            }
            for s in paper.sections
        ],
    }


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()