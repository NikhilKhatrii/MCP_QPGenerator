"""FastMCP server exposing question-paper tools."""

from __future__ import annotations

import base64
import sys
from pathlib import Path
from typing import Any

# Allow this file to be loaded directly (e.g. `fastmcp run server.py` or by
# FastMCP Cloud) as well as imported as part of the package.
_SRC = str(Path(__file__).resolve().parent.parent)
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from fastmcp import FastMCP

from mcp_qpgenerator.paper import (
    QuestionPaper,
    build_markdown,
    paper_to_markdown_file,
    parse_markdown,
)
from mcp_qpgenerator.docx_gen import paper_to_docx, paper_to_docx_bytes
from mcp_qpgenerator.gmail import share_docx
from mcp_qpgenerator.secrets import get_gmail_address, get_gmail_app_password

mcp = FastMCP("Question Paper Generator")


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

    def _q(item: dict[str, Any]) -> Any:
        from mcp_qpgenerator.paper import Question

        num = int(item.get("number", 0)) if item.get("number") is not None else 0
        marks = item.get("marks")
        marks = int(marks) if marks is not None and str(marks).strip() else None
        return Question(number=num, text=str(item.get("text", "")).strip(), marks=marks)

    for q in questions or []:
        paper.questions.append(_q(q))
    for sec in sections or []:
        from mcp_qpgenerator.paper import Section

        name = str(sec.get("title") or sec.get("section") or sec.get("name") or "")
        section = Section(name=name, instructions=str(sec.get("instructions", "") or ""))
        for q in sec.get("questions", []):
            section.questions.append(_q(q))
        paper.sections.append(section)
    _renumber(paper)
    return paper


def _renumber(paper: QuestionPaper) -> None:
    n = 1
    for sec in paper.sections:
        for q in sec.questions:
            if not q.number:
                q.number = n
            n += 1
    for q in paper.questions:
        if not q.number:
            q.number = n
        n += 1


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
    output_path: str = "",
) -> dict[str, Any]:
    """Create a .md question paper.

    The Markdown text is returned directly, so no filesystem write is needed.
    If output_path is given, the file is also written there.
    """
    paper = _build_paper(
        college_name, subject, exam, program, date, time,
        duration, max_marks, questions, sections,
    )
    content = build_markdown(paper)
    path = None
    if output_path:
        path = paper_to_markdown_file(paper, output_path)
    return {"path": path, "content": content}


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
    output_path: str = "",
) -> dict[str, Any]:
    """Generate a .docx question paper.

    The docx is built in memory and returned as base64, so no filesystem
    write is needed and the tool works on read-only sandboxes. If
    output_path is given, the file is also written there.
    """
    paper = _build_paper(
        college_name, subject, exam, program, date, time,
        duration, max_marks, questions, sections,
    )
    docx_b64 = base64.b64encode(paper_to_docx_bytes(paper)).decode()
    path = None
    if output_path:
        path = paper_to_docx(paper, output_path)
    return {"path": path, "docx_b64": docx_b64}


@mcp.tool
def markdown_to_docx(
    markdown: str,
    output_path: str = "",
) -> dict[str, Any]:
    """Convert a Markdown question paper into a .docx file."""
    paper = parse_markdown(markdown)
    docx_b64 = base64.b64encode(paper_to_docx_bytes(paper)).decode()
    path = None
    if output_path:
        path = paper_to_docx(paper, output_path)
    return {"path": path, "docx_b64": docx_b64}


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


@mcp.tool
def generate_and_email_paper(
    recipients: list[str],
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
    attachment_name: str = "question_paper.docx",
    email_subject: str = "Question Paper",
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
    """Build a question paper, render it to .docx in memory, and email it.

    No temp file is written, so this works in sandboxes where /tmp is
    read-only or ephemeral. Credentials default to GMAIL_ADDRESS /
    GMAIL_APP_PASSWORD from the environment.
    """
    paper = _build_paper(
        college_name, subject, exam, program, date, time,
        duration, max_marks, questions, sections,
    )
    docx_bytes = paper_to_docx_bytes(paper)
    username = username or get_gmail_address()
    password = password or get_gmail_app_password()
    sender = sender or username
    return share_docx(
        recipients=recipients,
        attachment_bytes=docx_bytes,
        attachment_name=attachment_name,
        subject=email_subject,
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


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()