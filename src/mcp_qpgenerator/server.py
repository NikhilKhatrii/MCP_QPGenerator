"""FastMCP server: three tools for building and sharing question papers.

1. create_question_paper_markdown - fill a pydantic model and return Markdown
2. markdown_to_docx - convert Markdown text to a .docx file
3. send_docx_email - email a .docx to one or more recipients
"""

from __future__ import annotations

import base64
import sys
from pathlib import Path
from typing import Any

# Allow this file to be loaded directly (e.g. FastMCP Cloud) as well as
# imported as part of the package.
_SRC = str(Path(__file__).resolve().parent.parent)
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from fastmcp import FastMCP

from mcp_qpgenerator.paper import (
    QuestionPaper,
    renumber,
    build_markdown,
    paper_to_markdown_file,
)
from mcp_qpgenerator.md2docx import markdown_to_docx_bytes as md_to_docx
from mcp_qpgenerator.gmail import share_docx
from mcp_qpgenerator.secrets import get_gmail_address, get_gmail_app_password

mcp = FastMCP("Question Paper Generator")


@mcp.tool
def create_question_paper_markdown(
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
    """Build a question paper from metadata + questions and return Markdown.

    Fill in the college/subject/exam fields and the questions. Each question
    is {"number": 1, "text": "...", "marks": 5}. Group questions into sections
    with {"section": "A", "instructions": "...", "questions": [...]}.

    The Markdown text is returned directly; pass output_path to also write it.
    """
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

    def _q(item: dict[str, Any]):
        from mcp_qpgenerator.paper import Question

        num = item.get("number")
        num = int(num) if num is not None else 0
        marks = item.get("marks")
        marks = int(marks) if marks is not None and str(marks).strip() else None
        return Question(number=num, text=str(item.get("text", "")).strip(), marks=marks)

    for q in questions or []:
        paper.questions.append(_q(q))
    for sec in sections or []:
        from mcp_qpgenerator.paper import Section

        name = str(sec.get("section") or sec.get("title") or sec.get("name") or "")
        section = Section(name=name, instructions=str(sec.get("instructions", "") or ""))
        for q in sec.get("questions", []):
            section.questions.append(_q(q))
        paper.sections.append(section)
    renumber(paper)

    content = build_markdown(paper)
    path = None
    if output_path:
        path = paper_to_markdown_file(paper, output_path)
    return {"path": path, "content": content}


@mcp.tool
def markdown_to_docx(
    markdown: str,
    output_path: str = "",
) -> dict[str, Any]:
    """Convert Markdown text into a .docx file.

    Uses the markdown2docx library, so no regex parsing is needed. The docx
    is built in memory and returned as base64; pass output_path to also write
    it to disk.
    """
    docx_b64 = base64.b64encode(md_to_docx(markdown)).decode()
    path = None
    if output_path:
        from mcp_qpgenerator.md2docx import markdown_to_docx_file

        path = markdown_to_docx_file(markdown, output_path)
    return {"path": path, "docx_b64": docx_b64}


@mcp.tool
def send_docx_email(
    recipients: list[str],
    docx_path: str | None = None,
    docx_bytes: bytes | None = None,
    docx_b64: str | None = None,
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
    """Email a .docx question paper to one or more recipients.

    Pass a file path, raw bytes, or base64 (the output of markdown_to_docx).
    Credentials default to GMAIL_ADDRESS / GMAIL_APP_PASSWORD from the
    environment (set them in the FastMCP Cloud Secrets UI, or in a local
    .env file which is gitignored).

    SMTP auth is the default (Gmail App Password). Set use_smtp=False for the
    Gmail API with an OAuth token, or auth=False for a local MTA.
    """
    if docx_b64 and docx_bytes is None:
        docx_bytes = base64.b64decode(docx_b64)
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


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()