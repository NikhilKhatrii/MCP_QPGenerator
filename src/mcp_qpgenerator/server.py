"""FastMCP server: three tools for building and sharing question papers.

1. create_question_paper_markdown - fill a pydantic model and write Markdown
2. markdown_to_docx - convert a Markdown file into a .docx file
3. send_docx_email - email a .docx file to one or more recipients

Every tool writes its output to disk and returns the path, so the LLM never
has to carry large base64 payloads between calls.
"""

from __future__ import annotations

import sys
import tempfile
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
from mcp_qpgenerator.md2docx import markdown_to_docx as _md_to_docx
from mcp_qpgenerator.gmail import share_docx
from mcp_qpgenerator.secrets import get_gmail_address, get_gmail_app_password

mcp = FastMCP("Question Paper Generator")

_DEFAULT_DIR = Path(tempfile.gettempdir()) / "mcp_qpgenerator"
_DEFAULT_DIR.mkdir(parents=True, exist_ok=True)


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
    """Build a question paper from metadata + questions and write Markdown.

    Fill in the college/subject/exam fields and the questions. Each question
    is {"text": "...", "marks": 5}. Group questions into sections with
    {"section": "A", "instructions": "...", "questions": [...]}.

    The paper is written to a .md file and its path returned, along with the
    Markdown text for quick preview. Pass output_path to choose the location.
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
        section = Section(
            name=name,
            instructions=str(sec.get("instructions", "") or ""),
        )
        for q in sec.get("questions", []):
            section.questions.append(_q(q))
        paper.sections.append(section)
    renumber(paper)

    content = build_markdown(paper)
    if not output_path:
        output_path = str(_DEFAULT_DIR / "question_paper.md")
    path = paper_to_markdown_file(paper, output_path)
    return {"path": path, "content": content}


@mcp.tool
def markdown_to_docx(
    markdown_path: str = "",
    markdown: str = "",
    output_path: str = "",
) -> dict[str, Any]:
    """Convert a Markdown file into a .docx file and return its path.

    Pass markdown_path (a .md file on disk) or raw markdown text. The docx is
    written to disk and its path returned.
    """
    if markdown_path:
        text = Path(markdown_path).read_text(encoding="utf-8")
    else:
        text = markdown
    if not output_path:
        output_path = str(_DEFAULT_DIR / "question_paper.docx")
    path = _md_to_docx(text, output_path)
    return {"path": path, "size": Path(path).stat().st_size}


@mcp.tool
def send_docx_email(
    recipients: list[str],
    docx_path: str,
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
    """Email a .docx file to one or more recipients.

    Pass the path returned by markdown_to_docx. Credentials default to
    GMAIL_ADDRESS / GMAIL_APP_PASSWORD from the environment (set them in the
    FastMCP Cloud Secrets UI, or in a local .env file which is gitignored).

    SMTP auth is the default (Gmail App Password). Set use_smtp=False for the
    Gmail API with an OAuth token, or auth=False for a local MTA.
    """
    username = username or get_gmail_address()
    password = password or get_gmail_app_password()
    sender = sender or username
    return share_docx(
        recipients=recipients,
        attachment_path=docx_path,
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