"""
MCP Question Paper Generator
Tools:
  1. generate_question_paper  - Builds a Markdown question paper from structured inputs
  2. send_email               - Converts the .md to .docx and emails it to recipients
"""

import os
import smtplib
import pathlib
import logging
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dotenv import load_dotenv
import pypandoc
from fastmcp import FastMCP

# ── Logging ──────────────────────────────────────────────────────────────────
logger = logging.getLogger("mcp_qpgenerator")
logging.basicConfig(level=logging.INFO)

# ── Load .env if present (local dev) ─────────────────────────────────────────
load_dotenv()

# ── Output folder (/tmp is always writable in Docker; fallback for local dev) ─
OUTPUT_DIR = pathlib.Path("/tmp") if pathlib.Path("/tmp").exists() else pathlib.Path(__file__).parent.parent.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

mcp = FastMCP("Question Paper Generator")


# ─────────────────────────────────────────────────────────────────────────────
# TOOL 1 – Generate Markdown question paper
# ─────────────────────────────────────────────────────────────────────────────
@mcp.tool()
def generate_question_paper(
    university_name: str,
    exam_title: str,
    subject_name: str,
    subject_code: str,
    course: str,
    semester: str,
    exam_date: str,
    total_marks: int,
    duration: str,
    instructions: list[str],
    sections: list[dict],
) -> str:
    """
    Generate a professional Markdown question paper and save it to disk.

    Parameters
    ----------
    university_name : str
        Full name of the university / institution.
    exam_title : str
        E.g. "End Semester Examination" or "Mid-Term Test".
    subject_name : str
        Full subject name.
    subject_code : str
        Subject / paper code.
    course : str
        Degree programme, e.g. "B.Tech (Computer Science)".
    semester : str
        E.g. "5th Semester".
    exam_date : str
        Date of the exam, e.g. "November 2026".
    total_marks : int
        Total marks for the paper.
    duration : str
        Duration of the exam, e.g. "3 Hours".
    instructions : list[str]
        General instructions shown at the top of the paper.
        Example: ["Attempt all questions.", "Use blue/black pen only."]
    sections : list[dict]
        Each section is a dict with keys:
          - section_title (str)   : e.g. "Section A"
          - section_marks (int)   : marks per question in this section
          - attempt_info (str)    : e.g. "Attempt any 5 out of 8 questions"
          - questions (list[str]) : list of question strings

    Returns
    -------
    str
        Absolute path to the saved Markdown file.
    """
    lines: list[str] = []

    # ── Header ────────────────────────────────────────────────────────────────
    lines += [
        '<div align="center">',
        f"<h1>{university_name}</h1>",
        f"<h2>{exam_title}</h2>",
        f"<h3>{course} - Semester {semester}</h3>",
        f"<h3>{subject_name} ({subject_code})</h3>",
        "</div>",
        "",
        f'<table style="width:100%"><tr><td style="text-align:left"><strong>Time:</strong> {duration}</td><td style="text-align:right"><strong>Max Marks:</strong> {total_marks}</td></tr></table><hr>',
        "",
    ]

    # ── General Instructions ──────────────────────────────────────────────────
    if instructions:
        lines += [
            "### General Instructions",
            "",
        ]
        for i, inst in enumerate(instructions, 1):
            lines.append(f"{i}. {inst}")
        lines += ["", "<hr>", ""]

    # ── Sections ──────────────────────────────────────────────────────────────
    for sec in sections:
        title = sec.get("section_title", "Section")
        sec_marks = sec.get("section_marks", "")
        attempt_info = sec.get("attempt_info", "")
        questions: list[str] = sec.get("questions", [])

        lines += [
            f'<h3 align="center">{title}</h3>',
            "",
        ]
        if attempt_info:
            lines.append(f"*{attempt_info}*")
            lines.append("")
        if sec_marks:
            lines.append(f"*(Each question carries {sec_marks} marks)*")
            lines.append("")

        for q_num, question in enumerate(questions, 1):
            marks_suffix = f" **[{sec_marks}]**" if sec_marks else ""
            lines.append(f"{q_num}. {question}{marks_suffix}")
            lines.append("")

        lines += ["<hr>", ""]

    md_content = "\n".join(lines)

    # ── Save file ─────────────────────────────────────────────────────────────
    safe_subject = subject_name.replace(" ", "_").replace("/", "-")
    md_path = OUTPUT_DIR / f"{safe_subject}_question_paper.md"
    md_path.write_text(md_content, encoding="utf-8")

    return str(md_path)


# ─────────────────────────────────────────────────────────────────────────────
# TOOL 2 – Convert Markdown → DOCX and email it
# ─────────────────────────────────────────────────────────────────────────────
@mcp.tool()
def send_email(
    markdown_file_path: str,
    recipients: list[str],
    subject: str = "Question Paper",
    body: str = "Please find the attached question paper.",
) -> str:
    """
    Convert the Markdown question paper to DOCX and email it to recipients.

    Parameters
    ----------
    markdown_file_path : str
        Absolute path to the .md file produced by generate_question_paper.
    recipients : list[str]
        List of recipient email addresses.
    subject : str
        Email subject line.
    body : str
        Plain-text body of the email.

    Returns
    -------
    str
        A success message listing all recipients the email was sent to.
    """
    try:
        gmail_address = os.environ["GMAIL_ADDRESS"]
        gmail_password = os.environ["GMAIL_APP_PASSWORD"]
    except KeyError as e:
        logger.exception("Missing email credentials")
        raise RuntimeError("Missing email credentials: ensure GMAIL_ADDRESS and GMAIL_APP_PASSWORD are set in environment") from e

    md_path = pathlib.Path(markdown_file_path)
    if not md_path.exists():
        raise FileNotFoundError(f"Markdown file not found: {markdown_file_path}")

    if not recipients:
        raise ValueError("No recipients provided.")

    # ── Convert .md → .docx via pypandoc ──────────────────────────────────────
    docx_path = md_path.with_suffix(".docx")
    try:
        pypandoc.convert_file(str(md_path), "docx", outputfile=str(docx_path))
    except Exception as e:
        logger.exception("Failed to convert Markdown to DOCX")
        raise RuntimeError(f"DOCX conversion failed: {e}") from e

    # ── Build and send email ──────────────────────────────────────────────────
    try:
        msg = MIMEMultipart()
        msg["From"] = gmail_address
        msg["To"] = ", ".join(recipients)
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        with open(docx_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{docx_path.name}"',
        )
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(gmail_address, gmail_password)
            server.sendmail(gmail_address, recipients, msg.as_string())
    except Exception as e:
        logger.exception("Failed to send email")
        raise RuntimeError(f"Email sending failed: {e}") from e

    return f"Email sent successfully to: {', '.join(recipients)}"


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────
def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
