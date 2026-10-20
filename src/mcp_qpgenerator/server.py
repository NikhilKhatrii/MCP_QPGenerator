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
import subprocess
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dotenv import load_dotenv
import pypandoc
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from fastmcp import FastMCP

# --- Logging ---------------------------------------------------------------
logger = logging.getLogger("mcp_qpgenerator")
logging.basicConfig(level=logging.INFO)

# --- Load .env if present (local dev) --------------------------------------
load_dotenv()

# --- Output folder (/tmp is always writable in Docker; fallback for local) --
OUTPUT_DIR = (
    pathlib.Path("/tmp")
    if pathlib.Path("/tmp").exists()
    else pathlib.Path(__file__).parent.parent.parent / "output"
)
OUTPUT_DIR.mkdir(exist_ok=True)

FONT_NAME = "Times New Roman"

mcp = FastMCP("Question Paper Generator")


# ---------------------------------------------------------------------------
# Helper: build a reference .docx that controls how the paper looks in Word
# ---------------------------------------------------------------------------
def _set_font(style, size: int, bold: bool | None = None) -> None:
    """Force a font on a style and strip theme fonts/colors that override it."""
    style.font.name = FONT_NAME
    style.font.size = Pt(size)
    style.font.color.rgb = RGBColor(0, 0, 0)
    if bold is not None:
        style.font.bold = bold

    rfonts = style.element.get_or_add_rPr().get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT_NAME)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


def _build_reference_docx() -> pathlib.Path:
    """Create a pandoc reference.docx with exam-paper styling."""
    ref_path = OUTPUT_DIR / "qp_reference.docx"

    # Start from pandoc's own default reference so all its styles exist
    with open(ref_path, "wb") as f:
        subprocess.run(
            [pypandoc.get_pandoc_path(), "--print-default-data-file", "reference.docx"],
            stdout=f,
            check=True,
        )

    doc = Document(str(ref_path))
    styles = doc.styles

    # Body text styles pandoc uses
    for name in ("Normal", "Body Text", "First Paragraph", "Compact"):
        if name in [s.name for s in styles]:
            _set_font(styles[name], 11)
            styles[name].paragraph_format.space_after = Pt(6)

    # Section headings (### in markdown): black, bold, centered
    if "Heading 3" in [s.name for s in styles]:
        h3 = styles["Heading 3"]
        _set_font(h3, 13, bold=True)
        h3.font.italic = False
        h3.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        h3.paragraph_format.space_before = Pt(14)
        h3.paragraph_format.space_after = Pt(4)

    # Custom styles used by the markdown (matched by name)
    uni = styles.add_style("QPUniversity", WD_STYLE_TYPE.PARAGRAPH)
    uni.base_style = styles["Normal"]
    _set_font(uni, 16, bold=True)
    uni.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    uni.paragraph_format.space_after = Pt(4)

    ctr = styles.add_style("QPCenter", WD_STYLE_TYPE.PARAGRAPH)
    ctr.base_style = styles["Normal"]
    _set_font(ctr, 12)
    ctr.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ctr.paragraph_format.space_after = Pt(2)

    doc.save(str(ref_path))
    return ref_path


# ---------------------------------------------------------------------------
# TOOL 1 - Generate Markdown question paper
# ---------------------------------------------------------------------------
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

    # --- Title block (centered via custom styles in the reference docx) ----
    lines += [
        '::: {custom-style="QPUniversity"}',
        university_name.upper(),
        ":::",
        "",
        '::: {custom-style="QPCenter"}',
        f"**{exam_title}**",
        "",
        f"{course}, {semester}",
        "",
        f"**{subject_name} ({subject_code})**",
        "",
        f"Date: {exam_date}",
        ":::",
        "",
    ]

    # --- Time (left) and Max Marks (right) on one full-width row -----------
    # The long separator line makes pandoc use the dash counts as relative
    # column widths, so the table spans the full page width.
    sep_left = ":" + "-" * 45
    sep_right = "-" * 45 + ":"
    lines += [
        "| | |",
        f"|{sep_left}|{sep_right}|",
        f"| **Time: {duration}** | **Max Marks: {total_marks}** |",
        "",
        "---",
        "",
    ]

    # --- General Instructions ----------------------------------------------
    if instructions:
        lines += ["**General Instructions:**", ""]
        for i, inst in enumerate(instructions, 1):
            lines.append(f"{i}. {inst}")
        lines += [""]

    # --- Sections ------------------------------------------------------------
    for sec in sections:
        title = sec.get("section_title", "Section")
        sec_marks = sec.get("section_marks", "")
        attempt_info = sec.get("attempt_info", "")
        questions: list[str] = sec.get("questions", [])

        lines += [f"### {title}", ""]

        note_parts = []
        if attempt_info:
            note_parts.append(attempt_info.rstrip("."))
        if sec_marks:
            note_parts.append(f"Each question carries {sec_marks} marks")
        if note_parts:
            lines += [
                '::: {custom-style="QPCenter"}',
                f"*({'. '.join(note_parts)}.)*",
                ":::",
                "",
            ]

        for q_num, question in enumerate(questions, 1):
            marks_suffix = f" **[{sec_marks}]**" if sec_marks else ""
            lines.append(f"{q_num}. {question}{marks_suffix}")
            lines.append("")

    md_content = "\n".join(lines)

    # --- Save file -----------------------------------------------------------
    safe_subject = subject_name.replace(" ", "_").replace("/", "-")
    md_path = OUTPUT_DIR / f"{safe_subject}_question_paper.md"
    md_path.write_text(md_content, encoding="utf-8")

    return str(md_path)


# ---------------------------------------------------------------------------
# TOOL 2 - Convert Markdown to DOCX and email it
# ---------------------------------------------------------------------------
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
        raise RuntimeError(
            "Missing email credentials: ensure GMAIL_ADDRESS and "
            "GMAIL_APP_PASSWORD are set in environment"
        ) from e

    md_path = pathlib.Path(markdown_file_path)
    if not md_path.exists():
        raise FileNotFoundError(f"Markdown file not found: {markdown_file_path}")

    if not recipients:
        raise ValueError("No recipients provided.")

    # --- Convert .md to .docx using the styled reference document ----------
    docx_path = md_path.with_suffix(".docx")
    try:
        ref_docx = _build_reference_docx()
        pypandoc.convert_file(
            str(md_path),
            "docx",
            format="markdown",
            outputfile=str(docx_path),
            extra_args=[f"--reference-doc={ref_docx}"],
        )
    except Exception as e:
        logger.exception("Failed to convert Markdown to DOCX")
        raise RuntimeError(f"DOCX conversion failed: {e}") from e

    # --- Build and send email ------------------------------------------------
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


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()