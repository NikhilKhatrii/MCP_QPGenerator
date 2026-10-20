"""
MCP Question Paper Generator
Tools:
  1. generate_question_paper  - Builds a Markdown question paper from structured inputs
  2. send_email               - Converts the .md to .docx and emails it to recipients
"""

import os
import re
import smtplib
import pathlib
import logging
import subprocess
import asyncio
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dotenv import load_dotenv
import pypandoc
import uuid
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt, RGBColor
from pydantic import BaseModel, Field
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

# Maximum age for temporary question paper files (30 minutes)
MAX_TEMP_FILE_AGE_SECONDS = 30 * 60


def cleanup_stale_temp_files(max_age_seconds: int = MAX_TEMP_FILE_AGE_SECONDS) -> int:
    """
    Scan OUTPUT_DIR (and /tmp) for temporary question paper files (.md, .docx)
    older than max_age_seconds (default 30 mins) and delete them to prevent disk bloat.
    """
    now = time.time()
    cleaned = 0
    dirs_to_scan = {OUTPUT_DIR}
    if pathlib.Path("/tmp").exists():
        dirs_to_scan.add(pathlib.Path("/tmp"))

    for directory in dirs_to_scan:
        try:
            for item in directory.glob("*_*.md"):
                try:
                    if item.is_file() and (now - item.stat().st_mtime > max_age_seconds):
                        item.unlink(missing_ok=True)
                        cleaned += 1
                except Exception:
                    pass
            for item in directory.glob("*_*.docx"):
                try:
                    if item.is_file() and (now - item.stat().st_mtime > max_age_seconds):
                        item.unlink(missing_ok=True)
                        cleaned += 1
                except Exception:
                    pass
        except Exception as e:
            logger.warning("Error during temp file cleanup in %s: %s", directory, e)

    if cleaned > 0:
        logger.info("Cleaned up %d stale temporary file(s)", cleaned)
    return cleaned

FONT_NAME = "Times New Roman"

# --- Page layout (change these to adjust the paper) -------------------------
PAGE_WIDTH = Mm(210)        # A4. For US Letter use Inches(8.5)
PAGE_HEIGHT = Mm(297)       # A4. For US Letter use Inches(11)
MARGIN_LEFT_RIGHT = Inches(0.9)
MARGIN_TOP_BOTTOM = Inches(0.8)
ROLL_NO_BOXES = 10

mcp = FastMCP(
    "Question Paper Generator",
    version="0.2.0",
    description="MCP server to generate university question papers, convert to DOCX via Pandoc, and email them with safety confirmation guardrails.",
)


# ---------------------------------------------------------------------------
# Reference docx: controls fonts, styles and page layout of the output
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
    """Create a pandoc reference.docx with exam-paper styling and margins."""
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

    # Page size and margins (pandoc copies these into the output document)
    for section in doc.sections:
        section.page_width = PAGE_WIDTH
        section.page_height = PAGE_HEIGHT
        section.left_margin = MARGIN_LEFT_RIGHT
        section.right_margin = MARGIN_LEFT_RIGHT
        section.top_margin = MARGIN_TOP_BOTTOM
        section.bottom_margin = MARGIN_TOP_BOTTOM

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

    # Italic note under each section title, with a bit more room below
    note = styles.add_style("QPNote", WD_STYLE_TYPE.PARAGRAPH)
    note.base_style = styles["Normal"]
    _set_font(note, 11)
    note.font.italic = True
    note.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    note.paragraph_format.space_after = Pt(8)

    # Placeholder paragraph that is replaced by the Roll No. boxes later
    roll = styles.add_style("QPRollNo", WD_STYLE_TYPE.PARAGRAPH)
    roll.base_style = styles["Normal"]
    _set_font(roll, 11)

    doc.save(str(ref_path))
    return ref_path


# ---------------------------------------------------------------------------
# Post-processing helpers (run on the .docx after pandoc)
# ---------------------------------------------------------------------------
def _set_cell_borders(cell, sz: int = 8) -> None:
    """Draw a thin black border on all four sides of a table cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tc_pr.append(borders)


def _build_roll_no_table(doc):
    """Return a right-aligned 'Roll No.:' label followed by empty boxes."""
    table = doc.add_table(rows=1, cols=ROLL_NO_BOXES + 1)
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    table.autofit = False

    row = table.rows[0]
    row.height = Inches(0.3)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    for i, cell in enumerate(row.cells):
        cell.width = Inches(0.95) if i == 0 else Inches(0.28)
        para = cell.paragraphs[0]
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.space_before = Pt(0)
        if i == 0:
            run = para.add_run("Roll No.:")
            run.bold = True
            run.font.name = FONT_NAME
            run.font.size = Pt(11)
        else:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_cell_borders(cell)
    return table


def _replace_roll_no_placeholder(doc) -> None:
    """Swap the QPRollNo placeholder paragraph for the real boxes."""
    for para in doc.paragraphs:
        if para.style is not None and para.style.name == "QPRollNo":
            table = _build_roll_no_table(doc)
            anchor = para._p
            anchor.addprevious(table._tbl)
            anchor.getparent().remove(anchor)
            break


def _right_align_marks(doc) -> None:
    """Move the trailing '[5]' style marks of each question to the right edge."""
    section = doc.sections[0]
    tab_pos = section.page_width - section.left_margin - section.right_margin

    for para in doc.paragraphs:
        if len(para.runs) < 2:
            continue
        last = para.runs[-1]
        if not re.fullmatch(r"\[\d+\]", last.text.strip()):
            continue

        # Drop the space before the marks, then add a right-aligned tab stop
        prev = para.runs[-2]
        prev.text = prev.text.rstrip()
        para.paragraph_format.tab_stops.add_tab_stop(tab_pos, WD_TAB_ALIGNMENT.RIGHT)
        last.text = "\t" + last.text.strip()


def _postprocess_docx(docx_path: pathlib.Path) -> None:
    doc = Document(str(docx_path))
    _replace_roll_no_placeholder(doc)
    _right_align_marks(doc)
    doc.save(str(docx_path))


# ---------------------------------------------------------------------------
# TOOL 1 - Generate Markdown question paper
# ---------------------------------------------------------------------------
@mcp.tool()
async def generate_question_paper(
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
    show_roll_no: bool = True,
) -> str:
    """
    Generate a professional Markdown question paper and save it to a unique temporary file.

    IMPORTANT SAFETY & INTERACTION DIRECTIVE FOR AI AGENTS:
    - Never assume the question paper is finalized immediately without the user's consent.
    - Always present the generated Markdown preview to the user in chat first.
    - Ask the user if they would like to make any adjustments (e.g. marks, question wording, instructions).
    - Explicitly ask the user to confirm the content and provide/confirm the target recipient email addresses before invoking `send_email`.

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
    show_roll_no : bool
        If True (default), a Roll No. box row is shown at the top right.

    Returns
    -------
    str
        The file path to the generated temporary Markdown file.
    """
    # Trigger non-blocking background cleanup of any stale temp files (> 30 mins)
    await asyncio.to_thread(cleanup_stale_temp_files)

    lines: list[str] = []

    # --- Roll No. placeholder (turned into boxes when the docx is built) ---
    if show_roll_no:
        lines += [
            '::: {custom-style="QPRollNo"}',
            "ROLLNO",
            ":::",
            "",
        ]

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
                '::: {custom-style="QPNote"}',
                f"({'. '.join(note_parts)}.)",
                ":::",
                "",
            ]

        for q_num, question in enumerate(questions, 1):
            marks_suffix = f" **[{sec_marks}]**" if sec_marks else ""
            lines.append(f"{q_num}. {question}{marks_suffix}")
            lines.append("")

    md_content = "\n".join(lines)

    # --- Save to unique temporary Markdown file with UUID --------------------
    run_id = str(uuid.uuid4())[:8]
    safe_code = re.sub(r"[^\w\-]", "_", subject_code.strip()) or "exam"
    md_path = OUTPUT_DIR / f"{safe_code}_{run_id}.md"

    # Non-blocking file write
    await asyncio.to_thread(md_path.write_text, md_content, encoding="utf-8")

    return (
        f"Question paper generated successfully and stored temporarily at:\n{md_path}\n\n"
        "NEXT STEPS FOR AGENT (SAFETY PROTOCOL):\n"
        "1. Present the generated markdown preview to the user in chat for review.\n"
        "2. Ask the user if any edits (questions, marks, duration, etc.) are needed.\n"
        "3. Ask the user for explicit confirmation to email the paper and request the target recipient email addresses.\n"
        "4. DO NOT invoke `send_email` until the user has explicitly confirmed the content and recipients (setting user_confirmed=True)."
    )


# ---------------------------------------------------------------------------
# Helper: synchronous DOCX conversion and email sending executed in thread
# ---------------------------------------------------------------------------
def _convert_and_send_email_sync(
    markdown_file_path: str,
    recipients: list[str],
    subject: str,
    body: str,
) -> str:
    gmail_address = os.environ.get("GMAIL_ADDRESS")
    gmail_password = os.environ.get("GMAIL_APP_PASSWORD")
    if not gmail_address or not gmail_password:
        raise RuntimeError(
            "Missing email credentials: ensure GMAIL_ADDRESS and "
            "GMAIL_APP_PASSWORD are set in environment"
        )

    md_path = pathlib.Path(markdown_file_path)
    if not md_path.exists():
        raise FileNotFoundError(f"Markdown file not found: {markdown_file_path}")

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
        _postprocess_docx(docx_path)

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
    finally:
        # Immediately delete temporary files to clean up disk space
        if md_path.exists():
            try:
                md_path.unlink()
            except Exception:
                pass
        if docx_path.exists():
            try:
                docx_path.unlink()
            except Exception:
                pass

    return f"Email sent successfully to: {', '.join(recipients)}"


# ---------------------------------------------------------------------------
# TOOL 2 - Convert Markdown to DOCX and email it
# ---------------------------------------------------------------------------
class SendEmailArgs(BaseModel):
    """Explicit Pydantic schema model for send_email tool parameters."""
    markdown_file_path: str = Field(
        ...,
        description="Absolute path to the temporary .md file produced by generate_question_paper.",
    )
    recipients: list[str] = Field(
        ...,
        min_length=1,
        description="Explicit list of recipient email addresses. CRITICAL: The AI agent MUST NOT guess or auto-fill email addresses without asking the user directly in chat.",
    )
    subject: str = Field(
        default="Question Paper",
        description="Email subject line. MUST be explicitly confirmed by the user or default to 'Question Paper'.",
    )
    body: str = Field(
        default="Please find the attached question paper.",
        description="Plain-text body of the email.",
    )
    user_confirmed: bool = Field(
        default=False,
        description="Mandatory security confirmation flag. MUST only be set to True after the user has explicitly reviewed the preview and granted permission to dispatch the email. Never set to True autonomously.",
    )


@mcp.tool()
async def send_email(
    markdown_file_path: str = Field(
        ...,
        description="Absolute path to the temporary .md file produced by generate_question_paper.",
    ),
    recipients: list[str] = Field(
        ...,
        min_length=1,
        description="Explicit list of recipient email addresses. CRITICAL: The AI agent MUST NOT guess or auto-fill email addresses without asking the user directly in chat.",
    ),
    subject: str = Field(
        default="Question Paper",
        description="Email subject line. MUST be explicitly confirmed by the user or default to 'Question Paper'.",
    ),
    body: str = Field(
        default="Please find the attached question paper.",
        description="Plain-text body of the email.",
    ),
    user_confirmed: bool = Field(
        default=False,
        description="Mandatory security confirmation flag. MUST only be set to True after the user has explicitly reviewed the preview and granted permission to dispatch the email. Never set to True autonomously.",
    ),
) -> str:
    """
    Convert the Markdown question paper to DOCX, email it to recipients, and delete temporary files.

    IMPORTANT SAFETY & INTERACTION DIRECTIVE FOR AI AGENTS:
    - Never assume recipient emails automatically.
    - Always ask the user for confirmation and explicit recipient emails before calling this tool.
    - Ensure the user has reviewed the generated paper content.
    - Set user_confirmed=True ONLY if the user has provided explicit verbal confirmation.

    Parameters
    ----------
    markdown_file_path : str
        Absolute path to the temporary .md file produced by generate_question_paper.
    recipients : list[str]
        List of recipient email addresses.
    subject : str
        Email subject line (default: 'Question Paper').
    body : str
        Plain-text body of the email.
    user_confirmed : bool
        Security confirmation flag (default: False; must be True to proceed).

    Returns
    -------
    str
        A success message listing all recipients the email was sent to.
    """
    # Validate arguments through Pydantic model for robust schema enforcement
    args = SendEmailArgs(
        markdown_file_path=markdown_file_path,
        recipients=recipients,
        subject=subject,
        body=body,
        user_confirmed=user_confirmed,
    )

    if args.user_confirmed is not True:
        raise ValueError(
            "Email dispatch rejected: User confirmation has not been explicitly granted. "
            "You must present the preview to the user and obtain their confirmation before setting user_confirmed=True."
        )

    clean_recipients = [r.strip() for r in args.recipients if r and r.strip()]
    if not clean_recipients:
        raise ValueError("No recipients provided. Please confirm the recipient email address with the user.")

    try:
        result = await asyncio.to_thread(
            _convert_and_send_email_sync,
            args.markdown_file_path,
            clean_recipients,
            args.subject,
            args.body,
        )
        return result
    except Exception as e:
        logger.exception("Failed to convert and send email asynchronously")
        raise RuntimeError(f"Email sending failed: {e}") from e


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()