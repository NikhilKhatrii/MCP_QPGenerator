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
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dotenv import load_dotenv
import pypandoc
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt, RGBColor
import psycopg
from fastmcp import FastMCP

# --- Logging ---------------------------------------------------------------
logger = logging.getLogger("mcp_qpgenerator")
logging.basicConfig(level=logging.INFO)

# --- Load .env if present (local dev) --------------------------------------
load_dotenv()

# --- Database --------------------------------------------------------------
def get_db_connection() -> psycopg.Connection:
    """Create a new PostgreSQL database connection using DATABASE_URL."""
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        raise RuntimeError("DATABASE_URL environment variable is not set")
    return psycopg.connect(conninfo=db_url)


def init_db() -> None:
    """Create question_papers table if it does not exist."""
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS question_papers (
                        thread_id VARCHAR PRIMARY KEY,
                        markdown_content TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    );
                    """
                )
            conn.commit()
        logger.info("Database initialized successfully: question_papers table verified")
    except Exception as e:
        logger.warning("Database initialization skipped or failed: %s", e)

# --- Output folder (/tmp is always writable in Docker; fallback for local) --
OUTPUT_DIR = (
    pathlib.Path("/tmp")
    if pathlib.Path("/tmp").exists()
    else pathlib.Path(__file__).parent.parent.parent / "output"
)
OUTPUT_DIR.mkdir(exist_ok=True)

FONT_NAME = "Times New Roman"

# --- Page layout (change these to adjust the paper) -------------------------
PAGE_WIDTH = Mm(210)        # A4. For US Letter use Inches(8.5)
PAGE_HEIGHT = Mm(297)       # A4. For US Letter use Inches(11)
MARGIN_LEFT_RIGHT = Inches(0.9)
MARGIN_TOP_BOTTOM = Inches(0.8)
ROLL_NO_BOXES = 10

mcp = FastMCP("Question Paper Generator")


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
def generate_question_paper(
    thread_id: str,
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
    Generate a professional Markdown question paper and store it in PostgreSQL.

    Parameters
    ----------
    thread_id : str
        Unique identifier for the session/thread to isolate multi-user state.
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
        A success message indicating the paper was generated and stored.
    """
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

    # --- Save to PostgreSQL (insert or update based on thread_id) ------------
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS question_papers (
                        thread_id VARCHAR PRIMARY KEY,
                        markdown_content TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    );
                    """
                )
                cur.execute(
                    """
                    INSERT INTO question_papers (thread_id, markdown_content, created_at)
                    VALUES (%s, %s, CURRENT_TIMESTAMP)
                    ON CONFLICT (thread_id)
                    DO UPDATE SET
                        markdown_content = EXCLUDED.markdown_content,
                        created_at = CURRENT_TIMESTAMP;
                    """,
                    (thread_id, md_content),
                )
            conn.commit()
    except Exception as e:
        logger.exception("Failed to store question paper in database for thread %s", thread_id)
        raise RuntimeError(f"Database storage failed: {e}") from e

    return f"Question paper for '{subject_name}' generated and saved successfully for thread '{thread_id}'."


# ---------------------------------------------------------------------------
# TOOL 2 - Convert Markdown to DOCX and email it
# ---------------------------------------------------------------------------
@mcp.tool()
def send_email(
    thread_id: str,
    recipients: list[str],
    subject: str = "Question Paper",
    body: str = "Please find the attached question paper.",
) -> str:
    """
    Retrieve Markdown question paper by thread_id from database, convert to DOCX, email it, and clean up.

    Parameters
    ----------
    thread_id : str
        Unique identifier for the session/thread to retrieve stored question paper.
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
    if not recipients:
        raise ValueError("No recipients provided.")

    try:
        gmail_address = os.environ["GMAIL_ADDRESS"]
        gmail_password = os.environ["GMAIL_APP_PASSWORD"]
    except KeyError as e:
        logger.exception("Missing email credentials")
        raise RuntimeError(
            "Missing email credentials: ensure GMAIL_ADDRESS and "
            "GMAIL_APP_PASSWORD are set in environment"
        ) from e

    # --- Query markdown from PostgreSQL using thread_id ----------------------
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT markdown_content FROM question_papers WHERE thread_id = %s;",
                    (thread_id,),
                )
                row = cur.fetchone()
    except Exception as e:
        logger.exception("Failed to query database for thread %s", thread_id)
        raise RuntimeError(f"Database query failed: {e}") from e

    if not row or not row[0]:
        raise ValueError(f"No question paper found for thread_id: {thread_id}")

    md_content: str = row[0]

    # --- Write temporary .md file and convert to .docx ----------------------
    md_path = OUTPUT_DIR / f"{thread_id}_question_paper.md"
    docx_path = md_path.with_suffix(".docx")
    try:
        md_path.write_text(md_content, encoding="utf-8")
        ref_docx = _build_reference_docx()
        pypandoc.convert_file(
            str(md_path),
            "docx",
            format="markdown",
            outputfile=str(docx_path),
            extra_args=[f"--reference-doc={ref_docx}"],
        )
        _postprocess_docx(docx_path)
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
    finally:
        # Clean up temporary local files
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

    # --- Clean up database row after successful send -------------------------
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "DELETE FROM question_papers WHERE thread_id = %s;",
                    (thread_id,),
                )
            conn.commit()
        logger.info("Successfully cleaned up database record for thread %s", thread_id)
    except Exception as e:
        logger.warning("Failed to clean up database row for thread %s: %s", thread_id, e)

    return f"Email sent successfully to: {', '.join(recipients)}"


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> None:
    init_db()
    mcp.run()


if __name__ == "__main__":
    main()