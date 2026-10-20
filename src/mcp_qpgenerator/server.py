"""
MCP Question Paper Generator - Server Entry Point.
Initializes FastMCP and exposes the tools for markdown generation and email dispatch.
"""

import logging
from dotenv import load_dotenv
from fastmcp import FastMCP
from pydantic import Field

from mcp_qpgenerator.generator import generate_question_paper_handler
from mcp_qpgenerator.emailer import send_email_handler

# --- Logging ---------------------------------------------------------------
logger = logging.getLogger("mcp_qpgenerator")
logging.basicConfig(level=logging.INFO)

# --- Load .env if present (local dev) --------------------------------------
load_dotenv()

# --- Initialize FastMCP App (v0.2.0 for gateway cache-busting) -------------
mcp = FastMCP(
    "Question Paper Generator",
    version="0.2.0",
)


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
        The file path to the generated temporary Markdown file and next-step safety directives.
    """
    return await generate_question_paper_handler(
        university_name=university_name,
        exam_title=exam_title,
        subject_name=subject_name,
        subject_code=subject_code,
        course=course,
        semester=semester,
        exam_date=exam_date,
        total_marks=total_marks,
        duration=duration,
        instructions=instructions,
        sections=sections,
        show_roll_no=show_roll_no,
    )


# ---------------------------------------------------------------------------
# TOOL 2 - Convert Markdown to DOCX and email it
# ---------------------------------------------------------------------------
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
    return await send_email_handler(
        markdown_file_path=markdown_file_path,
        recipients=recipients,
        subject=subject,
        body=body,
        user_confirmed=user_confirmed,
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()