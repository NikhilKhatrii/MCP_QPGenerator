"""
Markdown question paper generation and temporary file management.
Handles UUID temporary file creation and background stale file cleanup.
"""

import asyncio
import os
import pathlib
import re
import time
import uuid
import logging

logger = logging.getLogger("mcp_qpgenerator.generator")

# Output folder (/tmp is always writable in Docker; fallback for local dev)
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


def build_markdown_content(
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
    """Construct standard university examination markdown string."""
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

    return "\n".join(lines)


async def generate_question_paper_handler(
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
    Generate the question paper markdown and persist it to a unique temporary file.
    Triggers background cleanup and returns safety instructions for the calling agent.
    """
    # Trigger non-blocking background cleanup of any stale temp files (> 30 mins)
    await asyncio.to_thread(cleanup_stale_temp_files)

    md_content = build_markdown_content(
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

    # Save to unique temporary Markdown file with UUID
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
