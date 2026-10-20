# MCP Question Paper Generator

A production-ready **FastMCP** server that enables AI assistants (Claude, ChatGPT, etc.) to generate traditionally styled university question papers, convert them into formatted Microsoft Word (`.docx`) documents via Pandoc, and safely email them to multiple recipients with strict human-in-the-loop guardrails.

---

## Architecture Overview

### Lightweight UUID-Based Temporary File Isolation
To guarantee safe concurrency across multiple users and agents without database overhead or state leaks, the server uses a stateless, collision-free temporary file pattern:
- **Unique File Paths:** Every generated question paper is isolated with a unique run identifier:
  ```text
  /tmp/{subject_code}_{run_id}.md    (e.g., /tmp/CS501_a1b2c3d4.md)
  ```
  *(Falls back to the local `output/` directory during local development if `/tmp` is unavailable).*
- **Immediate Disk Cleanup:** Once `send_email` converts and dispatches the document, both the temporary `.md` and generated `.docx` files are unlinked in a `finally` block.
- **Background Stale File Pruning:** An asynchronous cleanup routine (`cleanup_stale_temp_files`) runs automatically in the background to scan and purge any temporary `.md` or `.docx` files older than 30 minutes, preventing disk bloat in containerized environments like FastMCP Cloud.
- **Deprecated Architecture Notice:** The previous Supabase/PostgreSQL database approaches (`feature/postgres-thread-id` and `feature/fix-postgres-table-init`) have been deprecated and retired in favor of this lightweight filesystem approach.

---

## Safety Guardrails & Human-in-the-Loop Workflow

To prevent autonomous AI models from sending unreviewed exam papers or auto-guessing recipient addresses, the tool schemas enforce strict Pydantic validation:

1. **Mandatory Preview Review:** When `generate_question_paper` finishes, it returns the file path alongside explicit instructions instructing the AI agent to display the full markdown preview in chat for the user to review.
2. **Explicit Confirmation (`user_confirmed: bool`):** `send_email` requires `user_confirmed: bool = False`. If an AI agent attempts to invoke `send_email` without explicit user permission (`user_confirmed=False`), execution is immediately rejected with:
   > `Email dispatch rejected: User confirmation has not been explicitly granted. You must present the preview to the user and obtain their confirmation before setting user_confirmed=True.`
3. **No Email Guessing:** The Pydantic schema for `recipients` strictly instructs the model never to auto-fill or guess emails without direct user input.

---

## Tool Reference

### 1. `generate_question_paper` (async)
Builds a professional Markdown question paper according to standard university examination layouts.

**Parameters:**
- `university_name` (*str*): Full university / institution name.
- `exam_title` (*str*): Exam name (e.g. `"End Semester Examination"`).
- `subject_name` (*str*): Full subject title.
- `subject_code` (*str*): Subject code (e.g. `"CS501"`).
- `course` (*str*): Degree/programme (e.g. `"B.Tech Computer Science"`).
- `semester` (*str*): Academic semester (e.g. `"5th Semester"`).
- `exam_date` (*str*): Date of examination (e.g. `"November 2026"`).
- `total_marks` (*int*): Maximum score for the paper.
- `duration` (*str*): Duration string (e.g. `"3 Hours"`).
- `instructions` (*list[str]*): General exam instructions.
- `sections` (*list[dict]*): Sections array containing `section_title`, `section_marks`, `attempt_info`, and `questions`.
- `show_roll_no` (*bool*, optional): If `True` (default), adds a Roll No. box row placeholder at the top right.

**Formatting Features:**
- **Title Block:** Centered university title, exam header, course, semester, and subject.
- **Time & Marks Row:** Full-width row with left-aligned Time and right-aligned Max Marks with divider.
- **Section Styling:** Centered Times New Roman headings with attempt guidance.
- **Right-Aligned Marks:** Automatically positions bracketed marks (e.g., `[10]`) flush to the right margin.

**Returns:** Success message containing the temporary Markdown file path and safety protocol instructions.

---

### 2. `send_email` (async)
Converts the Markdown question paper to DOCX using a dynamic reference styling document, emails it to one or more recipients, and purges the temporary files.

**Parameters:**
- `markdown_file_path` (*str*): Path to the temporary `.md` file produced by `generate_question_paper`.
- `recipients` (*list[str]*): Validated list of recipient email addresses. Supports multiple recipients.
- `subject` (*str*, default: `"Question Paper"`): Email subject line.
- `body` (*str*, default: `"Please find the attached question paper."`): Body text.
- `user_confirmed` (*bool*, default: `False`): Mandatory security confirmation flag. Must be explicitly `True` to authorize email dispatch.

**Returns:** Success confirmation string listing all recipients.

---

## Environment Variables

Configure these in `.env` (local dev) or in your FastMCP Cloud Environment Settings:

```env
GMAIL_ADDRESS=your-email@gmail.com
GMAIL_APP_PASSWORD=your-16-char-app-password
```

---

## Local Development

Install dependencies using `uv`:

```bash
# Install dependencies
uv sync

# Run the MCP server
uv run python -m src.mcp_qpgenerator.server
```

---

## Deployment (FastMCP Cloud / Docker)

The project includes an optimized [Dockerfile](file:///c:/Users/Ujjwal/Desktop/MCP/Dockerfile) built on `python:3.12-slim-bookworm`:
- Includes system dependencies (`pandoc`, `binutils`, `ca-certificates`).
- Uses `uv` for fast dependency caching and installation.
- Self-contained Word document styling powered by `pypandoc-binary` and `python-docx`.
