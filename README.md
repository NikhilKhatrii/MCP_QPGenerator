# MCP Question Paper Generator

An MCP server built with **FastMCP** that lets any AI assistant:

1. **Generate** a fully-formatted Markdown question paper
2. **Convert** it to a `.docx` file
3. **Email** the file to multiple recipients

---

## Tools

### `generate_question_paper`
Builds a professional Markdown question paper from structured inputs and saves it to `output/`.

**Key inputs:**
- `university_name`, `exam_title`, `subject_name`, `subject_code`
- `course`, `semester`, `exam_date`, `duration`, `total_marks`
- `instructions` — list of general instructions
- `sections` — list of dicts, each with:
  - `section_title` (e.g. `"Section A"`)
  - `section_marks` (marks per question)
  - `attempt_info` (e.g. `"Attempt any 5 out of 8"`)
  - `questions` — list of question strings

**Returns:** absolute path to the saved `.md` file

---

### `convert_to_docx`
Converts the Markdown file to a Word `.docx` document.

**Input:** `markdown_file_path` — path returned by `generate_question_paper`  
**Returns:** absolute path to the `.docx` file

---

### `send_email`
Emails the `.docx` file via Gmail to one or more recipients.

**Inputs:**
- `docx_file_path` — path returned by `convert_to_docx`
- `recipients` — list of email addresses
- `subject` (optional) — email subject line
- `body` (optional) — email body text

**Returns:** success confirmation message

---

## Running Locally

```bash
uv run python -m qpgen.server
```

## Publishing to FastMCP Cloud

Follow the [FastMCP Cloud deployment guide](https://fastmcp.com) to publish this server.

---

## Project Structure

```
MCP/
├── src/
│   └── qpgen/
│       ├── __init__.py
│       └── server.py       ← all 3 tools live here
├── output/                 ← generated files saved here
├── pyproject.toml
└── README.md
```
