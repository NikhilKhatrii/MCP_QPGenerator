# MCP Question Paper Generator

A FastMCP server with three tools for building college question papers and
emailing them. Every tool writes its output to disk and returns the path, so
the LLM never has to carry large base64 payloads between calls.

## Tools

| Tool | What it does |
|---|---|
| `create_question_paper_markdown` | Fill a paper model (college, subject, exam, questions) and write a `.md` file |
| `markdown_to_docx` | Convert a `.md` file into a `.docx` file (via the `markdown2docx` library) |
| `send_docx_email` | Email a `.docx` file to one or more recipients |

## Pipeline

```text
metadata + questions
        │
        ▼
create_question_paper_markdown ──► question_paper.md  (path returned)
        │
        ▼
markdown_to_docx ──► question_paper.docx  (path returned)
        │
        ▼
send_docx_email ──► sent to recipients
```

Files are written under `tempfile.gettempdir()/mcp_qpgenerator/` by default;
pass `output_path` to choose another location.

## Install

```bash
uv sync
```

## Run in dev / inspector mode

```bash
uv run fastmcp dev inspector -m mcp_qpgenerator.server --ui-port 6276 --server-port 6277
```

Opens `http://127.0.0.1:6276` in your browser with all three tools.

## Usage

```json
{
  "college_name": "Example College",
  "subject": "Mathematics",
  "exam": "Midterm",
  "program": "B.Sc.",
  "date": "2026-10-05",
  "time": "10:00",
  "duration": "2 Hours",
  "max_marks": "50",
  "sections": [
    {
      "section": "A",
      "instructions": "Answer all questions.",
      "questions": [
        { "text": "Define derivative.", "marks": 2 },
        { "text": "State Pythagoras theorem.", "marks": 2 }
      ]
    }
  ]
}
```

Use either `questions` (flat list) or `sections` (grouped). Each question needs
a `text` string and an optional `marks` number.

## Emailing

`send_docx_email` takes the path returned by `markdown_to_docx`. Credentials
default to `GMAIL_ADDRESS` and `GMAIL_APP_PASSWORD` from the environment.
Locally they come from a gitignored `.env` file; on FastMCP Cloud set them in
the platform's Secrets UI.

- SMTP auth (default) — `smtp.gmail.com:587` with a Gmail App Password
- `use_smtp=false` — Gmail API with an OAuth `token.json`
- `auth=false` — no credentials, for a local MTA (e.g. Postfix on localhost:25)