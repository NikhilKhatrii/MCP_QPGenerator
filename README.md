# MCP Question Paper Generator

A FastMCP server with three tools for building college question papers and
emailing them.

## Tools

| Tool | What it does |
|---|---|
| `create_question_paper_markdown` | Fill a paper model (college, subject, exam, questions) and get Markdown back |
| `markdown_to_docx` | Convert Markdown text to a .docx file (via the `markdown2docx` library) |
| `send_docx_email` | Email a .docx to one or more recipients |

## Pipeline

```text
metadata + questions
        │
        ▼
create_question_paper_markdown ──► Markdown text
        │
        ▼
markdown_to_docx ──► .docx (base64)
        │
        ▼
send_docx_email ──► sent to recipients
```

The docx is built in memory, so no filesystem write is needed and the tools
work on read-only sandboxes.

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

## Emailing

`send_docx_email` accepts a file path, raw bytes, or the base64 returned by
`markdown_to_docx`. Credentials default to `GMAIL_ADDRESS` and
`GMAIL_APP_PASSWORD` from the environment. Locally they come from a gitignored
`.env` file; on FastMCP Cloud set them in the platform's Secrets UI.

- SMTP auth (default) — `smtp.gmail.com:587` with a Gmail App Password
- `use_smtp=false` — Gmail API with an OAuth `token.json`
- `auth=false` — no credentials, for a local MTA (e.g. Postfix on localhost:25)