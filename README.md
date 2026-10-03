# MCP Question Paper Generator

A FastMCP server that builds college question papers and shares them.

## Tools

| Tool | Description |
|---|---|
| `generate_question_paper` | Build a paper object from metadata + questions/sections |
| `question_paper_to_markdown` | Create a `.md` question paper (college name, subject, exam, sections) |
| `question_paper_to_docx` | Generate a `.docx` question paper |
| `markdown_to_docx` | Convert an existing `.md` paper into a `.docx` |
| `send_docx_email` | Email the `.docx` to one or more recipients |

The intended pipeline is `.md` → `.docx`: write the paper as Markdown, then
convert it, so the formatting is easy to inspect before generating Word.

## Install

```bash
uv sync
```

## Run in dev / inspector mode

```bash
uv run fastmcp dev inspector -m mcp_qpgenerator.server --ui-port 6276 --server-port 6277
```

This starts the MCP Inspector and opens `http://127.0.0.1:6276` in your
browser automatically, with all five tools available in the UI.

The `--ui-port` / `--server-port` flags avoid the default ports (6274/6275)
being occupied by a previous run.

## Usage example

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
      "section": "Section A",
      "questions": [
        { "number": 1, "text": "Define derivative.", "marks": 2 }
      ]
    }
  ]
}
```

## Emailing

`send_docx_email` sends the `.docx` to one or more recipients.

- **SMTP with auth (default)** — `smtp.gmail.com:587`, username + password
  (for Gmail use an **App Password**). Change `smtp_server`/`smtp_port` for
  any other provider.
- **Gmail API** — `use_smtp=false` with an OAuth `token.json`.
- **No auth** — `auth=false`, connects to a local SMTP MTA without credentials
  (e.g. Postfix on `localhost:25`).

### Credentials

`GMAIL_ADDRESS` and `GMAIL_APP_PASSWORD` come from the environment. Locally
they are loaded from a gitignored `.env` file. On FastMCP Cloud (or any host)
set the same two variables in the platform's Secrets UI — the repo never
contains them.