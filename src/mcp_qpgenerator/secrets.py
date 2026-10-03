"""Credential loading for the question-paper email tool.

Secrets come from environment variables, with a local `.env` file as a
convenience during development. The `.env` file is gitignored, so it never
leaves your machine. On FastMCP Cloud (or any host) you set the same
variables in the platform's Secrets UI and they are injected at runtime.
"""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

# Load .env from the project root if present (ignored by git).
load_dotenv(Path(__file__).resolve().parent.parent.parent / ".env")


def get_gmail_address() -> str:
    return os.environ.get("GMAIL_ADDRESS", "")


def get_gmail_app_password() -> str:
    return os.environ.get("GMAIL_APP_PASSWORD", "")


def has_credentials() -> bool:
    return bool(get_gmail_address() and get_gmail_app_password())


__all__ = [
    "get_gmail_address",
    "get_gmail_app_password",
    "has_credentials",
]