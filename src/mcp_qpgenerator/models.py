"""
Pydantic models and validation schemas for MCP Question Paper Generator.
Enforces strict guardrails against autonomous auto-dispatching.
"""

from pydantic import BaseModel, Field


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
