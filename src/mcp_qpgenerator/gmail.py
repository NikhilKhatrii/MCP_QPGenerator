"""Gmail sharing tool: send a generated .docx to one or more recipients."""

from __future__ import annotations

import base64
import os
import smtplib
import ssl
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any

from mcp_qpgenerator.paper import QuestionPaper


def _build_message(
    sender: str,
    recipients: list[str],
    subject: str,
    body: str,
    attachment_path: str | None,
    attachment_bytes: bytes | None,
    attachment_name: str,
) -> MIMEMultipart:
    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    data = attachment_bytes
    if data is None and attachment_path:
        data = Path(attachment_path).read_bytes()

    if data:
        part = MIMEApplication(data)
        part.add_header(
            "Content-Disposition",
            "attachment",
            filename=Path(attachment_name).name,
        )
        msg.attach(part)
    return msg


def _send_smtp(
    msg: MIMEMultipart,
    recipients: list[str],
    smtp_server: str,
    smtp_port: int,
    username: str,
    password: str,
    use_tls: bool = True,
    auth: bool = True,
) -> str:
    """Send via SMTP. With auth=True (default) it logs in; with auth=False
    it connects without credentials (only works against a local/relay MTA
    that accepts unauthenticated mail, e.g. Postfix on localhost:25)."""
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.set_debuglevel(0)
        if auth:
            context = ssl.create_default_context()
            if use_tls:
                server.starttls(context=context)
            server.login(username, password)
        server.sendmail(msg["From"], recipients, msg.as_string())
    mode = "authenticated" if auth else "unauthenticated"
    return f"Sent via SMTP ({mode}) to {', '.join(recipients)}"


def _send_gmail_api(
    msg: MIMEMultipart,
    recipients: list[str],
    token_path: str,
    scopes: list[str],
) -> str:
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    creds = Credentials.from_authorized_user_file(token_path, scopes=scopes)

    if not creds or not creds.valid:
        raise RuntimeError(
            "Gmail token is missing or expired. Run the OAuth flow to refresh it."
        )

    service = build("gmail", "v1", credentials=creds)
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("ascii")
    sent = []
    for recipient in recipients:
        message = {"raw": raw}
        message["to"] = recipient
        sent.append(service.users().messages().send(userId="me", body=message).execute())
    return f"Sent via Gmail API to {len(sent)} recipient(s)"


def share_docx(
    recipients: list[str],
    attachment_path: str | None = None,
    attachment_bytes: bytes | None = None,
    attachment_name: str = "question_paper.docx",
    subject: str = "Question Paper",
    body: str = "Please find attached the question paper.",
    sender: str = "",
    smtp_server: str = "smtp.gmail.com",
    smtp_port: int = 587,
    username: str = "",
    password: str = "",
    use_smtp: bool = True,
    token_path: str = "token.json",
    auth: bool = True,
    scopes: list[str] = [
        "https://www.googleapis.com/auth/gmail.send",
    ],
) -> dict[str, Any]:
    """Send a .docx question paper to one or more recipients.

    Defaults to SMTP (works with any SMTP server, e.g. Gmail app passwords).
    Set use_smtp=False to use the Gmail API with an OAuth token file.
    Set auth=False to send without credentials (only works against a local
    SMTP MTA that accepts unauthenticated mail, e.g. Postfix on localhost:25).
    """
    if not recipients:
        raise ValueError("At least one recipient email address is required.")

    username = username or os.environ.get("GMAIL_ADDRESS", "")
    password = password or os.environ.get("GMAIL_APP_PASSWORD", "")
    if not sender:
        sender = username

    if not auth and not use_smtp:
        raise ValueError("auth=False only applies to the SMTP path.")

    msg = _build_message(
        sender=sender,
        recipients=recipients,
        subject=subject,
        body=body,
        attachment_path=attachment_path,
        attachment_bytes=attachment_bytes,
        attachment_name=attachment_name,
    )

    if use_smtp:
        if auth and not username:
            raise ValueError("SMTP auth requires a username.")
        if auth and not password:
            raise ValueError("SMTP auth requires a password.")
        detail = _send_smtp(
            msg, recipients, smtp_server, smtp_port, username, password,
            use_tls=True, auth=auth,
        )
    else:
        detail = _send_gmail_api(msg, recipients, token_path, scopes)

    return {
        "status": "sent",
        "detail": detail,
        "recipients": recipients,
        "attachment": attachment_name,
        "subject": subject,
    }


def send_question_paper(
    paper: QuestionPaper,
    recipients: list[str],
    docx_path: str,
    subject: str = "Question Paper",
    body: str = "Please find attached the question paper.",
    sender: str = "",
    username: str = "",
    password: str = "",
    smtp_server: str = "smtp.gmail.com",
    smtp_port: int = 587,
    use_smtp: bool = True,
    token_path: str = "token.json",
    auth: bool = True,
) -> dict[str, Any]:
    """Convenience wrapper: send an already-generated .docx file."""
    return share_docx(
        recipients=recipients,
        attachment_path=docx_path,
        attachment_name=Path(docx_path).name or "question_paper.docx",
        subject=subject,
        body=body,
        sender=sender,
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        username=username,
        password=password,
        use_smtp=use_smtp,
        token_path=token_path,
        auth=auth,
    )