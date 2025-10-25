import os
import secrets
from typing import Optional
import requests

from core.config import settings

try:
    from sendgrid import SendGridAPIClient
    from sendgrid.helpers.mail import Mail
except Exception:
    SendGridAPIClient = None

FRONTEND_URL = getattr(settings, "frontend_url", os.getenv("FRONTEND_URL", "http://localhost:3000"))
EMAIL_FROM = getattr(settings, "email_from", os.getenv("EMAIL_FROM", "no-reply@example.com"))


def generate_code(length: int = 8) -> str:
    """Generate a random verification code (default format: XXXX-XXXX)."""
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ23456789"
    raw = "".join(secrets.choice(alphabet) for _ in range(length))
    if length == 8:
        return raw[:4] + "-" + raw[4:]
    return raw


async def send_email_via_sendgrid(to_email: str, subject: str, content: str):
    """Send an email using SendGrid (async wrapper around the sync client)."""
    if not SendGridAPIClient:
        raise ImportError("SendGrid API client is not available.")
    api_key = getattr(settings, "sendgrid_api_key", None) or os.getenv("SENDGRID_API_KEY")
    if not api_key:
        raise RuntimeError("SendGrid API key is not configured.")

    message = Mail(
        from_email=EMAIL_FROM,
        to_emails=to_email,
        subject=subject,
        plain_text_content=content,
    )
    sg = SendGridAPIClient(api_key)
    # send() is synchronous; call it directly (small volumes) or run in thread if needed.
    resg = sg.send(message)
    return resg


async def send_email_via_brevo(to_email: str, subject: str, content: str):
    """Send email using Brevo (Transactional v3 API).

    Requires BREVO_API_KEY in Settings (.env) or environment.
    """
    api_key = getattr(settings, "brevo_api_key", None) or os.getenv("BREVO_API_KEY")
    if not api_key:
        raise RuntimeError("BREVO_API_KEY not set")

    url = os.getenv("BREVO_API_URL", "https://api.brevo.com/v3/smtp/email")
    sender_email = getattr(settings, "email_from", os.getenv("EMAIL_FROM", EMAIL_FROM))

    # Protect against accidental placeholder senders
    if sender_email.endswith("@example.com"):
        raise RuntimeError(
            f"EMAIL_FROM is set to a placeholder ({sender_email}). Set EMAIL_FROM to a verified sender (for example: no-reply@findeck.live) in your .env or environment."
        )

    payload = {
        "sender": {"name": "FinDeck", "email": sender_email},
        "to": [{"email": to_email}],
        "subject": subject,
        "textContent": content,
    }
    headers = {"accept": "application/json", "content-type": "application/json", "api-key": api_key}
    resp = requests.post(url, json=payload, headers=headers, timeout=10)
    resp.raise_for_status()
    return resp.json()


async def send_email(to_email: str, subject: str, content: str):
    """Send email using Brevo (preferred) or SendGrid (optional fallback).

    SMTP fallback has been removed — configure Brevo or SendGrid in your environment.
    """
    # Prefer Brevo if configured (check Settings first, then environment)
    brevo_key = getattr(settings, "brevo_api_key", None) or os.getenv("BREVO_API_KEY")
    if brevo_key:
        try:
            return await send_email_via_brevo(to_email, subject, content)
        except Exception:
            # bubble up the error to the caller; caller may log
            raise

    # Optional SendGrid fallback if available
    sendgrid_key = getattr(settings, "sendgrid_api_key", None) or os.getenv("SENDGRID_API_KEY")
    if sendgrid_key and SendGridAPIClient:
        try:
            return await send_email_via_sendgrid(to_email, subject, content)
        except Exception:
            raise

    raise RuntimeError(
        "No transactional email provider configured. Set BREVO_API_KEY (recommended) or SENDGRID_API_KEY in your environment."
    )


def make_verification_email(code: str, purpose: str = "email verification", expiration_minutes: int = 15):
    """Build a human-friendly verification email.

    For safety we present a clear, human-readable expiry string:
    - subscription/payment codes: show "30 days"
    - all other codes: show "15 minutes"

    This avoids exposing large raw-minute values like "43200 minutes" to users.
    """
    link = f"{FRONTEND_URL.rstrip('/')}/verify?code={code}"

    # Normalize purpose for matching; developers may pass different strings.
    p = (purpose or "").lower()
    if "payment" in p or "subscription" in p or "plan" in p:
        expiry_text = "30 days"
        purpose_line = "This is a subscription/payment activation code."
    else:
        # Force short expiry wording for non-subscription flows (email verification, password reset, etc.)
        expiry_text = "15 minutes"
        purpose_line = "Use this code to complete the requested action."

    return (
        f"Your FinDeck {purpose} code is: {code}\n\n"
        f"Or click the link below to verify:\n{link}\n\n"
        f"{purpose_line} This code will expire in {expiry_text}.\n\n"
        "If you did not request this, please ignore this email.\n\n"
    )