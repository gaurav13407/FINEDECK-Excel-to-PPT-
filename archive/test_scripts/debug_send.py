import os
import asyncio
import traceback
from pathlib import Path

# Import the Brevo sender from the project
from src.backend.app.services.emails_utils import send_email_via_brevo, send_email, generate_code, make_verification_email


def load_dotenv(path: str = ".env") -> dict:
    d = {}
    p = Path(path)
    if not p.exists():
        return d
    for line in p.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue
        k, v = line.split("=", 1)
        d[k.strip()] = v.strip().strip('"').strip("'")
    return d


def main():
    # Load repo .env first, then overlay with process env
    env = load_dotenv(".env")
    brevo_key = os.getenv("BREVO_API_KEY") or env.get("BREVO_API_KEY")
    email_from = os.getenv("EMAIL_FROM") or env.get("EMAIL_FROM") or "no-reply@example.com"
    target = os.getenv("TARGET_TEST_EMAIL") or "shriyanshtyagi005@gmail.com"

    if not brevo_key:
        print("BREVO_API_KEY not found in environment or .env — cannot send via Brevo")
        return

    # Set these in process env so the project's utility can pick them up
    os.environ["BREVO_API_KEY"] = brevo_key
    os.environ["EMAIL_FROM"] = email_from

    print(f"Attempting to send verification to: {target} using sender: {email_from}")

    # Generate a human-friendly verification code and email content
    code = generate_code(8)
    content = make_verification_email(code)
    subject = "FinDeck verification code"

    print("Generated verification code:", code)

    try:
        # Use the project's send_email which prefers Brevo and falls back
        res = asyncio.run(send_email(target, subject, content))
        print("Send returned:")
        print(res)
        print("Verification email sent. Check inbox or spam folder.")
    except Exception:
        print("Send raised an exception (full traceback):")
        traceback.print_exc()


if __name__ == "__main__":
    main()
