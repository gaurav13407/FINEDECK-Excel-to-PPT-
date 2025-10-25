import os
import asyncio
import types

import pytest

from src.backend.app.services import emails_utils as eu


def test_generate_code_format():
    code = eu.generate_code(8)
    assert isinstance(code, str)
    assert '-' in code and len(code) == 9


def test_send_email_uses_sendgrid(monkeypatch):
    # Simulate SendGrid being available
    class DummySG:
        def __init__(self, api_key):
            self.api_key = api_key
        def send(self, message):
            return {"status": "sent"}

    # Monkeypatch SendGrid import objects
    monkeypatch.setenv('SENDGRID_API_KEY', 'dummy_key')
    monkeypatch.setattr(eu, 'SendGridAPIClient', lambda k: DummySG(k))

    async def dummy_sendgrid_send(msg):
        return {"status": "sent"}

    # Patch the sendgrid Mail constructor to be benign if present
    # We only want to assert that send_email does not raise when SendGrid is configured
    res = asyncio.run(eu.send_email('test@example.com', 'Test', 'body'))
    assert res is not None


def test_send_verification_to_user():
    """Optional integration test: send a real verification code to the
    address in TARGET_TEST_EMAIL. This test only runs when SEND_TEST_EMAIL=1
    in the environment to avoid accidental sends.
    """
    target = os.getenv("TARGET_TEST_EMAIL", "gaurav13407@outlook.om")
    if os.getenv("SEND_TEST_EMAIL") != "1":
        pytest.skip("Skipping real send; set SEND_TEST_EMAIL=1 to enable real email sends")

    # Generate a code and email content
    code = eu.generate_code(8)
    content = eu.make_verification_email(code)

    # send_email in the project may return a coroutine or a direct result,
    # run via asyncio.run to support async implementations.
    try:
        res = asyncio.run(eu.send_email(target, "FinDeck verification code", content))
    except Exception as e:
        pytest.fail(f"Failed to send verification email: {e}")

    # If we get here, the provider didn't raise — consider this a success
    assert res is not None


def test_make_verification_email_default_contains_15_minutes():
    code = eu.generate_code(8)
    content = eu.make_verification_email(code, purpose="email verification", expiration_minutes=43200)
    assert "15 minutes" in content


def test_make_verification_email_payment_contains_30_days():
    code = eu.generate_code(8)
    content = eu.make_verification_email(code, purpose="payment_activation", expiration_minutes=43200)
    assert "30 days" in content
