from fastapi import APIRouter, Depends, HTTPException, status, Request
from pydantic import BaseModel, EmailStr
from datetime import datetime, timedelta
import os
from typing import Optional
from core.database import get_collection
from core.config import settings
from services.emails_utils import generate_code, send_email, make_verification_email

router = APIRouter()

# canonical purpose names
PURPOSE_EMAIL = "email_verification"
PURPOSE_PAYMENT = "payment_activation"
PURPOSE_PASSWORD_RESET = "password_reset"

CODE_EXPIRY_MINUTES = int(os.getenv("CODE_EXPIRY_MINUTES", "15"))


class SendCodeRequest(BaseModel):
    email: EmailStr
    purpose: str = PURPOSE_EMAIL  # 'email_verification' or 'payment_activation'
    plan: Optional[str] = None


class VerifyCodeRequest(BaseModel):
    email: EmailStr
    code: str
    purpose: str = PURPOSE_EMAIL


class RedeemCodeRequest(BaseModel):
    code: str


@router.post("/send", status_code=200)
async def send_code(payload: SendCodeRequest, request: Request):
    """Create a code record and send email."""
    coll = get_collection("activation_codes")
    code = generate_code(8)
    now = datetime.utcnow()
    # Use short-lived codes for most flows (email verification, password reset, etc.).
    # Reserve long-lived (30 days) expiry only for payment/activation codes.
    expiry_minutes = CODE_EXPIRY_MINUTES if payload.purpose != PURPOSE_PAYMENT else 60 * 24 * 30
    expires_at = now + timedelta(minutes=expiry_minutes)
    doc = {
        "code": code,
        "purpose": payload.purpose,
        "plan": payload.plan or None,
        "target_email": payload.email,
        "created_at": now,
        "expires_at": expires_at,
        "redeemed": False,
        "redeemed_by": None,
        "paypal_txn_id": None,
        "meta": {
            "ip": request.client.host if request.client else None,
            "user_agent": request.headers.get("user-agent")
        }
    }

    await coll.insert_one(doc)

    readable_purpose = payload.purpose.replace("_", " ")
    content = make_verification_email(code, purpose=readable_purpose, expiration_minutes=expiry_minutes)

    send_error = None
    try:
        await send_email(payload.email, f"FinDeck {readable_purpose.title()} Code", content)
    except Exception as e:
        send_error = str(e)
        print("❌ Failed to send email:", e)

    response = {"ok": True, "message": "Code generated and email queued (if configured)."}
    # In development mode return the generated code to make testing easier
    if getattr(settings, 'debug', False):
        response['code'] = code
        if send_error:
            response['send_error'] = send_error

    return response


@router.post("/verify", status_code=200)
async def verify_code(payload: VerifyCodeRequest):
    coll = get_collection("activation_codes")
    now = datetime.utcnow()
    q = {
        "code": payload.code,
        "purpose": payload.purpose,
        "target_email": payload.email,
        "redeemed": False,
        "expires_at": {"$gt": now}
    }
    doc = await coll.find_one(q)
    if not doc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Invalid or expired code")
    await coll.update_one({"_id": doc["_id"]}, {"$set": {"redeemed": True, "redeemed_at": now, "redeemed_by": payload.email}})
    return {"ok": True, "message": "Code verified"}


# Placeholder auth dependency
async def get_current_user():
    raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Auth dependency not configured")


@router.post("/redeem", status_code=200)
async def redeem_code(payload: RedeemCodeRequest, current_user=Depends(get_current_user)):
    coll = get_collection("activation_codes")
    now = datetime.utcnow()
    q = {"code": payload.code, "purpose": PURPOSE_PAYMENT, "redeemed": False, "expires_at": {"$gt": now}}
    doc = await coll.find_one(q)
    if not doc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Invalid or expired code")

    await coll.update_one({"_id": doc["_id"]}, {"$set": {"redeemed": True, "redeemed_by": str(current_user.get("id")), "redeemed_at": now}})

    users = get_collection("users")
    await users.update_one({"_id": current_user.get("id")}, {"$set": {"subscription.plan": doc.get("plan", "basic")}})

    return {"ok": True, "message": "Code redeemed and plan applied"}