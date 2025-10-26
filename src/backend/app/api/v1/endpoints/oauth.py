from fastapi import APIRouter, Request
from fastapi.responses import RedirectResponse, JSONResponse
from typing import Optional
from datetime import datetime, timedelta
import os

from core.config import settings
from core.database import get_collection
from services.emails_utils import generate_code, send_email, make_verification_email
from core.security import get_password_hash

try:
    from authlib.integrations.starlette_client import OAuth
except Exception:
    OAuth = None

# Initialize OAuth client if available and credentials present
oauth = None
# Support local/prod Google credentials
GOOGLE_CLIENT_ID_LOCAL = getattr(settings, "GOOGLE_CLIENT_ID_LOCAL", None) or os.getenv("GOOGLE_CLIENT_ID_LOCAL")
GOOGLE_CLIENT_SECRET_LOCAL = getattr(settings, "GOOGLE_CLIENT_SECRET_LOCAL", None) or os.getenv("GOOGLE_CLIENT_SECRET_LOCAL")
GOOGLE_CLIENT_ID_PROD = getattr(settings, "GOOGLE_CLIENT_ID_PROD", None) or os.getenv("GOOGLE_CLIENT_ID_PROD")
GOOGLE_CLIENT_SECRET_PROD = getattr(settings, "GOOGLE_CLIENT_SECRET_PROD", None) or os.getenv("GOOGLE_CLIENT_SECRET_PROD")
GOOGLE_CLIENT_ID_FALLBACK = getattr(settings, "GOOGLE_CLIENT_ID", None) or os.getenv("GOOGLE_CLIENT_ID")
GOOGLE_CLIENT_SECRET_FALLBACK = getattr(settings, "GOOGLE_CLIENT_SECRET", None) or os.getenv("GOOGLE_CLIENT_SECRET")

use_local = bool(getattr(settings, "debug", False) or getattr(settings, "environment", "").lower() == "development")
if use_local:
    google_client_id = GOOGLE_CLIENT_ID_LOCAL or GOOGLE_CLIENT_ID_FALLBACK or GOOGLE_CLIENT_ID_PROD
    google_client_secret = GOOGLE_CLIENT_SECRET_LOCAL or GOOGLE_CLIENT_SECRET_FALLBACK or GOOGLE_CLIENT_SECRET_PROD
else:
    google_client_id = GOOGLE_CLIENT_ID_PROD or GOOGLE_CLIENT_ID_FALLBACK or GOOGLE_CLIENT_ID_LOCAL
    google_client_secret = GOOGLE_CLIENT_SECRET_PROD or GOOGLE_CLIENT_SECRET_FALLBACK or GOOGLE_CLIENT_SECRET_LOCAL

if OAuth and google_client_id and google_client_secret:
    oauth = OAuth()
    oauth.register(
        name="google",
        client_id=google_client_id,
        client_secret=google_client_secret,
        server_metadata_url="https://accounts.google.com/.well-known/openid-configuration",
        client_kwargs={"scope": "openid email profile"},
    )
# Register Twitter (X) OAuth2 if configured
X_CLIENT_ID = getattr(settings, "X_CLIENT_ID", None) or os.getenv("X_CLIENT_ID")
X_CLIENT_SECRET = getattr(settings, "X_CLIENT_SECRET", None) or os.getenv("X_CLIENT_SECRET")
if OAuth and X_CLIENT_ID and X_CLIENT_SECRET:
    if not oauth:
        oauth = OAuth()
    oauth.register(
        name="twitter",
        client_id=X_CLIENT_ID,
        client_secret=X_CLIENT_SECRET,
        authorize_url="https://twitter.com/i/oauth2/authorize",
        access_token_url="https://api.twitter.com/2/oauth2/token",
        client_kwargs={"scope": "tweet.read users.read offline.access email"},
    )
else:
    # Support local/prod variants for X/Twitter
    X_CLIENT_ID_LOCAL = getattr(settings, "X_CLIENT_ID_LOCAL", None) or os.getenv("X_CLIENT_ID_LOCAL")
    X_CLIENT_SECRET_LOCAL = getattr(settings, "X_CLIENT_SECRET_LOCAL", None) or os.getenv("X_CLIENT_SECRET_LOCAL")
    X_CLIENT_ID_PROD = getattr(settings, "X_CLIENT_ID_PROD", None) or os.getenv("X_CLIENT_ID_PROD")
    X_CLIENT_SECRET_PROD = getattr(settings, "X_CLIENT_SECRET_PROD", None) or os.getenv("X_CLIENT_SECRET_PROD")
    X_CLIENT_ID_FALLBACK = X_CLIENT_ID
    X_CLIENT_SECRET_FALLBACK = X_CLIENT_SECRET
    if use_local:
        x_client_id = X_CLIENT_ID_LOCAL or X_CLIENT_ID_FALLBACK or X_CLIENT_ID_PROD
        x_client_secret = X_CLIENT_SECRET_LOCAL or X_CLIENT_SECRET_FALLBACK or X_CLIENT_SECRET_PROD
    else:
        x_client_id = X_CLIENT_ID_PROD or X_CLIENT_ID_FALLBACK or X_CLIENT_ID_LOCAL
        x_client_secret = X_CLIENT_SECRET_PROD or X_CLIENT_SECRET_FALLBACK or X_CLIENT_SECRET_LOCAL
    if OAuth and x_client_id and x_client_secret:
        if not oauth:
            oauth = OAuth()
        oauth.register(
            name="twitter",
            client_id=x_client_id,
            client_secret=x_client_secret,
            authorize_url="https://twitter.com/i/oauth2/authorize",
            access_token_url="https://api.twitter.com/2/oauth2/token",
            client_kwargs={"scope": "tweet.read users.read offline.access email"},
        )

# Register GitHub OAuth if configured (support local and production keys)
GITHUB_CLIENT_ID_LOCAL = getattr(settings, "GITHUB_CLIENT_ID_LOCAL", None) or os.getenv("GITHUB_CLIENT_ID_LOCAL")
GITHUB_CLIENT_SECRET_LOCAL = getattr(settings, "GITHUB_CLIENT_SECRET_LOCAL", None) or os.getenv("GITHUB_CLIENT_SECRET_LOCAL")
GITHUB_CLIENT_ID_PROD = getattr(settings, "GITHUB_CLIENT_ID_PROD", None) or os.getenv("GITHUB_CLIENT_ID_PROD")
GITHUB_CLIENT_SECRET_PROD = getattr(settings, "GITHUB_CLIENT_SECRET_PROD", None) or os.getenv("GITHUB_CLIENT_SECRET_PROD")
# Fallback to generic keys if explicit local/prod not provided
GITHUB_CLIENT_ID_FALLBACK = getattr(settings, "GITHUB_CLIENT_ID", None) or os.getenv("GITHUB_CLIENT_ID")
GITHUB_CLIENT_SECRET_FALLBACK = getattr(settings, "GITHUB_CLIENT_SECRET", None) or os.getenv("GITHUB_CLIENT_SECRET")

# Choose which credentials to use depending on environment/debug
use_local = bool(getattr(settings, "debug", False) or getattr(settings, "environment", "").lower() == "development")
if use_local:
    client_id = GITHUB_CLIENT_ID_LOCAL or GITHUB_CLIENT_ID_FALLBACK or GITHUB_CLIENT_ID_PROD
    client_secret = GITHUB_CLIENT_SECRET_LOCAL or GITHUB_CLIENT_SECRET_FALLBACK or GITHUB_CLIENT_SECRET_PROD
else:
    client_id = GITHUB_CLIENT_ID_PROD or GITHUB_CLIENT_ID_FALLBACK or GITHUB_CLIENT_ID_LOCAL
    client_secret = GITHUB_CLIENT_SECRET_PROD or GITHUB_CLIENT_SECRET_FALLBACK or GITHUB_CLIENT_SECRET_LOCAL

if OAuth and client_id and client_secret:
    if not oauth:
        oauth = OAuth()
    oauth.register(
        name="github",
        client_id=client_id,
        client_secret=client_secret,
        access_token_url="https://github.com/login/oauth/access_token",
        authorize_url="https://github.com/login/oauth/authorize",
        api_base_url="https://api.github.com/",
        client_kwargs={"scope": "user:email"},
    )

router = APIRouter()


@router.get("/status")
async def oauth_status():
    """Debug endpoint: report OAuth readiness and configured providers.

    Use this to quickly determine whether Authlib is available and which
    providers were registered from environment/config.
    """
    authlib_installed = OAuth is not None
    configured = []
    if oauth:
        try:
            configured = list(getattr(oauth, "_clients", {}).keys())
        except Exception:
            configured = []
    return JSONResponse({"authlib_installed": authlib_installed, "configured_providers": configured})


@router.get("/oauth/{provider}")
async def oauth_start(provider: str, redirect: Optional[str] = None, test_email: Optional[str] = None, request: Request = None):
    """Simple OAuth start endpoint (development helper).

    In debug mode this endpoint simulates an OAuth login by creating a user-like
    email (or using `test_email` query param), generating a verification code,
    sending the code via the existing email utilities, and redirecting the user
    to the frontend verification page.

    In production this should be replaced with a proper OAuth flow using
    provider client IDs/secrets and secure callback handling.
    """
    # Only provide simulated flow in debug or when explicitly provided a test_email
    if not getattr(settings, "debug", False) and not test_email:
        return JSONResponse({"detail": "OAuth simulation disabled. Configure real OAuth providers."}, status_code=501)

    # Determine email to use
    if test_email:
        email = test_email
    else:
        # Create a deterministic demo email per provider
        ts = int(datetime.utcnow().timestamp())
        email = f"{provider}_user_{ts}@example.com"

    # Generate code and persist activation record
    coll = get_collection("activation_codes")
    code = generate_code(8)
    now = datetime.utcnow()
    expiry_minutes = int(os.getenv("CODE_EXPIRY_MINUTES", "15"))
    expires_at = now + timedelta(minutes=expiry_minutes)
    doc = {
        "code": code,
        "purpose": "email_verification",
        "plan": None,
        "target_email": email,
        "created_at": now,
        "expires_at": expires_at,
        "redeemed": False,
        "redeemed_by": None,
        "paypal_txn_id": None,
        "meta": {
            "ip": request.client.host if request and request.client else None,
            "user_agent": request.headers.get("user-agent") if request else None,
            "provider": provider,
        }
    }
    await coll.insert_one(doc)

    # Send the email (best-effort)
    try:
        content = make_verification_email(code, purpose="email verification", expiration_minutes=expiry_minutes)
        await send_email(email, f"FinDeck verification code", content)
    except Exception:
        # don't block the user if email fails in dev
        pass

    # Build redirect URL to frontend verify page
    frontend = getattr(settings, "frontend_url", "http://localhost:3000").rstrip("/")
    redirect_path = redirect or f"verify.html?email={email}&purpose=email_verification"
    # If redirect already contains query params, assume it's a path; otherwise ensure full path
    if redirect_path.startswith("http"):
        target = redirect_path
    else:
        if redirect_path.startswith("/"):
            redirect_path = redirect_path.lstrip("/")
        target = f"{frontend}/{redirect_path}"

    # In debug include the code in the query params to ease testing
    if getattr(settings, "debug", False):
        sep = '&' if '?' in target else '?'
        target = f"{target}{sep}code={code}"

    return RedirectResponse(url=target)


@router.get("/github/login")
async def github_login(request: Request, redirect: Optional[str] = None):
    """Start GitHub OAuth login flow."""
    if not oauth or "github" not in oauth._clients:
        return JSONResponse({"detail": "GitHub OAuth not configured (missing client id/secret or authlib)."}, status_code=501)

    redirect_uri = f"{request.base_url.scheme}://{request.client.host if request.client else request.base_url.hostname}{request.url.path.replace('/github/login','/github/callback')}"
    state = {"next": redirect or f"verify.html?purpose=email_verification"}
    return await oauth.github.authorize_redirect(request, redirect_uri, state=state)


@router.get("/github/callback")
async def github_callback(request: Request):
    """Handle GitHub OAuth callback, fetch user/email, create or update user, send verification code, and redirect."""
    if not oauth or "github" not in oauth._clients:
        return JSONResponse({"detail": "GitHub OAuth not configured."}, status_code=501)

    token = await oauth.github.authorize_access_token(request)
    userinfo = None
    try:
        resp = await oauth.github.get("user", token=token)
        userinfo = resp.json()
    except Exception:
        userinfo = None

    if not userinfo:
        return JSONResponse({"detail": "Failed to obtain user info from GitHub."}, status_code=400)

    github_id = userinfo.get("id")
    name = userinfo.get("name") or userinfo.get("login")

    # Fetch user emails to get primary/verified email
    email = None
    try:
        emails_resp = await oauth.github.get("user/emails", token=token)
        emails = emails_resp.json()
        if isinstance(emails, list):
            # Prefer primary & verified
            primary = next((e for e in emails if e.get("primary") and e.get("verified")), None)
            if not primary:
                primary = next((e for e in emails if e.get("verified")), None)
            if not primary and len(emails) > 0:
                primary = emails[0]
            if primary:
                email = primary.get("email")
    except Exception:
        email = None

    # Fallback placeholder email
    if not email:
        email = f"github_user_{github_id}@example.com"

    users = get_collection("users")
    now = datetime.utcnow()

    user = None
    if github_id:
        user = await users.find_one({"providers.github.id": github_id})

    if not user and email:
        user = await users.find_one({"email": email})

    if user:
        await users.update_one({"_id": user.get("_id")}, {"$set": {"last_login": now, f"providers.github": {"id": github_id, "profile": userinfo, "last_login": now}}})
    else:
        user_doc = {
            "email": email,
            "name": name,
            "password_hash": None,
            "email_verified": False,
            "providers": {"github": {"id": github_id, "profile": userinfo, "last_login": now}},
            "created_at": now,
            "last_login": now,
            "plan": "Free",
            "subscription": {},
        }
        res = await users.insert_one(user_doc)
        user = await users.find_one({"_id": res.inserted_id})

    # Create verification code & send email
    codes = get_collection("activation_codes")
    code = generate_code(8)
    expiry_minutes = int(os.getenv("CODE_EXPIRY_MINUTES", "15"))
    expires_at = now + timedelta(minutes=expiry_minutes)
    code_doc = {
        "code": code,
        "purpose": "email_verification",
        "plan": None,
        "target_email": email,
        "created_at": now,
        "expires_at": expires_at,
        "redeemed": False,
        "redeemed_by": None,
        "paypal_txn_id": None,
        "meta": {"provider": "github"}
    }
    await codes.insert_one(code_doc)

    try:
        content = make_verification_email(code, purpose="email verification", expiration_minutes=expiry_minutes)
        await send_email(email, "FinDeck verification code", content)
    except Exception:
        pass

    frontend = getattr(settings, "frontend_url", "http://localhost:3000").rstrip("/")
    target = f"{frontend}/verify.html?email={email}&purpose=email_verification"
    if getattr(settings, "debug", False):
        sep = '&' if '?' in target else '?'
        target = f"{target}{sep}code={code}"

    return RedirectResponse(url=target)



@router.get("/google/login")
async def google_login(request: Request, redirect: Optional[str] = None):
    """Start Google OAuth (redirect to Google's consent screen).

    Requires GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET to be configured. If they
    are missing or Authlib is not installed, returns a 501 with a helpful message.
    """
    if not oauth:
        return JSONResponse({"detail": "Google OAuth not configured (missing client id/secret or authlib)."}, status_code=501)

    frontend = getattr(settings, "frontend_url", "http://localhost:3000").rstrip("/")
    redirect_uri = f"{request.base_url.scheme}://{request.client.host if request.client else request.base_url.hostname}{request.url.path.replace('/google/login','/google/callback')}"
    # If user provided a frontend redirect path, include it in state
    state = {"next": redirect or f"verify.html?purpose=email_verification"}
    return await oauth.google.authorize_redirect(request, redirect_uri, state=state)


@router.get("/google/callback")
async def google_callback(request: Request):
    """Handle Google OAuth callback, create/find user in DB, send verification code, and redirect to frontend verify page."""
    if not oauth:
        return JSONResponse({"detail": "Google OAuth not configured."}, status_code=501)

    token = await oauth.google.authorize_access_token(request)
    userinfo = None
    try:
        # Try to parse id_token first (openid)
        userinfo = await oauth.google.parse_id_token(request, token)
    except Exception:
        # Fallback to userinfo endpoint
        try:
            resp = await oauth.google.get("userinfo", token=token)
            userinfo = resp.json()
        except Exception:
            userinfo = None

    if not userinfo:
        return JSONResponse({"detail": "Failed to obtain user info from Google."}, status_code=400)

    email = userinfo.get("email")
    email_verified = bool(userinfo.get("email_verified", False))
    google_sub = userinfo.get("sub") or userinfo.get("id")
    name = userinfo.get("name") or userinfo.get("given_name") or email.split("@")[0]

    users = get_collection("users")
    now = datetime.utcnow()

    user = None
    if google_sub:
        user = await users.find_one({"providers.google.id": google_sub})

    if not user and email:
        user = await users.find_one({"email": email})

    if user:
        # Update provider info and last_login
        await users.update_one({"_id": user.get("_id")}, {"$set": {"last_login": now, f"providers.google": {"id": google_sub, "profile": userinfo, "last_login": now}, "email_verified": email_verified if email else user.get("email_verified", False)}})
    else:
        # Create new user
        user_doc = {
            "email": email,
            "name": name,
            "password_hash": None,
            "email_verified": email_verified,
            "providers": {"google": {"id": google_sub, "profile": userinfo, "last_login": now}},
            "created_at": now,
            "last_login": now,
            "plan": "Free",
            "subscription": {},
        }
        res = await users.insert_one(user_doc)
        user = await users.find_one({"_id": res.inserted_id})

    # Create a verification code (we send a code even for social logins to reuse existing UX)
    codes = get_collection("activation_codes")
    code = generate_code(8)
    expiry_minutes = int(os.getenv("CODE_EXPIRY_MINUTES", "15"))
    expires_at = now + timedelta(minutes=expiry_minutes)
    code_doc = {
        "code": code,
        "purpose": "email_verification",
        "plan": None,
        "target_email": email,
        "created_at": now,
        "expires_at": expires_at,
        "redeemed": False,
        "redeemed_by": None,
        "paypal_txn_id": None,
        "meta": {"provider": "google"}
    }
    await codes.insert_one(code_doc)

    # Send email (best-effort)
    try:
        content = make_verification_email(code, purpose="email verification", expiration_minutes=expiry_minutes)
        await send_email(email, "FinDeck verification code", content)
    except Exception:
        pass

    # Redirect to frontend verify page
    frontend = getattr(settings, "frontend_url", "http://localhost:3000").rstrip("/")
    target = f"{frontend}/verify.html?email={email}&purpose=email_verification"
    if getattr(settings, "debug", False):
        sep = '&' if '?' in target else '?'
        target = f"{target}{sep}code={code}"

    return RedirectResponse(url=target)


@router.get("/twitter/login")
async def twitter_login(request: Request, redirect: Optional[str] = None):
    """Start Twitter/X OAuth2 login flow."""
    if not oauth or "twitter" not in oauth._clients:
        return JSONResponse({"detail": "Twitter/X OAuth not configured (missing client id/secret or authlib)."}, status_code=501)

    # Build redirect_uri to our callback
    redirect_uri = f"{request.base_url.scheme}://{request.client.host if request.client else request.base_url.hostname}{request.url.path.replace('/twitter/login','/twitter/callback')}"
    state = {"next": redirect or f"verify.html?purpose=email_verification"}
    return await oauth.twitter.authorize_redirect(request, redirect_uri, state=state)


@router.get("/twitter/callback")
async def twitter_callback(request: Request):
    """Handle Twitter/X OAuth2 callback, create/find user, send verification code, redirect to verify page."""
    if not oauth or "twitter" not in oauth._clients:
        return JSONResponse({"detail": "Twitter/X OAuth not configured."}, status_code=501)

    token = await oauth.twitter.authorize_access_token(request)
    userinfo = None
    try:
        # Twitter v2: fetch user profile
        resp = await oauth.twitter.get("https://api.twitter.com/2/users/me?user.fields=id,name,username,profile_image_url", token=token)
        data = resp.json()
        # data shape: {"data": {"id": ..., "name": ..., "username": ...}}
        user_obj = data.get("data") if isinstance(data, dict) else None
        if user_obj:
            userinfo = {
                "id": user_obj.get("id"),
                "name": user_obj.get("name"),
                "username": user_obj.get("username"),
                "profile_image_url": user_obj.get("profile_image_url")
            }
    except Exception:
        userinfo = None

    if not userinfo:
        return JSONResponse({"detail": "Failed to obtain user info from Twitter/X."}, status_code=400)

    twitter_id = userinfo.get("id")
    name = userinfo.get("name") or userinfo.get("username")
    # Twitter may not return email via v2; fallback to a provider placeholder email
    email = None
    # Try to read email from token response if present
    if token and isinstance(token, dict):
        email = token.get("email") or token.get("user_email")
    if not email:
        email = f"twitter_user_{twitter_id}@example.com"

    users = get_collection("users")
    now = datetime.utcnow()

    user = None
    if twitter_id:
        user = await users.find_one({"providers.twitter.id": twitter_id})

    if not user and email:
        user = await users.find_one({"email": email})

    if user:
        await users.update_one({"_id": user.get("_id")}, {"$set": {"last_login": now, f"providers.twitter": {"id": twitter_id, "profile": userinfo, "last_login": now}}})
    else:
        user_doc = {
            "email": email,
            "name": name,
            "password_hash": None,
            "email_verified": False,
            "providers": {"twitter": {"id": twitter_id, "profile": userinfo, "last_login": now}},
            "created_at": now,
            "last_login": now,
            "plan": "Free",
            "subscription": {},
        }
        res = await users.insert_one(user_doc)
        user = await users.find_one({"_id": res.inserted_id})

    # Create a verification code and email it
    codes = get_collection("activation_codes")
    code = generate_code(8)
    expiry_minutes = int(os.getenv("CODE_EXPIRY_MINUTES", "15"))
    expires_at = now + timedelta(minutes=expiry_minutes)
    code_doc = {
        "code": code,
        "purpose": "email_verification",
        "plan": None,
        "target_email": email,
        "created_at": now,
        "expires_at": expires_at,
        "redeemed": False,
        "redeemed_by": None,
        "paypal_txn_id": None,
        "meta": {"provider": "twitter"}
    }
    await codes.insert_one(code_doc)

    try:
        content = make_verification_email(code, purpose="email verification", expiration_minutes=expiry_minutes)
        await send_email(email, "FinDeck verification code", content)
    except Exception:
        pass

    frontend = getattr(settings, "frontend_url", "http://localhost:3000").rstrip("/")
    target = f"{frontend}/verify.html?email={email}&purpose=email_verification"
    if getattr(settings, "debug", False):
        sep = '&' if '?' in target else '?'
        target = f"{target}{sep}code={code}"

    return RedirectResponse(url=target)
