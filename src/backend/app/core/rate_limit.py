"""
Rate Limiting Configuration for FinDeck API
===========================================

Implements per-user rate limiting for API endpoints using SlowAPI.
"""

from slowapi import Limiter
from slowapi.util import get_remote_address
from fastapi import Request


def get_user_identifier(request: Request) -> str:
    """
    Extract user identifier for rate limiting.
    Prefers user ID from JWT token, falls back to IP address.
    
    Args:
        request: FastAPI request object
        
    Returns:
        String identifier for rate limiting (user_id or IP)
    """
    # Try to get user from request state (set by auth dependency)
    user = getattr(request.state, "user", None)
    if user and hasattr(user, "id"):
        return f"user:{str(user.id)}"
    
    # Fallback to IP address for unauthenticated requests
    return f"ip:{get_remote_address(request)}"


# Create rate limiter instance
limiter = Limiter(
    key_func=get_user_identifier,
    default_limits=["100/hour"],  # Default: 100 requests per hour per user
    headers_enabled=True,  # Include rate limit info in response headers
)


# Rate limit configurations for different tiers
RATE_LIMITS = {
    "free": "5/hour",           # Free tier: 5 conversions per hour
    "basic": "15/hour",         # Basic tier: 15 conversions per hour  
    "pro": "50/hour",           # Pro tier: 50 conversions per hour
    "ai": "200/hour",           # AI/Enterprise tier: 200 conversions per hour
    "default": "10/hour"        # Default for unknown tiers
}


def get_rate_limit_for_plan(plan: str) -> str:
    """
    Get rate limit string for a subscription plan.
    
    Args:
        plan: Subscription plan name (free, basic, pro, ai)
        
    Returns:
        Rate limit string (e.g., "50/hour")
    """
    return RATE_LIMITS.get(plan.lower(), RATE_LIMITS["default"])
