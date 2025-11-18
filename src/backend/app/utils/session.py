"""
Redis-based session management utility for FastAPI.
Provides secure session storage using Upstash Redis.
"""

import json
import secrets
from typing import Optional, Dict, Any
from datetime import datetime, timedelta
import redis


class SessionManager:
    """Manages user sessions using Redis."""
    
    def __init__(self, redis_client: redis.Redis, session_prefix: str = "session:", 
                 default_ttl: int = 86400):
        """
        Initialize the session manager.
        
        Args:
            redis_client: Redis client instance
            session_prefix: Prefix for session keys in Redis
            default_ttl: Default time-to-live for sessions in seconds (default: 24 hours)
        """
        self.redis = redis_client
        self.session_prefix = session_prefix
        self.default_ttl = default_ttl
    
    def create_session(self, user_data: Dict[str, Any], ttl: Optional[int] = None) -> str:
        """
        Create a new session and return the session ID.
        
        Args:
            user_data: Dictionary containing user data to store
            ttl: Time-to-live in seconds (optional, uses default if not provided)
        
        Returns:
            Session ID (token)
        """
        session_id = secrets.token_urlsafe(32)
        session_key = f"{self.session_prefix}{session_id}"
        
        # Add metadata
        session_data = {
            "user_data": user_data,
            "created_at": datetime.utcnow().isoformat(),
            "last_accessed": datetime.utcnow().isoformat()
        }
        
        # Store in Redis
        self.redis.setex(
            session_key,
            ttl or self.default_ttl,
            json.dumps(session_data)
        )
        
        return session_id
    
    def get_session(self, session_id: str) -> Optional[Dict[str, Any]]:
        """
        Retrieve session data by session ID.
        
        Args:
            session_id: The session ID
        
        Returns:
            Session data dictionary or None if not found/expired
        """
        session_key = f"{self.session_prefix}{session_id}"
        
        try:
            session_data_bytes = self.redis.get(session_key)
            
            if not session_data_bytes:
                return None
            
            session_data = json.loads(session_data_bytes)
            
            # Update last accessed time
            session_data["last_accessed"] = datetime.utcnow().isoformat()
            
            # Refresh TTL
            current_ttl = self.redis.ttl(session_key)
            if current_ttl > 0:
                self.redis.setex(session_key, current_ttl, json.dumps(session_data))
            
            return session_data.get("user_data")
            
        except Exception as e:
            print(f"Error retrieving session: {e}")
            return None
    
    def update_session(self, session_id: str, user_data: Dict[str, Any]) -> bool:
        """
        Update session data.
        
        Args:
            session_id: The session ID
            user_data: New user data to store
        
        Returns:
            True if successful, False otherwise
        """
        session_key = f"{self.session_prefix}{session_id}"
        
        try:
            existing_data_bytes = self.redis.get(session_key)
            
            if not existing_data_bytes:
                return False
            
            existing_data = json.loads(existing_data_bytes)
            
            # Update user data
            session_data = {
                "user_data": user_data,
                "created_at": existing_data.get("created_at", datetime.utcnow().isoformat()),
                "last_accessed": datetime.utcnow().isoformat()
            }
            
            # Get current TTL and preserve it
            current_ttl = self.redis.ttl(session_key)
            if current_ttl > 0:
                self.redis.setex(session_key, current_ttl, json.dumps(session_data))
                return True
            
            return False
            
        except Exception as e:
            print(f"Error updating session: {e}")
            return False
    
    def delete_session(self, session_id: str) -> bool:
        """
        Delete a session.
        
        Args:
            session_id: The session ID
        
        Returns:
            True if deleted, False if not found
        """
        session_key = f"{self.session_prefix}{session_id}"
        result = self.redis.delete(session_key)
        return result > 0
    
    def refresh_session(self, session_id: str, ttl: Optional[int] = None) -> bool:
        """
        Refresh a session's TTL.
        
        Args:
            session_id: The session ID
            ttl: New time-to-live in seconds (optional)
        
        Returns:
            True if successful, False otherwise
        """
        session_key = f"{self.session_prefix}{session_id}"
        
        try:
            if self.redis.exists(session_key):
                self.redis.expire(session_key, ttl or self.default_ttl)
                return True
            return False
        except Exception as e:
            print(f"Error refreshing session: {e}")
            return False
