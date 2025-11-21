#!/usr/bin/env python3
"""
Middleware diagnostic script - checks if middleware is working correctly
"""

import os
import sys
from pathlib import Path

# Add paths
project_root = Path(__file__).parent
app_dir = project_root / "src" / "backend" / "app"
sys.path.insert(0, str(project_root))
sys.path.insert(0, str(app_dir))

from dotenv import load_dotenv
load_dotenv()

print("=" * 80)
print("MIDDLEWARE DIAGNOSTIC REPORT")
print("=" * 80)

# Check 1: Environment Variables
print("\n1. ENVIRONMENT VARIABLES:")
print(f"   ENABLE_REDIS_SESSIONS: {os.getenv('ENABLE_REDIS_SESSIONS')}")
print(f"   UPSTASH_REDIS_URL: {'SET' if os.getenv('UPSTASH_REDIS_URL') else 'NOT SET'}")
print(f"   SESSION_SECRET_KEY: {'SET' if os.getenv('SESSION_SECRET_KEY') else 'NOT SET'}")

# Check 2: Redis Connection
print("\n2. REDIS CONNECTION TEST:")
UPSTASH_REDIS_URL = os.getenv("UPSTASH_REDIS_URL")
if UPSTASH_REDIS_URL:
    try:
        import redis
        client = redis.from_url(UPSTASH_REDIS_URL)
        if client.ping():
            print("   ✅ Redis PING successful")
        else:
            print("   ❌ Redis PING failed")
    except Exception as e:
        print(f"   ❌ Redis connection error: {e}")
else:
    print("   ⚠️  UPSTASH_REDIS_URL not set")

# Check 3: Module Imports
print("\n3. MODULE IMPORT TEST:")
try:
    from utils.session import SessionManager
    print("   ✅ SessionManager imported successfully")
except ImportError as e:
    print(f"   ❌ Failed to import SessionManager: {e}")

try:
    from middleware import SessionMiddleware
    print("   ✅ SessionMiddleware imported successfully")
except ImportError as e:
    print(f"   ❌ Failed to import SessionMiddleware: {e}")

try:
    # Import from the correct location
    sys.path.insert(0, str(app_dir))
    import main
    app = main.app
    print("   ✅ FastAPI app imported successfully")
except ImportError as e:
    print(f"   ❌ Failed to import app: {e}")
except AttributeError as e:
    print(f"   ❌ App object not found: {e}")

# Check 4: Middleware Registration
print("\n4. MIDDLEWARE REGISTRATION TEST:")
try:
    middleware_count = len(app.user_middleware)
    print(f"   Total middleware registered: {middleware_count}")
    
    if middleware_count > 0:
        print("   Registered middleware:")
        for i, mw in enumerate(app.user_middleware):
            print(f"      {i+1}. {mw.cls.__name__}")
            if 'Session' in mw.cls.__name__:
                print("         ✅ SessionMiddleware detected!")
    else:
        print("   ⚠️  No middleware registered - SessionMiddleware NOT active")
except Exception as e:
    print(f"   ❌ Error checking middleware: {e}")

# Check 5: App State
print("\n5. APP STATE TEST:")
try:
    if hasattr(app.state, 'redis'):
        if app.state.redis is not None:
            print("   ✅ app.state.redis is set")
        else:
            print("   ⚠️  app.state.redis is None - Redis not configured in app")
    else:
        print("   ❌ app.state.redis not found - not set by run_server.py")
    
    if hasattr(app.state, 'session_secret'):
        if app.state.session_secret:
            print("   ✅ app.state.session_secret is set")
        else:
            print("   ⚠️  app.state.session_secret is None")
    else:
        print("   ❌ app.state.session_secret not found - not set by run_server.py")
except NameError:
    print("   ❌ app is not defined - import failed")
except Exception as e:
    print(f"   ❌ Error checking app state: {e}")

# Check 6: Path Configuration
print("\n6. PYTHON PATH:")
print(f"   Project root: {project_root}")
print(f"   App directory: {app_dir}")
print(f"   App dir exists: {app_dir.exists()}")
print(f"   Utils dir exists: {(app_dir / 'utils').exists()}")
print(f"   Middleware dir exists: {(app_dir / 'middleware').exists()}")

print("\n" + "=" * 80)
print("DIAGNOSTIC COMPLETE")
print("=" * 80)
