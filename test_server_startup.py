"""
Simple test to verify server starts correctly with middleware
"""
import subprocess
import time
import requests

print("Testing FinDeck Server Startup with Redis Middleware...")
print("=" * 80)

# Start server in background
print("\n1. Starting server...")
process = subprocess.Popen(
    ["python", "src/backend/run_server.py"],
    stdout=subprocess.PIPE,
    stderr=subprocess.STDOUT,
    text=True,
    bufsize=1
)

# Wait and capture initial output
print("\nServer Output:")
print("-" * 80)
time.sleep(5)

# Read initial output
output_lines = []
while True:
    line = process.stdout.readline()
    if not line:
        break
    print(line.rstrip())
    output_lines.append(line)
    if "Uvicorn running" in line or "Application startup complete" in line:
        break

print("-" * 80)

# Check if middleware was enabled
middleware_enabled = any("Session middleware enabled" in line for line in output_lines)
redis_connected = any("Connected to Redis successfully" in line for line in output_lines)

print("\n2. Startup Analysis:")
print(f"   Redis Connected: {'✅ YES' if redis_connected else '❌ NO'}")
print(f"   Middleware Enabled: {'✅ YES' if middleware_enabled else '❌ NO'}")

# Test if server is responsive
print("\n3. Testing server response...")
time.sleep(2)
try:
    response = requests.get("http://localhost:8000/docs", timeout=5)
    if response.status_code == 200:
        print("   ✅ Server is responding (docs accessible)")
    else:
        print(f"   ⚠️  Server returned {response.status_code}")
except Exception as e:
    print(f"   ❌ Server not responding: {e}")

# Cleanup
print("\n4. Stopping server...")
process.terminate()
process.wait()
print("   ✅ Server stopped")

print("\n" + "=" * 80)
print("Test Complete!")
print("=" * 80)
