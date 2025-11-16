from src.backend.app.core.security import get_password_hash, verify_password, create_access_token, verify_token

# Test password hashing
print("=== Testing Password Hashing ===")
password = "mypassword123"
hashed = get_password_hash(password)
print(f"Original: {password}")
print(f"Hashed: {hashed}")

# Test password verification
is_valid = verify_password(password, hashed)
is_invalid = verify_password("wrongpassword", hashed)
print(f"Correct password verification: {is_valid}")  # Should be True
print(f"Wrong password verification: {is_invalid}")  # Should be False

print("\n=== Testing JWT Tokens ===")
# Test JWT token creation
user_data = {"user_id": "12345", "email": "test@example.com"}
token = create_access_token(user_data)
print(f"Created token: {token[:50]}...")

# Test JWT token verification
decoded = verify_token(token)
print(f"Decoded token: {decoded}")

# Test invalid token
invalid_decoded = verify_token("invalid_token")
print(f"Invalid token result: {invalid_decoded}")  # Should be None