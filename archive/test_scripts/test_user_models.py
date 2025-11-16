# Test User Models
from src.backend.app.models.user import (
    UserCreate, UserLogin, UserResponse, SubscriptionPlan, 
    TemplateCategory, has_template_access, has_ai_access,
    can_create_presentation, PLAN_CONFIGS
)
from datetime import datetime

print("🧪 Testing User Models...")

# Test 1: UserCreate validation
print("\n=== Test 1: User Creation Validation ===")
try:
    # Valid user
    valid_user = UserCreate(
        name="John Doe",
        email="john@example.com", 
        password="password123"
    )
    print("✅ Valid user creation: PASS")
    print(f"User: {valid_user.name}, {valid_user.email}")
except Exception as e:
    print(f"❌ Valid user creation failed: {e}")

try:
    # Invalid user (short password)
    invalid_user = UserCreate(
        name="John",
        email="john@example.com",
        password="123"  # Too short
    )
    print("❌ Invalid user validation: FAIL (should have failed)")
except Exception as e:
    print("✅ Invalid user validation: PASS (correctly rejected)")

# Test 2: UserLogin validation
print("\n=== Test 2: Login Validation ===")
try:
    login_data = UserLogin(
        email="john@example.com",
        password="password123"
    )
    print("✅ Login validation: PASS")
except Exception as e:
    print(f"❌ Login validation failed: {e}")

# Test 3: Subscription Plans
print("\n=== Test 3: Subscription Plans ===")
for plan in SubscriptionPlan:
    config = PLAN_CONFIGS[plan]
    print(f"✅ {config['name']}: ${config['price']}, {config['presentations_limit']} PPTs")

# Test 4: Template Access Logic
print("\n=== Test 4: Template Access Logic ===")
# Mock user data for testing
class MockUser:
    def __init__(self, plan):
        self.subscription = type('obj', (object,), {'plan': plan})

basic_user = MockUser(SubscriptionPlan.BASIC)
pro_user = MockUser(SubscriptionPlan.PRO)
enterprise_user = MockUser(SubscriptionPlan.ENTERPRISE)

# Test basic templates (all should have access)
print("Basic Templates:")
print(f"  Basic user: {has_template_access(basic_user, TemplateCategory.BASIC)}")
print(f"  Pro user: {has_template_access(pro_user, TemplateCategory.BASIC)}")
print(f"  Enterprise user: {has_template_access(enterprise_user, TemplateCategory.BASIC)}")

# Test professional templates (Pro + Enterprise only)
print("Professional Templates:")
print(f"  Basic user: {has_template_access(basic_user, TemplateCategory.PROFESSIONAL)}")
print(f"  Pro user: {has_template_access(pro_user, TemplateCategory.PROFESSIONAL)}")
print(f"  Enterprise user: {has_template_access(enterprise_user, TemplateCategory.PROFESSIONAL)}")

# Test premium templates (Enterprise only)
print("Premium Templates:")
print(f"  Basic user: {has_template_access(basic_user, TemplateCategory.PREMIUM)}")
print(f"  Pro user: {has_template_access(pro_user, TemplateCategory.PREMIUM)}")
print(f"  Enterprise user: {has_template_access(enterprise_user, TemplateCategory.PREMIUM)}")

# Test 5: AI Access
print("\n=== Test 5: AI Access ===")
# Mock users with AI features
class MockUserWithAI:
    def __init__(self, plan, ai_enabled=True):
        self.subscription = type('obj', (object,), {
            'plan': plan, 
            'ai_features_enabled': ai_enabled
        })

basic_user_ai = MockUserWithAI(SubscriptionPlan.BASIC)
enterprise_user_ai = MockUserWithAI(SubscriptionPlan.ENTERPRISE, True)

print(f"Basic user AI access: {has_ai_access(basic_user_ai)}")
print(f"Enterprise user AI access: {has_ai_access(enterprise_user_ai)}")

print("\n🎉 All tests completed!")
print("If you see ✅ for most tests, your models are working correctly!")