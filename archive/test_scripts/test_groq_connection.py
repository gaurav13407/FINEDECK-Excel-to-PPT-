"""
Test Groq API Connection
Verify that the API key works and test basic functionality
"""

import os
from dotenv import load_dotenv
from groq import Groq

# Load environment variables
load_dotenv()

def test_groq_connection():
    """Test basic Groq API connection"""
    
    print("🔧 Testing Groq API Connection...")
    print("=" * 60)
    
    # Get API key
    api_key = os.getenv("GROQ_API_KEY")
    
    if not api_key:
        print("❌ ERROR: GROQ_API_KEY not found in .env file")
        return False
    
    print(f"✅ API Key found: {api_key[:20]}...")
    
    try:
        # Initialize client
        client = Groq(api_key=api_key)
        print("✅ Groq client initialized")
        
        # Test with a simple completion
        print("\n📡 Sending test request to Groq API...")
        
        response = client.chat.completions.create(
            model="llama-3.1-8b-instant",  # Fastest model for testing
            messages=[
                {
                    "role": "user",
                    "content": "Say 'Hello from FinDeck AI!' in exactly 5 words."
                }
            ],
            temperature=0.5,
            max_tokens=50
        )
        
        # Get response
        message = response.choices[0].message.content
        print(f"✅ Response received: {message}")
        
        # Check usage
        print(f"\n📊 Token Usage:")
        print(f"   - Prompt tokens: {response.usage.prompt_tokens}")
        print(f"   - Completion tokens: {response.usage.completion_tokens}")
        print(f"   - Total tokens: {response.usage.total_tokens}")
        
        # Cost calculation (Groq pricing: $0.27 per 1M tokens)
        cost = (response.usage.total_tokens / 1_000_000) * 0.27
        print(f"   - Estimated cost: ${cost:.6f}")
        
        print("\n" + "=" * 60)
        print("🎉 SUCCESS! Groq API is working perfectly!")
        print("=" * 60)
        
        return True
        
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        print("\nTroubleshooting:")
        print("1. Check if API key is correct")
        print("2. Visit https://console.groq.com to verify key")
        print("3. Check if you have remaining quota")
        return False

def test_all_models():
    """Test all available Groq models"""
    
    print("\n\n🚀 Testing Available Groq Models...")
    print("=" * 60)
    
    api_key = os.getenv("GROQ_API_KEY")
    client = Groq(api_key=api_key)
    
    models = [
        ("llama-3.1-8b-instant", "Fastest - Good for titles, quick insights"),
        ("mixtral-8x7b-32768", "Balanced - Good for analysis"),
        ("llama-3.1-70b-versatile", "Most Capable - Best for complex analysis"),
    ]
    
    for model, description in models:
        try:
            print(f"\n📝 Testing: {model}")
            print(f"   Description: {description}")
            
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "user", "content": "What is 2+2? Answer in one word."}
                ],
                temperature=0.3,
                max_tokens=10
            )
            
            result = response.choices[0].message.content
            tokens = response.usage.total_tokens
            
            print(f"   ✅ Response: {result}")
            print(f"   📊 Tokens used: {tokens}")
            
        except Exception as e:
            print(f"   ❌ Error: {str(e)}")
    
    print("\n" + "=" * 60)

if __name__ == "__main__":
    # Test basic connection
    success = test_groq_connection()
    
    if success:
        # Test all models
        test_all_models()
        
        print("\n\n✨ Next Steps:")
        print("1. ✅ Groq API is ready to use")
        print("2. 📝 Create AI service (ai_service.py)")
        print("3. 🎨 Implement 5 AI features")
        print("4. 🚀 Start generating AI-powered presentations!")
