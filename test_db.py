# test_mongo.py
import asyncio
from motor.motor_asyncio import AsyncIOMotorClient
from src.backend.app.core.config import settings

async def test_connection():
    try:
        # Create client
        client = AsyncIOMotorClient(settings.database_url)
        
        # Test connection
        await client.admin.command('ping')
        print("✅ MongoDB connection successful!")
        
        # List databases (to verify access)
        db_list = await client.list_database_names()
        print(f"📊 Available databases: {db_list}")
        
        # Test our specific database
        db = client[settings.database_name]
        collections = await db.list_collection_names()
        print(f"📁 Collections in {settings.database_name}: {collections}")
        
        client.close()
        
    except Exception as e:
        print(f"❌ Connection failed: {e}")

# Run the test
if __name__ == "__main__":
    asyncio.run(test_connection())