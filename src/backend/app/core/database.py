# MongoDB Database Connection Manager
# Handles all database connectivity and initialization
# - Creates async MongoDB client using Motor driver
# - Manages database connection lifecycle
# - Sets up database collections and indexes
# - Provides database dependency injection for endpoints
# - Handles connection pooling and error recovery
# - Database health check utilities
from motor.motor_asyncio import AsyncIOMotorClient
from .config import settings
#Global variables for connection
client=None
database=None

# Functions Created
async def connect_to_mongo():
    global client,database
    try:
        #create the client connection settings
        client=AsyncIOMotorClient(settings.database_url,maxPoolSize=20,minPoolSize=5,maxIdleTimeMS=30000,serverSelectionTimeoutMS=5000)
        # test the connection
        await client.admin.command('ping')

        # Get the database
        database=client[settings.database_name]
        print("✅ MongoDB connection established")

    except Exception as e:
        print(f"❌ MongoDB connection error: {e}")
        raise e
async def close_mongo_connection():
    global client,database
    try:
        if client is not None:
            client.close()
            print("✅ MongoDB connection closed")
            client = None
            database = None
    except Exception as e:
        print(f"❌ MongoDB disconnection error: {e}")
        raise e

def get_database():
    if database is None:
        raise Exception("Database not connected! Call connect_to_mongo() first")
    return database
def get_collection(name:str):
    if database is None:
        raise Exception("Database not connected! Call connect_to_mongo() first")
    return database[name]

async def setup_database_indexes():
    """Create indexes for better query performance"""
    try:
        # User collection indexes
        user_collection=get_collection("users")
        await user_collection.create_index("email",unique=True)
        await user_collection.create_index("created_at")

        # Files colection indexes
        files_collection=get_collection("files")
        await files_collection.create_index("user_id")
        await files_collection.create_index("uploaded_date")

        # Conversion collection indexes
        conversion_collection=get_collection("conversions")
        await conversion_collection.create_index("user_id")
        await conversion_collection.create_index("status")

        # Temp uploads with TTL
        temp_collection=get_collection("temp_uploads")
        await temp_collection.create_index("created_at",expireAfterSeconds=86400)
        print("✅ Database indexes created successfully")

    except Exception as e:
        print(f"Index creation failed:{e}")