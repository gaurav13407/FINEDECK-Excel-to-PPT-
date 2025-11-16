# Test FastAPI App - How Dependencies Are Really Used
from fastapi import FastAPI, Depends, HTTPException
from src.backend.app.core.dependencies import get_current_user, get_current_active_user, get_db
from src.backend.app.core.database import connect_to_mongo, close_mongo_connection
from src.backend.app.core.security import create_access_token

app = FastAPI(title="FinDeck API Test")

# Startup event - connect to database
@app.on_event("startup")
async def startup_event():
    await connect_to_mongo()
    print("🚀 FastAPI app started and connected to MongoDB!")

# Shutdown event - close database
@app.on_event("shutdown")
async def shutdown_event():
    await close_mongo_connection()
    print("👋 FastAPI app shutting down")

# Public endpoint (no authentication needed)
@app.get("/")
async def root():
    return {"message": "FinDeck API is running!"}

# Test endpoint - create a token
@app.post("/test/create-token")
async def create_test_token():
    user_data = {"user_id": "test123", "email": "test@example.com"}
    token = create_access_token(user_data)
    return {"access_token": token, "token_type": "bearer"}

# Protected endpoint - uses your dependencies!
@app.get("/test/protected")
async def protected_route(current_user: dict = Depends(get_current_user)):
    return {
        "message": "This is a protected route!",
        "user": current_user,
        "status": "Your dependencies are working perfectly!"
    }

# Database test endpoint
@app.get("/test/database")
async def test_database(db = Depends(get_db)):
    return {
        "message": "Database connection working!",
        "database_type": str(type(db))
    }

if __name__ == "__main__":
    import uvicorn
    print("Starting FastAPI test server...")
    print("Visit: http://localhost:8000")
    print("API docs: http://localhost:8000/docs")
    uvicorn.run(app, host="0.0.0.0", port=8000)