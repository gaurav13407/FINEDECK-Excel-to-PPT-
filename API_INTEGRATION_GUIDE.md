# API Integration Quick Reference

## Converter Usage

### Basic Usage

```python
from src.converter.excel_to_ppt_converter import convert_excel_to_ppt

result = convert_excel_to_ppt(
    excel_path="input.xlsx",
    output_path="output.pptx",
    user_tier="basic",  # 'free', 'basic', 'pro', 'ai_pro'
    user_ppt_count=0    # Current month's count
)
```

### Result Structure

```python
{
    'success': True/False,
    'output_path': 'path/to/output.pptx',
    'slides_created': 5,
    'template_used': 'financial_green',
    'ai_features_used': ['title', 'template_selection'],
    'ai_usage': {
        'tokens': 1067,
        'cost': 0.000288
    },
    'error': None  # or error message if failed
}
```

### FastAPI Endpoint Example

```python
from fastapi import APIRouter, UploadFile, Depends, HTTPException
from src.converter.excel_to_ppt_converter import convert_excel_to_ppt
import tempfile
import os

router = APIRouter()

@router.post("/convert")
async def convert(
    file: UploadFile,
    current_user: User = Depends(get_current_user)
):
    """Convert Excel to PPT with user's tier"""
    
    # 1. Check user subscription
    subscription = await get_user_subscription(current_user.id)
    user_tier = subscription.tier  # 'free', 'basic', 'pro', 'ai_pro'
    
    # 2. Get current month's PPT count
    ppt_count = await get_monthly_ppt_count(current_user.id)
    
    # 3. Save uploaded file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        content = await file.read()
        tmp.write(content)
        excel_path = tmp.name
    
    try:
        # 4. Convert
        output_path = f"temp/{current_user.id}_{file.filename}.pptx"
        result = convert_excel_to_ppt(
            excel_path=excel_path,
            output_path=output_path,
            user_tier=user_tier,
            user_ppt_count=ppt_count
        )
        
        # 5. Handle result
        if not result['success']:
            if result.get('upgrade_required'):
                raise HTTPException(
                    status_code=402,
                    detail=result['error']
                )
            raise HTTPException(
                status_code=400,
                detail=result['error']
            )
        
        # 6. Track usage
        await increment_ppt_count(current_user.id)
        await log_ai_usage(
            user_id=current_user.id,
            tokens=result['ai_usage']['tokens'],
            cost=result['ai_usage']['cost'],
            features=result['ai_features_used']
        )
        
        # 7. Return file
        return FileResponse(
            output_path,
            filename=f"{file.filename.split('.')[0]}.pptx",
            media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    finally:
        # Cleanup
        if os.path.exists(excel_path):
            os.remove(excel_path)
```

## Database Schema

### user_subscriptions
```sql
CREATE TABLE user_subscriptions (
    id SERIAL PRIMARY KEY,
    user_id INTEGER REFERENCES users(id),
    tier VARCHAR(20) NOT NULL,  -- 'free', 'basic', 'pro', 'ai_pro'
    status VARCHAR(20) DEFAULT 'active',
    started_at TIMESTAMP DEFAULT NOW(),
    expires_at TIMESTAMP,
    stripe_subscription_id VARCHAR(255),
    created_at TIMESTAMP DEFAULT NOW()
);
```

### usage_tracking
```sql
CREATE TABLE usage_tracking (
    id SERIAL PRIMARY KEY,
    user_id INTEGER REFERENCES users(id),
    month_year VARCHAR(7) NOT NULL,  -- '2024-12'
    ppt_count INTEGER DEFAULT 0,
    total_tokens INTEGER DEFAULT 0,
    total_cost DECIMAL(10,6) DEFAULT 0,
    last_reset TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);
```

### ai_usage_logs
```sql
CREATE TABLE ai_usage_logs (
    id SERIAL PRIMARY KEY,
    user_id INTEGER REFERENCES users(id),
    file_name VARCHAR(255),
    tier VARCHAR(20),
    tokens_used INTEGER,
    cost DECIMAL(10,6),
    features_used JSON,  -- ['title', 'summary', ...]
    slides_created INTEGER,
    template_used VARCHAR(50),
    created_at TIMESTAMP DEFAULT NOW()
);
```

## MongoDB Schema (Alternative)

```python
# users collection
{
    "_id": ObjectId(),
    "email": "user@example.com",
    "subscription": {
        "tier": "pro",
        "status": "active",
        "started_at": ISODate(),
        "expires_at": ISODate(),
        "stripe_id": "sub_xxx"
    },
    "usage": {
        "month_year": "2024-12",
        "ppt_count": 5,
        "total_tokens": 5335,
        "total_cost": 0.0014,
        "last_reset": ISODate()
    }
}

# ai_usage_logs collection
{
    "_id": ObjectId(),
    "user_id": ObjectId(),
    "file_name": "Sample_pnl.xlsx",
    "tier": "ai_pro",
    "tokens_used": 3553,
    "cost": 0.000959,
    "features_used": ["title", "summary", "insights", "template_selection", "layout", "chart_type"],
    "slides_created": 2,
    "template_used": "financial_green",
    "created_at": ISODate()
}
```

## Usage Tracking Functions

```python
async def get_monthly_ppt_count(user_id: int) -> int:
    """Get user's PPT count for current month"""
    month_year = datetime.now().strftime("%Y-%m")
    usage = await db.usage_tracking.find_one({
        "user_id": user_id,
        "month_year": month_year
    })
    return usage['ppt_count'] if usage else 0

async def increment_ppt_count(user_id: int):
    """Increment user's monthly PPT count"""
    month_year = datetime.now().strftime("%Y-%m")
    await db.usage_tracking.update_one(
        {"user_id": user_id, "month_year": month_year},
        {
            "$inc": {"ppt_count": 1},
            "$set": {"updated_at": datetime.now()}
        },
        upsert=True
    )

async def log_ai_usage(user_id: int, tokens: int, cost: float, features: list):
    """Log AI usage for billing/analytics"""
    await db.ai_usage_logs.insert_one({
        "user_id": user_id,
        "tokens_used": tokens,
        "cost": cost,
        "features_used": features,
        "created_at": datetime.now()
    })
    
    # Update monthly totals
    month_year = datetime.now().strftime("%Y-%m")
    await db.usage_tracking.update_one(
        {"user_id": user_id, "month_year": month_year},
        {
            "$inc": {
                "total_tokens": tokens,
                "total_cost": cost
            }
        },
        upsert=True
    )

async def reset_monthly_usage():
    """Reset usage counts (run monthly via cron)"""
    month_year = datetime.now().strftime("%Y-%m")
    await db.usage_tracking.update_many(
        {"month_year": {"$ne": month_year}},
        {
            "$set": {
                "ppt_count": 0,
                "total_tokens": 0,
                "total_cost": 0,
                "last_reset": datetime.now(),
                "month_year": month_year
            }
        }
    )
```

## Cron Job for Monthly Reset

```python
# Add to your scheduler (APScheduler, Celery, etc.)
from apscheduler.schedulers.asyncio import AsyncIOScheduler

scheduler = AsyncIOScheduler()

@scheduler.scheduled_job('cron', day=1, hour=0, minute=0)
async def monthly_reset_job():
    """Reset usage on 1st of each month at midnight"""
    print(f"Resetting monthly usage for {datetime.now().strftime('%Y-%m')}")
    await reset_monthly_usage()
    print("Monthly reset complete!")

scheduler.start()
```

## Error Handling

```python
# Error response structure
{
    "success": False,
    "error": "PPT limit reached. Basic ($25/month) allows 7 PPTs/month.",
    "upgrade_required": True,
    "current_tier": "basic",
    "suggested_tier": "pro",
    "upgrade_benefits": [
        "15 PPTs per month instead of 7",
        "Access to all 10 professional templates",
        "AI template selection"
    ]
}
```

## Frontend Integration

```typescript
// React/Next.js example
async function convertExcel(file: File) {
    const formData = new FormData();
    formData.append('file', file);
    
    try {
        const response = await fetch('/api/v1/convert', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`
            },
            body: formData
        });
        
        if (response.status === 402) {
            // Upgrade required
            const data = await response.json();
            showUpgradeModal(data);
            return;
        }
        
        if (!response.ok) {
            throw new Error('Conversion failed');
        }
        
        // Download PPT
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = file.name.replace('.xlsx', '.pptx');
        a.click();
        
    } catch (error) {
        console.error('Conversion error:', error);
    }
}
```

## Testing

```bash
# Test Free tier
curl -X POST http://localhost:8000/api/v1/convert \
  -H "Authorization: Bearer $TOKEN" \
  -F "file=@examples/Sample_pnl.xlsx" \
  -o output.pptx

# Test with tier override (for testing)
curl -X POST http://localhost:8000/api/v1/convert?tier=ai_pro \
  -H "Authorization: Bearer $TOKEN" \
  -F "file=@examples/Sample_pnl.xlsx" \
  -o output.pptx
```

## Pricing Display

```python
TIER_FEATURES = {
    'free': {
        'price': 0,
        'ppt_limit': 1,
        'features': [
            '1 presentation per month',
            'Basic template',
            'Single sheet support',
            'Basic charts'
        ]
    },
    'basic': {
        'price': 25,
        'ppt_limit': 7,
        'features': [
            '7 presentations per month',
            'Basic template',
            'Multi-sheet support (up to 5)',
            'AI-generated titles',
            'Basic charts'
        ]
    },
    'pro': {
        'price': 49,
        'ppt_limit': 15,
        'features': [
            '15 presentations per month',
            'All 10 professional templates',
            'Multi-sheet support (up to 20)',
            'AI-generated titles',
            'AI template selection',
            'Advanced charts'
        ]
    },
    'ai_pro': {
        'price': 99,
        'ppt_limit': -1,
        'features': [
            'Unlimited presentations',
            'All 10 professional templates',
            'Unlimited sheets',
            'AI-generated titles',
            'AI template selection',
            'AI-generated summaries',
            'AI data insights (5 bullets)',
            'AI layout optimization',
            'AI chart recommendations',
            'Premium support'
        ]
    }
}
```
