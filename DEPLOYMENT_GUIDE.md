# End-to-End Testing & Deployment Guide

## 🚀 Complete Deployment Steps

### 1. Prerequisites

**Required Software:**
- Python 3.12+
- MongoDB 4.4+
- Node.js 16+ (for frontend)
- Git

**Required Accounts:**
- Groq API account (console.groq.com)
- MongoDB Atlas (or local MongoDB)
- Backblaze B2 (for file storage)

### 2. Backend Setup

```bash
# Navigate to backend directory
cd "src/backend"

# Create virtual environment
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Install additional packages for tiered converter
pip install groq python-pptx pandas numpy openpyxl
```

### 3. Environment Configuration

Create `.env` file in `src/backend/app/`:

```env
# MongoDB
MONGODB_URL=mongodb://localhost:27017
DATABASE_NAME=findeck

# JWT Authentication
SECRET_KEY=your-secret-key-here-change-in-production
ALGORITHM=HS256
ACCESS_TOKEN_EXPIRE_MINUTES=30

# Groq API
GROQ_API_KEY=your-groq-api-key-here

# Backblaze B2
B2_KEY_ID=your-b2-key-id
B2_APPLICATION_KEY=your-b2-application-key
B2_BUCKET_NAME=your-bucket-name

# Server
HOST=0.0.0.0
PORT=8000
ENVIRONMENT=development
```

### 4. Database Setup

**Create MongoDB indexes:**

```javascript
// Connect to MongoDB
use findeck;

// Users collection indexes
db.users.createIndex({ "email": 1 }, { unique: true });
db.users.createIndex({ "subscription.plan": 1 });
db.users.createIndex({ "usage_stats.last_reset_month": 1 });

// Conversions collection (for analytics)
db.conversions.createIndex({ "user_id": 1 });
db.conversions.createIndex({ "created_at": -1 });
db.conversions.createIndex({ "tier": 1 });

// Files collection
db.files.createIndex({ "user_id": 1 });
db.files.createIndex({ "created_at": -1 });
```

**Add usage_stats field to existing users:**

```javascript
db.users.updateMany(
  { "usage_stats": { $exists: false } },
  {
    $set: {
      "usage_stats": {
        "total_conversions": 0,
        "this_month_conversions": 0,
        "last_conversion_date": null,
        "last_reset_month": null,
        "total_ai_tokens": 0,
        "total_ai_cost": 0
      }
    }
  }
);
```

### 5. Start Backend Server

```bash
# Navigate to backend app directory
cd src/backend/app

# Start with uvicorn
uvicorn main:app --reload --host 0.0.0.0 --port 8000

# Or for production:
uvicorn main:app --workers 4 --host 0.0.0.0 --port 8000
```

**Verify server is running:**
- API Docs: http://localhost:8000/docs
- Health Check: http://localhost:8000/health

### 6. Test API Endpoints

**Create test users with different tiers:**

```bash
# Free tier user
curl -X POST http://localhost:8000/api/v1/auth/register \
  -H "Content-Type: application/json" \
  -d '{
    "name": "Free User",
    "email": "free@test.com",
    "password": "password123"
  }'

# Basic tier user
curl -X POST http://localhost:8000/api/v1/auth/register \
  -H "Content-Type: application/json" \
  -d '{
    "name": "Basic User",
    "email": "basic@test.com",
    "password": "password123"
  }'

# Pro tier user
curl -X POST http://localhost:8000/api/v1/auth/register \
  -H "Content-Type: application/json" \
  -d '{
    "name": "Pro User",
    "email": "pro@test.com",
    "password": "password123"
  }'

# AI Pro tier user
curl -X POST http://localhost:8000/api/v1/auth/register \
  -H "Content-Type: application/json" \
  -d '{
    "name": "AI Pro User",
    "email": "aipro@test.com",
    "password": "password123"
  }'
```

**Update subscription tiers in MongoDB:**

```javascript
// Free tier (default - no change needed)

// Basic tier
db.users.updateOne(
  { email: "basic@test.com" },
  {
    $set: {
      "subscription.plan": "basic",
      "subscription.price_per_month": 25.00,
      "subscription.presentations_limit": 7,
      "subscription.monthly_credits_limit": 10
    }
  }
);

// Pro tier
db.users.updateOne(
  { email: "pro@test.com" },
  {
    $set: {
      "subscription.plan": "pro",
      "subscription.price_per_month": 49.00,
      "subscription.presentations_limit": 15,
      "subscription.monthly_credits_limit": 50
    }
  }
);

// AI Pro tier
db.users.updateOne(
  { email: "aipro@test.com" },
  {
    $set: {
      "subscription.plan": "ai_pro",
      "subscription.price_per_month": 99.00,
      "subscription.presentations_limit": -1,
      "subscription.monthly_credits_limit": 1000,
      "subscription.ai_features_enabled": true
    }
  }
);
```

**Login and get tokens:**

```bash
# Login free user
curl -X POST http://localhost:8000/api/v1/auth/login \
  -H "Content-Type: application/json" \
  -d '{
    "email": "free@test.com",
    "password": "password123"
  }'

# Save the access_token from response
```

### 7. Test Tiered Conversion

**Test Free Tier:**

```bash
TOKEN="your-free-user-token"

# Get tier features
curl -X GET http://localhost:8000/api/v1/tiered/tier-features \
  -H "Authorization: Bearer $TOKEN"

# Get usage stats
curl -X GET http://localhost:8000/api/v1/tiered/usage-stats \
  -H "Authorization: Bearer $TOKEN"

# Convert Excel to PPT
curl -X POST http://localhost:8000/api/v1/tiered/tiered-convert \
  -H "Authorization: Bearer $TOKEN" \
  -F "file=@examples/Sample_pnl.xlsx" \
  -F "presentation_title=Free Tier Test" \
  -o free_output.pptx
```

**Test Basic Tier:**

```bash
TOKEN="your-basic-user-token"

# Preview AI features
curl -X POST http://localhost:8000/api/v1/tiered/preview-ai \
  -H "Authorization: Bearer $TOKEN" \
  -F "file=@examples/Sample_pnl.xlsx"

# Convert with AI titles
curl -X POST http://localhost:8000/api/v1/tiered/tiered-convert \
  -H "Authorization: Bearer $TOKEN" \
  -F "file=@examples/Sample_pnl.xlsx" \
  -F "presentation_title=Basic Tier Test" \
  -o basic_output.pptx
```

**Test Pro Tier:**

```bash
TOKEN="your-pro-user-token"

# Convert with AI titles + template selection
curl -X POST http://localhost:8000/api/v1/tiered/tiered-convert \
  -H "Authorization: Bearer $TOKEN" \
  -F "file=@examples/Portfolio Allocation Data.xlsx" \
  -F "presentation_title=Pro Tier Test" \
  -o pro_output.pptx
```

**Test AI Pro Tier:**

```bash
TOKEN="your-ai-pro-user-token"

# Full AI features
curl -X POST http://localhost:8000/api/v1/tiered/tiered-convert \
  -H "Authorization: Bearer $TOKEN" \
  -F "file=@examples/Risk Metrics Data.xlsx" \
  -F "presentation_title=AI Pro Tier Test" \
  -o ai_pro_output.pptx
```

### 8. Frontend Integration

**Example React/Next.js code:**

```typescript
// app/convert/page.tsx
'use client';

import { useState } from 'react';
import { useAuth } from '@/hooks/useAuth';

export default function ConvertPage() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const { user, token } = useAuth();

  const handleConvert = async () => {
    if (!file) return;

    setLoading(true);
    const formData = new FormData();
    formData.append('file', file);
    formData.append('presentation_title', file.name.split('.')[0]);

    try {
      const response = await fetch('/api/v1/tiered/tiered-convert', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`
        },
        body: formData
      });

      if (response.status === 402) {
        // Show upgrade modal
        const data = await response.json();
        alert(`Upgrade Required: ${data.detail}`);
        return;
      }

      if (!response.ok) {
        throw new Error('Conversion failed');
      }

      // Download file
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = file.name.replace('.xlsx', '.pptx');
      a.click();

      // Get usage from headers
      const slidesCreated = response.headers.get('X-Slides-Created');
      const aiFeatures = response.headers.get('X-AI-Features');
      
      alert(`Success! Created ${slidesCreated} slides with features: ${aiFeatures}`);

    } catch (error) {
      console.error('Conversion error:', error);
      alert('Conversion failed');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="container mx-auto p-8">
      <h1 className="text-3xl font-bold mb-8">Convert Excel to PowerPoint</h1>
      
      <div className="mb-4">
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => setFile(e.target.files?.[0] || null)}
          className="border p-2 rounded"
        />
      </div>

      <button
        onClick={handleConvert}
        disabled={!file || loading}
        className="bg-blue-500 text-white px-6 py-2 rounded hover:bg-blue-600 disabled:bg-gray-400"
      >
        {loading ? 'Converting...' : 'Convert to PPT'}
      </button>

      {user && (
        <div className="mt-8 p-4 bg-gray-100 rounded">
          <h2 className="font-bold mb-2">Your Plan: {user.subscription.plan}</h2>
          <p>PPT Limit: {user.subscription.presentations_limit === -1 ? 'Unlimited' : user.subscription.presentations_limit}</p>
        </div>
      )}
    </div>
  );
}
```

### 9. Monitoring & Analytics

**Create analytics dashboard:**

```javascript
// Get conversion statistics
db.conversions.aggregate([
  {
    $group: {
      _id: "$tier",
      total_conversions: { $sum: 1 },
      total_slides: { $sum: "$slides_created" },
      total_ai_tokens: { $sum: "$ai_usage.tokens" },
      total_ai_cost: { $sum: "$ai_usage.cost" },
      avg_slides_per_ppt: { $avg: "$slides_created" }
    }
  },
  {
    $sort: { total_conversions: -1 }
  }
]);

// Monthly revenue calculation
db.users.aggregate([
  {
    $match: {
      "subscription.status": "active"
    }
  },
  {
    $group: {
      _id: "$subscription.plan",
      count: { $sum: 1 },
      monthly_revenue: {
        $sum: "$subscription.price_per_month"
      }
    }
  }
]);
```

### 10. Production Deployment

**Update `.env` for production:**

```env
ENVIRONMENT=production
SECRET_KEY=generate-secure-random-key
MONGODB_URL=mongodb+srv://user:pass@cluster.mongodb.net/findeck
GROQ_API_KEY=your-production-groq-key
```

**Deploy with Docker:**

```dockerfile
# Dockerfile
FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "4"]
```

```bash
# Build and run
docker build -t findeck-backend .
docker run -p 8000:8000 --env-file .env findeck-backend
```

**Deploy to cloud:**

```bash
# Heroku
heroku create findeck-api
git push heroku main

# DigitalOcean App Platform
doctl apps create --spec app.yaml

# AWS Elastic Beanstalk
eb init -p python-3.12 findeck-api
eb create findeck-prod
```

### 11. Monitoring

**Setup monitoring:**

```python
# Add to main.py
from prometheus_client import Counter, Histogram
import time

# Metrics
conversions_total = Counter('conversions_total', 'Total conversions', ['tier'])
conversion_duration = Histogram('conversion_duration_seconds', 'Conversion duration')
ai_tokens_used = Counter('ai_tokens_total', 'Total AI tokens', ['tier'])

@app.middleware("http")
async def add_metrics(request: Request, call_next):
    start_time = time.time()
    response = await call_next(request)
    duration = time.time() - start_time
    
    if request.url.path.startswith("/api/v1/tiered/tiered-convert"):
        conversion_duration.observe(duration)
    
    return response
```

### 12. Backup & Recovery

**MongoDB backup script:**

```bash
#!/bin/bash
# backup.sh

DATE=$(date +%Y%m%d_%H%M%S)
BACKUP_DIR="backups"

mkdir -p $BACKUP_DIR

# Backup MongoDB
mongodump --uri="$MONGODB_URL" --out="$BACKUP_DIR/mongodb_$DATE"

# Backup files
tar -czf "$BACKUP_DIR/files_$DATE.tar.gz" user_uploads/

echo "Backup completed: $DATE"
```

## 🎯 Testing Checklist

### Pre-Deployment Tests

- [ ] All 4 tiers working (Free, Basic, Pro, AI Pro)
- [ ] PPT limits enforced correctly
- [ ] Template restrictions working
- [ ] AI features distributed correctly
- [ ] Usage tracking accurate
- [ ] Monthly reset working
- [ ] Authentication working
- [ ] File upload/download working
- [ ] Error handling proper
- [ ] Rate limiting working

### Load Testing

```bash
# Install Apache Bench
sudo apt-get install apache2-utils

# Test conversion endpoint (100 requests, 10 concurrent)
ab -n 100 -c 10 -T 'multipart/form-data' \
  -H "Authorization: Bearer $TOKEN" \
  http://localhost:8000/api/v1/tiered/tiered-convert
```

### Security Tests

- [ ] SQL injection prevention
- [ ] XSS prevention
- [ ] CSRF protection
- [ ] Rate limiting
- [ ] JWT expiration
- [ ] File upload validation
- [ ] Input sanitization

## 🚨 Common Issues

### Issue: "Module not found: converter"
**Solution:** Add converter path to Python path in endpoint files

### Issue: "AI service initialization failed"
**Solution:** Check GROQ_API_KEY in .env file

### Issue: "PPT limit not enforcing"
**Solution:** Ensure usage_stats.last_reset_month is set correctly

### Issue: "Templates not found"
**Solution:** Check src/templates/built_in/ folder exists with template JSON files

## 📊 Success Metrics

**Track these KPIs:**
- Daily Active Users (DAU)
- Monthly Recurring Revenue (MRR)
- Conversion Rate (free → paid)
- Churn Rate
- Average slides per PPT
- AI feature usage by tier
- API response times
- Error rates

**Target Metrics:**
- Conversion success rate: > 99%
- API response time: < 5 seconds
- AI accuracy: > 95% user satisfaction
- Uptime: 99.9%

## 🎉 Launch Checklist

- [ ] Backend deployed and tested
- [ ] Frontend connected to backend
- [ ] All 4 tiers tested end-to-end
- [ ] Payment integration complete (Stripe)
- [ ] Email notifications setup
- [ ] Analytics tracking configured
- [ ] Error monitoring (Sentry)
- [ ] Backup system running
- [ ] Documentation complete
- [ ] Legal pages (Terms, Privacy)
- [ ] Marketing site live
- [ ] Support system ready

**You're now ready to launch! 🚀**
