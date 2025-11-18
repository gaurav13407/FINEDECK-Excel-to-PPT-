# Render Deployment Guide - Redis Session Management

## 🚀 Deploying to Render with Redis Sessions

This guide explains how to enable/disable Redis session management on Render.com.

---

## Quick Setup

### Option 1: Deploy WITHOUT Redis Sessions (Recommended for Initial Deploy)

This is the **default configuration** and requires no additional setup. The app will use your existing JWT authentication.

**Steps:**
1. Push your code to GitHub
2. Render will automatically deploy
3. App will start with message: `ℹ️ Redis sessions disabled`

**No additional configuration needed!** ✅

---

### Option 2: Deploy WITH Redis Sessions (Advanced Security)

If you want enhanced session-based security, follow these steps:

#### Step 1: Set Environment Variables in Render Dashboard

Go to your Render service → **Environment** tab and add:

```bash
# Enable Redis sessions
ENABLE_REDIS_SESSIONS=true

# Upstash Redis URL (get from Upstash dashboard)
UPSTASH_REDIS_URL=rediss://default:YOUR_PASSWORD@YOUR_HOST.upstash.io:6379

# Session secret key (generate a random 64-character key)
SESSION_SECRET_KEY=your_64_character_random_key_here
```

#### Step 2: Generate SESSION_SECRET_KEY

Run this locally to generate a secure key:
```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

Copy the output and paste it as `SESSION_SECRET_KEY` in Render.

#### Step 3: Get Upstash Redis URL

1. Go to [Upstash Console](https://console.upstash.io/)
2. Create a new Redis database (if you haven't already)
3. Copy the **Redis URL** (format: `rediss://...`)
4. Paste it as `UPSTASH_REDIS_URL` in Render

#### Step 4: Deploy

1. Save environment variables in Render
2. Trigger a manual deploy or push to GitHub
3. App will start with message: `✅ Connected to Redis successfully!`
4. You'll see: `✅ Session middleware enabled - routes are protected`

---

## Environment Variables Reference

### Required for Basic Deployment (No Redis)
```bash
ENVIRONMENT=production
DATABASE_URL=your_mongodb_url
DATABASE_NAME=findeck_db
JWT_SECRET=your_jwt_secret
CORS_ORIGINS=["https://www.findeck.live","https://findeck.live"]
```

### Optional for Redis Sessions
```bash
ENABLE_REDIS_SESSIONS=true                    # Set to 'true' to enable
UPSTASH_REDIS_URL=rediss://...                # From Upstash dashboard
SESSION_SECRET_KEY=your_64_char_key           # Generated randomly
```

### Optional REST API (if using Upstash REST)
```bash
UPSTASH_REDIS_REST_URL=https://...upstash.io  # For REST API access
UPSTASH_REDIS_REST_TOKEN=your_rest_token      # For REST API access
```

---

## How to Toggle Redis Sessions

### Enable Redis Sessions
In Render dashboard:
1. Set `ENABLE_REDIS_SESSIONS=true`
2. Add `UPSTASH_REDIS_URL` and `SESSION_SECRET_KEY`
3. Redeploy

### Disable Redis Sessions
In Render dashboard:
1. Set `ENABLE_REDIS_SESSIONS=false` (or delete the variable)
2. Redeploy

**The app will automatically adapt!** No code changes needed.

---

## Troubleshooting

### Error: "UPSTASH_REDIS_URL not found"

**Solution:**
- Either set `ENABLE_REDIS_SESSIONS=false` in Render
- Or add `UPSTASH_REDIS_URL` to Render environment variables

### Warning: "Failed to connect to Redis"

**Possible causes:**
1. Invalid UPSTASH_REDIS_URL format
2. Network connectivity issues
3. Upstash database is paused/deleted

**Solution:**
- Verify URL format: `rediss://default:PASSWORD@HOST:6379` (note: `rediss` with double 's')
- Check Upstash dashboard to ensure database is active
- Test locally first with `.env` file

### Server starts but middleware is disabled

**Check logs for:**
- `⚠️ WARNING: ENABLE_REDIS_SESSIONS is true but...`
- This means a required variable is missing

**Solution:**
- Ensure all three variables are set:
  - `ENABLE_REDIS_SESSIONS=true`
  - `UPSTASH_REDIS_URL=rediss://...`
  - `SESSION_SECRET_KEY=...`

---

## Local Development

### With Redis Sessions
In your `.env` file:
```bash
ENABLE_REDIS_SESSIONS=true
UPSTASH_REDIS_URL=rediss://default:PASSWORD@HOST:6379
SESSION_SECRET_KEY=your_64_char_key
```

Run server:
```bash
python src/backend/run_server.py
```

You should see:
```
✅ Connected to Redis successfully!
✅ Session middleware enabled - routes are protected
```

### Without Redis Sessions
In your `.env` file:
```bash
ENABLE_REDIS_SESSIONS=false
# Or just comment out/remove the line
```

Run server:
```bash
python src/backend/run_server.py
```

You should see:
```
ℹ️ Redis sessions disabled (ENABLE_REDIS_SESSIONS not set to 'true')
ℹ️ Session middleware disabled - using existing auth methods
```

---

## Security Recommendations

### For Production (Render):
1. **Use HTTPS**: Render provides this automatically
2. **Strong SECRET_KEY**: Use a randomly generated 64-character key
3. **Rotate keys**: Change SESSION_SECRET_KEY periodically (invalidates all sessions)
4. **Monitor Upstash**: Check usage in Upstash dashboard

### For Free Tier:
- Upstash free tier: 10,000 commands/day
- Sufficient for small to medium apps
- Monitor usage to avoid hitting limits

---

## render.yaml Configuration

The `render.yaml` file is pre-configured with:

```yaml
# Redis sessions disabled by default
- key: ENABLE_REDIS_SESSIONS
  value: "false"

# Optional variables (sync: false means set in dashboard)
- key: UPSTASH_REDIS_URL
  sync: false
- key: SESSION_SECRET_KEY
  sync: false
```

**To enable Redis sessions:**
1. Go to Render dashboard → Environment
2. Override `ENABLE_REDIS_SESSIONS` to `"true"`
3. Add values for `UPSTASH_REDIS_URL` and `SESSION_SECRET_KEY`

---

## Testing After Deployment

### Test 1: Check Server Logs
Look for one of these messages:
```
✅ Connected to Redis successfully!
✅ Session middleware enabled
```
OR
```
ℹ️ Redis sessions disabled
ℹ️ Session middleware disabled
```

### Test 2: Test API
```bash
# Without Redis (should work)
curl https://your-app.onrender.com/docs

# With Redis (should require login for protected routes)
curl https://your-app.onrender.com/api/v1/session/protected
# Should return: 401 Unauthorized
```

---

## Cost Comparison

### Without Redis Sessions
- **Cost**: Free (uses existing JWT auth)
- **Security**: Standard JWT token-based
- **Use case**: Small projects, MVP, testing

### With Redis Sessions
- **Cost**: Free tier available (Upstash)
- **Security**: Enhanced (server-side session validation)
- **Use case**: Production apps, sensitive data, compliance requirements

---

## Next Steps

1. **Initial Deploy**: Use Option 1 (no Redis) for quick setup
2. **Add Redis Later**: Enable when needed using Option 2
3. **Monitor Usage**: Check Upstash dashboard if enabled
4. **Scale Up**: Upgrade Upstash plan when you exceed free tier

---

## Support

- **Render Docs**: https://render.com/docs
- **Upstash Docs**: https://docs.upstash.com/redis
- **FastAPI Sessions**: See `docs/REDIS_SESSION_SECURITY.md`

---

**Last Updated**: November 18, 2025
