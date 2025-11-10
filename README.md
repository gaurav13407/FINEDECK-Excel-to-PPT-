"# 📊 FinDeck - Excel to PowerPoint Converter

> **Professional SaaS platform for converting Excel financial data into polished PowerPoint presentations**

[![Live Demo](https://img.shields.io/badge/demo-findeck.live-blue)](https://www.findeck.live)
[![Python](https://img.shields.io/badge/python-3.11+-green)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.115.0-teal)](https://fastapi.tiangolo.com/)
[![MongoDB](https://img.shields.io/badge/MongoDB-Atlas-darkgreen)](https://www.mongodb.com/)

---

## 🚀 Features

- **📈 Excel Parsing**: Intelligent data extraction from `.xlsx` and `.xls` files
- **🎨 PPT Generation**: Automated creation of professional presentations using python-pptx
- **� Advanced Finance Charts**: 12+ chart types with intelligent detection
  - **Performance & Growth**: Line, Area, Column, Bar charts
  - **Portfolio & Assets**: Pie, Donut, Stacked Column charts
  - **P&L & Cash Flow**: Waterfall, Stacked Bar charts
  - **Market Analysis**: Candlestick, Scatter, Bubble charts
  - **Smart Detection**: Auto-selects best chart type based on data structure
- **�💳 Subscription Tiers**: Free, Basic, Pro, and AI (Enterprise) plans with credit system
- **🔐 Authentication**: JWT tokens + OAuth (GitHub, Google, Twitter)
- **☁️ Cloud Storage**: Multi-backend support (Azure Blob Storage, Backblaze B2)
- **⚡ Async Processing**: FastAPI + Motor for high-performance async operations
- **🛡️ Rate Limiting**: Per-user rate limits based on subscription tier
- **📊 Usage Tracking**: Real-time credit consumption and presentation count tracking

---

## 📋 Table of Contents

- [Tech Stack](#-tech-stack)
- [Architecture](#-architecture)
- [Installation](#-installation)
- [Configuration](#-configuration)
- [Database Setup](#-database-setup)
- [Running the Application](#-running-the-application)
- [API Documentation](#-api-documentation)
- [Deployment](#-deployment)
- [Project Structure](#-project-structure)
- [Contributing](#-contributing)

---

## 🛠 Tech Stack

### Backend
- **FastAPI 0.115.0** - Modern async web framework
- **Motor 3.4.0** - Async MongoDB driver
- **Pydantic v2** - Data validation with type hints
- **PyJWT** - JSON Web Token authentication
- **Redis + RQ** - Background job processing
- **APScheduler** - Scheduled task management

### Data Processing
- **pandas 2.2.2** - Excel data manipulation
- **openpyxl 3.1.5** - Excel file reading
- **python-pptx 1.0.2** - PowerPoint generation
- **matplotlib/plotly** - Chart generation

### Storage
- **Azure Blob Storage 12.20.0** - Primary cloud storage
- **Backblaze B2 (b2sdk 1.17.4)** - Secondary cloud storage
- **Local filesystem** - Fallback storage

### Frontend
- **Vanilla JavaScript (ES6+)** - No framework overhead
- **Static HTML/CSS** - Netlify deployment
- **Custom API client** - RESTful communication

### Database
- **MongoDB Atlas** - Cloud-hosted document database

---

## 🏗 Architecture

```
┌─────────────────┐         ┌──────────────────┐         ┌─────────────────┐
│   Frontend      │         │   FastAPI API    │         │    MongoDB      │
│  (Netlify)      │◄───────►│   (Backend)      │◄───────►│    Atlas        │
│                 │  REST   │                  │  Motor  │                 │
└─────────────────┘         └──────────────────┘         └─────────────────┘
                                     │                             
                                     │                             
                            ┌────────▼────────┐                   
                            │  Cloud Storage  │                   
                            │  Azure/B2/Local │                   
                            └─────────────────┘                   
```

### Key Components

1. **Authentication Layer**: JWT + OAuth with secure token management
2. **Service Layer**: Business logic for users, files, conversions
3. **Storage Layer**: Multi-backend cloud storage with fallback
4. **Converter Engine**: Excel → DataFrame → PPT pipeline
5. **Rate Limiter**: Per-user throttling based on subscription tier

---

## 💻 Installation

### Prerequisites

- **Python 3.11+** ([Download](https://www.python.org/downloads/))
- **MongoDB** (local or Atlas account)
- **Redis** (optional, for background jobs)
- **Git**

### 1. Clone Repository

```bash
git clone https://github.com/gaurav13407/FINEDECK-Excel-to-PPT-.git
cd FINEDECK-Excel-to-PPT-
```

### 2. Create Virtual Environment

**Windows:**
```cmd
python -m venv .venv
.venv\Scripts\activate
```

**macOS/Linux:**
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Install SlowAPI (Rate Limiting)

```bash
pip install slowapi==0.1.9
```

---

## ⚙️ Configuration

### Environment Variables

Create a `.env` file in the project root:

```properties
# Database Configuration
DATABASE_URL=mongodb+srv://user:password@cluster.mongodb.net/?retryWrites=true&w=majority
DATABASE_NAME=findeck_db

# JWT Security (GENERATE NEW SECRET!)
JWT_SECRET=<use-python-secrets-token_urlsafe-to-generate>
JWT_ALGORITHM=HS256
JWT_EXPIRATION_MINUTES=1440

# App Settings
DEBUG=false
ENVIRONMENT=production
APP_NAME=FinDeck API
API_VERSION=v1

# File Upload Settings
MAX_FILE_SIZE_MB=50
ALLOWED_FILE_EXTENSIONS=[".xlsx",".xls"]

# CORS Origins
CORS_ORIGINS=["https://www.findeck.live","https://findeck.live","http://localhost:3000"]

# Backblaze B2 Storage
USE_B2_STORAGE=true
B2_APPLICATION_KEY_ID=<your-b2-key-id>
B2_APPLICATION_KEY=<your-b2-secret-key>
B2_BUCKET_NAME=<your-bucket-name>
B2_ENDPOINT=s3.us-east-005.backblazeb2.com

# Azure Storage (Alternative)
AZURE_STORAGE_CONNECTION_STRING=<your-azure-connection-string>
AZURE_CONTAINER_NAME=user-files

# Email Service (Brevo)
BREVO_API_KEY=<your-brevo-api-key>
BREVO_TEMPLATE_ID=1
EMAIL_FROM=no-reply@findeck.live
MAIL_FROM_NAME=FinDeck

# OAuth Providers
GOOGLE_CLIENT_ID=<your-google-client-id>
GOOGLE_CLIENT_SECRET=<your-google-client-secret>
```

### Generate Secure JWT Secret

```python
import secrets
print(secrets.token_urlsafe(64))
```

---

## 🗄 Database Setup

### MongoDB Atlas (Recommended)

1. Create account at [mongodb.com/cloud/atlas](https://www.mongodb.com/cloud/atlas)
2. Create a new cluster (free tier available)
3. Create database user with read/write access
4. Whitelist your IP address (or `0.0.0.0/0` for development)
5. Copy connection string to `.env`

### Local MongoDB (Alternative)

```bash
# Install MongoDB Community Edition
# Windows: https://www.mongodb.com/try/download/community
# macOS: brew install mongodb-community

# Start MongoDB service
mongod --dbpath ./data/db
```

### Database Migration

Run the subscription normalization script:

```bash
# Preview changes (dry run)
python scripts/migrate_subscriptions.py --dry-run

# Apply changes
python scripts/migrate_subscriptions.py
```

---

## 🚀 Running the Application

### Backend (FastAPI)

**Development Mode:**
```bash
cd src/backend/app
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

**Production Mode:**
```bash
uvicorn src.backend.app.main:app --host 0.0.0.0 --port 8000 --workers 4
```

### Frontend (Static Files)

**Local Development:**
```bash
# Using Python's built-in server
cd src/ui
python -m http.server 5500
```

**Live Server (VS Code Extension):**
- Install "Live Server" extension
- Right-click `src/ui/index.html` → "Open with Live Server"

### Access Points

- **Frontend**: http://localhost:5500
- **Backend API**: http://localhost:8000
- **API Documentation**: http://localhost:8000/api/docs
- **ReDoc**: http://localhost:8000/api/redoc

---

## 📖 API Documentation

### Interactive Swagger UI

Visit: `http://localhost:8000/api/docs`

### Key Endpoints

#### Authentication
```http
POST /api/v1/auth/register
POST /api/v1/auth/login
POST /api/v1/auth/refresh
```

#### File Management
```http
POST /api/v1/files/upload
GET /api/v1/files/
GET /api/v1/files/{file_id}
DELETE /api/v1/files/{file_id}
```

#### Conversions
```http
POST /api/v1/conversions/convert
GET /api/v1/conversions/templates
```

#### User Management
```http
GET /api/v1/users/me
GET /api/v1/users/stats
GET /api/v1/users/subscription
```

### Rate Limits (Per Hour)

| Tier | Conversions/Hour | Files Upload/Month |
|------|------------------|-------------------|
| Free | 5 | 1 |
| Basic | 15 | 5 |
| Pro | 50 | 15 |
| AI (Enterprise) | 200 | 100 |

---

## 🌐 Deployment

### Backend (Railway/Render/Heroku)

1. **Set environment variables** in hosting platform dashboard
2. **Update `Procfile`** (if needed):
   ```
   web: uvicorn src.backend.app.main:app --host 0.0.0.0 --port $PORT
   ```
3. **Deploy** via Git push or GitHub integration

### Frontend (Netlify)

1. Connect repository to Netlify
2. **Build settings**:
   - Base directory: `src/ui`
   - Publish directory: `.` (no build needed)
3. **Custom domain**: Configure DNS to point to Netlify

### Database (MongoDB Atlas)

- Already cloud-hosted
- Enable IP whitelist for production servers
- Create backup schedule

---

## 📁 Project Structure

```
FinDeck/
├── src/
│   ├── backend/
│   │   └── app/
│   │       ├── api/
│   │       │   └── v1/
│   │       │       └── endpoints/
│   │       │           ├── users.py
│   │       │           ├── auth.py
│   │       │           ├── files.py
│   │       │           └── conversions.py
│   │       ├── core/
│   │       │   ├── config.py
│   │       │   ├── security.py
│   │       │   ├── database.py
│   │       │   └── rate_limit.py
│   │       ├── models/
│   │       │   ├── user.py
│   │       │   └── file.py
│   │       ├── services/
│   │       │   ├── user_service.py
│   │       │   └── file_service.py
│   │       ├── storage/
│   │       │   └── b2_storage.py
│   │       └── main.py
│   ├── converter/
│   │   ├── excel_reader.py
│   │   └── ppt_writer.py
│   └── ui/
│       ├── index.html
│       ├── mainpage.html
│       └── assets/
│           ├── css/
│           └── js/
├── scripts/
│   └── migrate_subscriptions.py
├── tests/
│   ├── test_auth.py
│   ├── test_files_api.py
│   └── test_conversions.py
├── examples/
│   └── *.xlsx (sample files)
├── .env
├── requirements.txt
└── README.md
```

---

## 🧪 Testing

### Run All Tests

```bash
pytest tests/ -v
```

### Run Specific Test Suite

```bash
pytest tests/test_auth.py -v
pytest tests/test_conversions.py -v
```

### Coverage Report

```bash
pytest --cov=src/backend/app --cov-report=html
```

---

## 🤝 Contributing

1. Fork the repository
2. Create feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add AmazingFeature'`)
4. Push to branch (`git push origin feature/AmazingFeature`)
5. Open Pull Request

---

## 📄 License

Proprietary - © 2025 FinDeck. All rights reserved.

---

## 📧 Contact

- **Website**: [www.findeck.live](https://www.findeck.live)
- **Support**: support@findeck.live
- **GitHub**: [@gaurav13407](https://github.com/gaurav13407)

---

## 🙏 Acknowledgments

- [FastAPI](https://fastapi.tiangolo.com/) - Modern web framework
- [python-pptx](https://python-pptx.readthedocs.io/) - PowerPoint generation
- [MongoDB](https://www.mongodb.com/) - Database platform
- [Netlify](https://www.netlify.com/) - Frontend hosting

---

**Made with ❤️ by the FinDeck Team**" 
