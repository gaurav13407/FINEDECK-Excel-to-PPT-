"""
Backend Integration Guide: Professional Slide Structure
========================================================

This guide shows how to integrate the new professional 7-slide structure
into your FastAPI backend.
"""

# ============================================================================
# STEP 1: Update FastAPI Endpoint
# ============================================================================

from fastapi import APIRouter, UploadFile, File, Depends
from src.converter.excel_to_ppt_converter import ExcelToPPTConverter
from src.backend.app.models.user import User
from src.backend.app.services.auth_service import get_current_user

router = APIRouter()

@router.post("/api/convert/professional")
async def convert_to_professional_ppt(
    file: UploadFile = File(...),
    presentation_title: str = None,
    current_user: User = Depends(get_current_user)
):
    """
    Convert Excel to Professional 7-Slide PPT
    
    Features:
    - Branded title slide with company info
    - AI-powered executive summary (Basic+)
    - KPI cards overview
    - Charts dashboard
    - Category comparison with AI insights (Pro+)
    - Advanced AI insights (AI Pro only)
    - Closing slide with branding
    """
    
    # Extract user metadata for branding
    user_metadata = {
        'name': current_user.full_name or current_user.email.split('@')[0],
        'company': current_user.company_name or 'Your Company',
        'email': current_user.email
    }
    
    # Create converter with user's tier
    converter = ExcelToPPTConverter(
        user_tier=current_user.subscription_tier,  # 'free', 'basic', 'pro', 'ai_pro'
        user_id=str(current_user.id),
        user_metadata=user_metadata
    )
    
    # Save uploaded file temporarily
    temp_excel_path = f"temp/uploads/{current_user.id}_{file.filename}"
    with open(temp_excel_path, "wb") as f:
        f.write(await file.read())
    
    # Generate output path
    output_filename = f"{presentation_title or file.filename.replace('.xlsx', '')}.pptx"
    output_path = f"temp/outputs/{current_user.id}_{output_filename}"
    
    # Convert with professional structure
    result = converter.convert_professional(
        excel_path=temp_excel_path,
        output_path=output_path,
        presentation_title=presentation_title or file.filename.replace('.xlsx', '').replace('_', ' ').title(),
        user_ppt_count=current_user.ppts_created_this_month
    )
    
    if result['success']:
        # Update user's PPT count
        current_user.ppts_created_this_month += 1
        await current_user.save()
        
        return {
            'success': True,
            'message': 'Professional presentation created successfully',
            'slides_created': result['slides_created'],
            'template_used': result['template_used'],
            'ai_features_used': result['ai_features_used'],
            'download_url': f"/api/download/{output_filename}"
        }
    else:
        return {
            'success': False,
            'error': result['error'],
            'upgrade_required': result.get('upgrade_required', False)
        }


# ============================================================================
# STEP 2: Update User Model
# ============================================================================

"""
Add these fields to your User model (src/backend/app/models/user.py):
"""

from pydantic import BaseModel, EmailStr
from typing import Optional

class User(BaseModel):
    id: str
    email: EmailStr
    full_name: Optional[str] = None
    company_name: Optional[str] = None  # NEW: For branding
    subscription_tier: str = 'free'  # 'free', 'basic', 'pro', 'ai_pro'
    ppts_created_this_month: int = 0
    
    # ... other fields


# ============================================================================
# STEP 3: Update Frontend Form
# ============================================================================

"""
Add company name field to user profile/registration form:

<form>
    <input type="text" name="full_name" placeholder="John Doe" required />
    <input type="email" name="email" placeholder="john@example.com" required />
    <input type="text" name="company_name" placeholder="Company Name" />  <!-- NEW -->
    <select name="subscription_tier">
        <option value="free">Free (1 PPT/month)</option>
        <option value="basic">Basic - $25/month (7 PPTs)</option>
        <option value="pro">Pro - $49/month (15 PPTs)</option>
        <option value="ai_pro">AI Pro - $99/month (Unlimited)</option>
    </select>
</form>
"""


# ============================================================================
# STEP 4: Test Endpoint
# ============================================================================

"""
Test with curl:

curl -X POST "http://localhost:8000/api/convert/professional" \
  -H "Authorization: Bearer YOUR_JWT_TOKEN" \
  -F "file=@examples/Sample_pnl.xlsx" \
  -F "presentation_title=Q4 Financial Report"

Expected Response:
{
  "success": true,
  "message": "Professional presentation created successfully",
  "slides_created": 6,
  "template_used": "minimal_white",
  "ai_features_used": ["slide_2_executive_summary"],
  "download_url": "/api/download/Q4_Financial_Report.pptx"
}
"""


# ============================================================================
# STEP 5: Feature Comparison (Show to Users)
# ============================================================================

TIER_FEATURES = {
    'free': {
        'name': 'Free',
        'price': '$0',
        'ppts_per_month': 1,
        'slides': '5-7 slides',
        'features': [
            '✅ Title slide with branding',
            '❌ No AI executive summary',
            '✅ Key metrics overview (4 KPI cards)',
            '✅ Basic charts dashboard',
            '✅ Category comparison (no AI insight)',
            '❌ No AI insights',
            '✅ Closing slide',
            '❌ No template selection'
        ]
    },
    'basic': {
        'name': 'Basic',
        'price': '$25/month',
        'ppts_per_month': 7,
        'slides': '7 slides',
        'features': [
            '✅ Title slide with branding',
            '✅ AI-powered executive summary (4-5 insights)',
            '✅ Key metrics overview (4 KPI cards)',
            '✅ Charts dashboard',
            '✅ Category comparison (no AI insight yet)',
            '❌ No advanced AI insights',
            '✅ Closing slide',
            '✅ Basic templates'
        ]
    },
    'pro': {
        'name': 'Pro',
        'price': '$49/month',
        'ppts_per_month': 15,
        'slides': '8-9 slides',
        'features': [
            '✅ Title slide with branding',
            '✅ AI-powered executive summary',
            '✅ Key metrics overview (4 KPI cards)',
            '✅ Enhanced charts dashboard',
            '✅ Category comparison with AI insight',
            '❌ No advanced AI insights',
            '✅ Closing slide',
            '✅ All 10 professional templates',
            '✅ Extra chart slides'
        ]
    },
    'ai_pro': {
        'name': 'AI Pro',
        'price': '$99/month',
        'ppts_per_month': -1,  # Unlimited
        'slides': '9-10 slides',
        'features': [
            '✅ Title slide with branding',
            '✅ AI-powered executive summary',
            '✅ Key metrics overview (4 KPI cards)',
            '✅ Advanced charts dashboard',
            '✅ Category comparison with AI insight',
            '✅ AI Insights: Top performers, Anomalies, Predictions',
            '✅ Closing slide',
            '✅ All 10 professional templates',
            '✅ Multiple extra slides',
            '✅ Advanced AI analysis'
        ]
    }
}


# ============================================================================
# STEP 6: Error Handling
# ============================================================================

"""
Common errors and solutions:
"""

ERROR_RESPONSES = {
    'limit_reached': {
        'error': 'PPT limit reached',
        'message': 'You have reached your monthly PPT limit',
        'upgrade_required': True,
        'suggestion': 'Upgrade to Basic ($25/month) for 7 PPTs or Pro ($49/month) for 15 PPTs'
    },
    'no_data': {
        'error': 'No data found in Excel file',
        'message': 'The uploaded Excel file contains no readable data',
        'upgrade_required': False,
        'suggestion': 'Please upload a valid Excel file with at least one sheet containing data'
    },
    'ai_unavailable': {
        'error': 'AI service temporarily unavailable',
        'message': 'AI features are currently unavailable, generating basic presentation',
        'upgrade_required': False,
        'suggestion': 'The presentation will be created without AI insights'
    }
}


# ============================================================================
# STEP 7: Analytics Tracking
# ============================================================================

"""
Track these metrics for each conversion:
"""

ANALYTICS_EVENTS = {
    'ppt_created': {
        'user_id': 'string',
        'tier': 'string',
        'slides_created': 'int',
        'ai_features_used': 'list',
        'template': 'string',
        'file_size_kb': 'float',
        'processing_time_ms': 'int',
        'errors': 'list'
    },
    'upgrade_prompted': {
        'user_id': 'string',
        'current_tier': 'string',
        'suggested_tier': 'string',
        'reason': 'string'  # 'limit_reached', 'feature_locked', etc.
    },
    'ai_feature_used': {
        'user_id': 'string',
        'tier': 'string',
        'feature': 'string',  # 'executive_summary', 'category_insight', 'advanced_insights'
        'tokens_used': 'int',
        'cost_usd': 'float'
    }
}


# ============================================================================
# QUICK START EXAMPLE
# ============================================================================

if __name__ == "__main__":
    """
    Quick test without backend - standalone usage
    """
    
    from src.converter.excel_to_ppt_converter import ExcelToPPTConverter
    
    # Simulate user data
    user_metadata = {
        'name': 'John Doe',
        'company': 'FinDeck Analytics Inc.',
        'email': 'john@findeck.com'
    }
    
    # Create converter (change tier to test different levels)
    converter = ExcelToPPTConverter(
        user_tier='ai_pro',  # Try: 'free', 'basic', 'pro', 'ai_pro'
        user_id='test_user_123',
        user_metadata=user_metadata
    )
    
    # Convert
    result = converter.convert_professional(
        excel_path='examples/Sample_pnl.xlsx',
        output_path='output/test_professional.pptx',
        presentation_title='Q4 Financial Performance Report',
        user_ppt_count=0  # Current month's PPT count
    )
    
    # Check result
    if result['success']:
        print(f"✅ Success!")
        print(f"   Slides: {result['slides_created']}")
        print(f"   Template: {result['template_used']}")
        print(f"   AI Features: {result['ai_features_used']}")
    else:
        print(f"❌ Error: {result['error']}")
        if result.get('upgrade_required'):
            print(f"   💡 Upgrade needed to continue")
