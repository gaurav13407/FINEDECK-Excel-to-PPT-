# Tiered Excel to PowerPoint Conversion Endpoints
# Advanced conversion with AI features based on subscription tier
# - POST /tiered-convert - Tier-based Excel to PPT conversion with AI features
# - POST /preview-ai - Preview AI recommendations before conversion
# - GET /tier-features - Get available features for user's tier
# - GET /usage-stats - Get current month's usage statistics

from fastapi import APIRouter, Depends, HTTPException, status, UploadFile, File, Form, Request
from fastapi.responses import FileResponse, JSONResponse
from typing import Optional, Dict, Any, List
from datetime import datetime
import os
import sys
import tempfile
from pathlib import Path

# Add path to find converter modules
converter_path = os.path.join(os.path.dirname(__file__), "../../../../../")
if converter_path not in sys.path:
    sys.path.insert(0, converter_path)

# Import tiered converter
from src.converter.excel_to_ppt_converter import convert_excel_to_ppt, ExcelToPPTConverter, TIER_CONFIG
from src.backend.app.services.ai_service import create_ai_service
from src.converter.excel_reader import excel_reader_all_sheets

# Import dependencies
from api.deps import get_current_active_user
from models.user import UserInDB, PLAN_CONFIGS, SubscriptionPlan
from services.file_service import get_file_by_id
from core.database import get_collection
from core.rate_limit import limiter
from bson import ObjectId

router = APIRouter()

# Helper function to map subscription plan to converter tier
def get_converter_tier(subscription_plan: SubscriptionPlan) -> str:
    """Map subscription plan to converter tier"""
    tier_mapping = {
        SubscriptionPlan.FREE: "free",
        SubscriptionPlan.BASIC: "basic",
        SubscriptionPlan.PRO: "pro",
        SubscriptionPlan.AI_PRO: "ai_pro",
        SubscriptionPlan.ENTERPRISE: "ai_pro"  # Legacy enterprise gets AI Pro features
    }
    return tier_mapping.get(subscription_plan, "free")

@router.post("/tiered-convert")
@limiter.limit("50/hour")
async def tiered_convert_excel_to_ppt(
    request: Request,
    file: UploadFile = File(...),
    template_name: Optional[str] = Form(None),
    presentation_title: Optional[str] = Form(None),
    use_finance_charts: Optional[bool] = Form(True),  # ✅ DEFAULT TO TRUE - Use Advanced Finance Charts
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Convert Excel to PowerPoint with tier-based AI features
    
    Features enabled based on subscription:
    - Free: Basic template, no AI
    - Basic: AI titles
    - Pro: AI titles + template selection
    - AI Pro: All 6 AI features (titles, summaries, insights, layout, chart recommendations)
    
    Args:
        file: Excel file to convert
        template_name: Optional template name (must be allowed for tier)
        presentation_title: Optional custom title
        use_finance_charts: Use Advanced Finance Charts with 12+ chart types (default: True)
    """
    
    # Get user's subscription tier
    user_tier = get_converter_tier(current_user.subscription.plan)
    config = PLAN_CONFIGS[current_user.subscription.plan]
    
    # DEBUG: Log received template_name
    print(f"\n🎨 ========== BACKEND TEMPLATE DEBUG ==========")
    print(f"🎨 Received template_name from frontend: {repr(template_name)}")
    print(f"🎨 Template type: {type(template_name)}")
    print(f"🎨 template_name is None: {template_name is None}")
    print(f"🎨 template_name == 'None': {template_name == 'None'}")
    print(f"🎨 template_name length: {len(template_name) if template_name else 0}")
    print(f"🎨 User tier: {user_tier}")
    print(f"🎨 Current plan: {current_user.subscription.plan}")
    print(f"🎨 ==========================================\n")
    
    # Get current month's PPT count
    users_collection = get_collection('users')
    user_obj_id = ObjectId(str(current_user.id)) if hasattr(current_user, 'id') else ObjectId(str(current_user._id))
    
    # Get or initialize usage stats
    user_doc = await users_collection.find_one({"_id": user_obj_id})
    current_month = datetime.utcnow().strftime("%Y-%m")
    
    usage_stats = user_doc.get('usage_stats', {})
    ppt_count = 0
    
    if usage_stats.get('last_reset_month') == current_month:
        ppt_count = usage_stats.get('this_month_conversions', 0)
    
    try:
        # Validate file type
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="File must be an Excel file (.xlsx or .xls)"
            )
        
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
            content = await file.read()
            tmp_excel.write(content)
            excel_path = tmp_excel.name
        
        # Create output path
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_ppt:
            output_path = tmp_ppt.name
        
        # Convert with tier-based features using professional slide builder
        # Safely extract filename without extension
        filename_parts = file.filename.rsplit('.', 1)
        default_title = filename_parts[0] if filename_parts else 'Presentation'
        
        # Use the professional converter for better styling and layouts
        converter = ExcelToPPTConverter(
            user_tier=user_tier,
            user_id=str(current_user.id),
            user_metadata={
                'name': current_user.name or current_user.email.split('@')[0],
                'email': current_user.email,
                'company': getattr(current_user, 'company', 'FinDeck User')
            },
            use_finance_charts=use_finance_charts  # Pass finance chart preference
        )
        
        # Use convert_professional for AI_PRO and PRO tiers to get nice slides
        if user_tier in ['ai_pro', 'pro']:
            print(f"\n🎨 ========== CALLING CONVERTER ==========")
            print(f"🎨 Template being passed to converter: {template_name}")
            print(f"🎨 Presentation title: {presentation_title or default_title}")
            print(f"🎨 User tier: {user_tier}")
            print(f"🎨 =====================================\n")
            
            result = converter.convert_professional(
                excel_path=excel_path,
                output_path=output_path,
                template_name=template_name,
                presentation_title=presentation_title or default_title,
                user_ppt_count=ppt_count,
                use_professional_structure=True
            )
        elif tier == "basic":
            # BASIC tier: Better than FREE - includes titles, clean charts, and proper formatting
            result = converter.convert_professional(
                excel_path=excel_path,
                output_path=output_path,
                template_name=template_name,
                presentation_title=presentation_title or default_title,
                user_ppt_count=ppt_count,
                use_professional_structure=False  # Standard structure (not full professional)
            )
        else:
            # Use basic convert for FREE tier only
            result = converter.convert(
                excel_path=excel_path,
                output_path=output_path,
                template_name=template_name,
                presentation_title=presentation_title or default_title,
                user_ppt_count=ppt_count
            )
        
        # Check if conversion was successful
        if not result['success']:
            # Clean up temp files with retry logic
            if os.path.exists(excel_path):
                try:
                    import time
                    import gc
                    gc.collect()
                    time.sleep(0.1)
                    os.unlink(excel_path)
                except PermissionError:
                    print(f"⚠️ Could not delete temp Excel file (locked): {excel_path}")
            
            if os.path.exists(output_path):
                try:
                    os.unlink(output_path)
                except PermissionError:
                    print(f"⚠️ Could not delete temp PPT file (locked): {output_path}")
            
            # Check if upgrade is required
            if result.get('upgrade_required'):
                raise HTTPException(
                    status_code=status.HTTP_402_PAYMENT_REQUIRED,
                    detail=result['error']
                )
            
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=result.get('error', 'Conversion failed')
            )
        
        # Update user statistics
        update_data = {
            "$inc": {
                "presentations_created": 1,
                "usage_stats.total_conversions": 1,
                "usage_stats.this_month_conversions": 1
            },
            "$set": {
                "updated_at": datetime.utcnow(),
                "usage_stats.last_conversion_date": datetime.utcnow(),
                "usage_stats.last_reset_month": current_month
            }
        }
        
        # Track AI usage if any AI features were used
        if result.get('ai_usage') and result['ai_usage'].get('tokens', 0) > 0:
            update_data["$inc"]["usage_stats.total_ai_tokens"] = result['ai_usage']['tokens']
            update_data["$inc"]["usage_stats.total_ai_cost"] = result['ai_usage']['cost']
        
        await users_collection.update_one(
            {"_id": user_obj_id},
            update_data
        )
        
        # Log conversion for analytics
        conversions_collection = get_collection('conversions')
        await conversions_collection.insert_one({
            "user_id": user_obj_id,
            "filename": file.filename,
            "tier": user_tier,
            "template_used": result.get('template_used'),
            "slides_created": result.get('slides_created', 0),
            "ai_features_used": result.get('ai_features_used', []),
            "ai_usage": result.get('ai_usage', {}),
            "created_at": datetime.utcnow()
        })
        
        # Clean up input file with retry logic (Windows file locking issue)
        if os.path.exists(excel_path):
            try:
                import time
                import gc
                gc.collect()  # Force garbage collection to release file handles
                time.sleep(0.1)  # Small delay to allow file handles to close
                os.unlink(excel_path)
            except PermissionError:
                # If file is still locked, schedule it for deletion later
                print(f"⚠️ Could not delete temp file immediately (file locked): {excel_path}")
                # File will be cleaned up by OS temp folder cleanup
        
        # Generate filename for download
        ppt_filename = f"{Path(file.filename).stem}_converted.pptx"
        
        # Return the file
        return FileResponse(
            path=output_path,
            filename=ppt_filename,
            media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            headers={
                "X-Slides-Created": str(result.get('slides_created', 0)),
                "X-Template-Used": result.get('template_used', 'unknown'),
                "X-AI-Features": ','.join(result.get('ai_features_used', [])),
                "X-AI-Tokens": str(result.get('ai_usage', {}).get('tokens', 0)),
                "X-AI-Cost": str(result.get('ai_usage', {}).get('cost', 0))
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        # Clean up temp files with retry logic
        if 'excel_path' in locals() and os.path.exists(excel_path):
            try:
                import time
                import gc
                gc.collect()
                time.sleep(0.1)
                os.unlink(excel_path)
            except PermissionError:
                print(f"⚠️ Could not delete temp Excel file (locked): {excel_path}")
        
        if 'output_path' in locals() and os.path.exists(output_path):
            try:
                os.unlink(output_path)
            except PermissionError:
                print(f"⚠️ Could not delete temp PPT file (locked): {output_path}")
        
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Conversion failed: {str(e)}"
        )


@router.post("/preview-ai")
@limiter.limit("20/hour")
async def preview_ai_features(
    request: Request,
    file: UploadFile = File(...),
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Preview AI recommendations without converting
    
    Returns AI suggestions for:
    - Slide titles
    - Template selection
    - Chart types
    - Layout optimization
    
    Only available for Pro and AI Pro tiers
    """
    
    user_tier = get_converter_tier(current_user.subscription.plan)
    config = PLAN_CONFIGS[current_user.subscription.plan]
    
    # Check if user has AI features
    if user_tier == "free":
        raise HTTPException(
            status_code=status.HTTP_402_PAYMENT_REQUIRED,
            detail="AI preview requires Basic, Pro, or AI Pro subscription"
        )
    
    try:
        # Validate file type
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="File must be an Excel file (.xlsx or .xls)"
            )
        
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
            content = await file.read()
            tmp_excel.write(content)
            excel_path = tmp_excel.name
        
        # Read Excel data
        sheets_dict = excel_reader_all_sheets(excel_path)
        sheets_data = [(name, df) for name, df in sheets_dict.items() if df is not None]
        
        if not sheets_data:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="No data found in Excel file"
            )
        
        # Limit sheets based on tier
        max_sheets = config.get('max_sheets', 1)
        if max_sheets > 0:
            sheets_data = sheets_data[:max_sheets]
        
        # Create AI service if tier supports it
        ai_service = None
        if len(config.get('ai_features', [])) > 0:
            ai_service = create_ai_service()
        
        preview_results = {
            "filename": file.filename,
            "sheets_found": len(sheets_data),
            "tier": user_tier,
            "ai_features_available": config.get('ai_features', []),
            "recommendations": []
        }
        
        # Generate AI recommendations for each sheet
        for sheet_name, df in sheets_data[:3]:  # Preview first 3 sheets only
            sheet_preview = {
                "sheet_name": sheet_name,
                "rows": len(df),
                "columns": len(df.columns)
            }
            
            # AI Title (Basic, Pro, AI Pro)
            if ai_service and 'title' in config.get('ai_features', []):
                try:
                    title = ai_service.generate_slide_title(df, sheet_name)
                    sheet_preview['ai_title'] = title
                except Exception as e:
                    sheet_preview['ai_title_error'] = str(e)
            
            # AI Template Selection (Pro, AI Pro)
            if ai_service and 'template_selection' in config.get('ai_features', []):
                try:
                    sheets_info = [{
                        'name': sheet_name,
                        'data_type': 'time_series',
                        'rows': len(df),
                        'cols': len(df.columns)
                    }]
                    
                    converter = ExcelToPPTConverter(user_tier)
                    allowed_templates = converter.get_allowed_templates()
                    
                    template_rec = ai_service.recommend_template(
                        file.filename,
                        sheets_info,
                        allowed_templates
                    )
                    sheet_preview['ai_template'] = template_rec
                except Exception as e:
                    sheet_preview['ai_template_error'] = str(e)
            
            # AI Chart Recommendation (AI Pro only)
            if ai_service and 'chart_type' in config.get('ai_features', []):
                try:
                    chart_rec = ai_service.recommend_chart_type(df, sheet_name)
                    sheet_preview['ai_chart'] = chart_rec
                except Exception as e:
                    sheet_preview['ai_chart_error'] = str(e)
            
            preview_results['recommendations'].append(sheet_preview)
        
        # Get AI usage stats
        if ai_service:
            preview_results['ai_usage'] = ai_service.get_usage_stats()
        
        # Clean up temp file
        if os.path.exists(excel_path):
            os.unlink(excel_path)
        
        return JSONResponse(content=preview_results)
        
    except HTTPException:
        raise
    except Exception as e:
        # Clean up temp file
        if 'excel_path' in locals() and os.path.exists(excel_path):
            os.unlink(excel_path)
        
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Preview failed: {str(e)}"
        )


@router.get("/tier-features")
async def get_tier_features(
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Get available features for user's current tier
    """
    
    user_tier = get_converter_tier(current_user.subscription.plan)
    config = PLAN_CONFIGS[current_user.subscription.plan]
    tier_config = TIER_CONFIG.get(user_tier, {})
    
    return {
        "tier": user_tier,
        "tier_name": config.get('name', 'Unknown'),
        "price": config.get('price', 0),
        "ppt_limit": config.get('presentations_limit', 0),
        "max_sheets": tier_config.get('max_sheets', 1),
        "templates_available": tier_config.get('templates', []),
        "ai_features": {
            "ai_titles": config.get('ai_titles', False),
            "ai_template_selection": config.get('ai_template_selection', False),
            "ai_summaries": config.get('ai_summaries', False),
            "ai_insights": config.get('ai_insights', False),
            "ai_layout": config.get('ai_layout', False),
            "ai_chart_recommendations": config.get('ai_chart_recommendations', False)
        },
        "features": config.get('features', []),
        "upgrade_options": get_upgrade_options(user_tier)
    }


@router.get("/usage-stats")
async def get_usage_statistics(
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Get current month's usage statistics
    """
    
    users_collection = get_collection('users')
    user_obj_id = ObjectId(str(current_user.id)) if hasattr(current_user, 'id') else ObjectId(str(current_user._id))
    
    user_doc = await users_collection.find_one({"_id": user_obj_id})
    current_month = datetime.utcnow().strftime("%Y-%m")
    
    usage_stats = user_doc.get('usage_stats', {})
    config = PLAN_CONFIGS[current_user.subscription.plan]
    
    # Reset if new month
    if usage_stats.get('last_reset_month') != current_month:
        ppt_count = 0
        ai_tokens = 0
        ai_cost = 0.0
    else:
        ppt_count = usage_stats.get('this_month_conversions', 0)
        ai_tokens = usage_stats.get('total_ai_tokens', 0)
        ai_cost = usage_stats.get('total_ai_cost', 0.0)
    
    ppt_limit = config.get('presentations_limit', 1)
    remaining = ppt_limit - ppt_count if ppt_limit > 0 else -1  # -1 means unlimited
    
    return {
        "current_month": current_month,
        "tier": get_converter_tier(current_user.subscription.plan),
        "ppt_created": ppt_count,
        "ppt_limit": ppt_limit,
        "ppt_remaining": remaining,
        "usage_percentage": (ppt_count / ppt_limit * 100) if ppt_limit > 0 else 0,
        "ai_tokens_used": ai_tokens,
        "ai_cost_total": round(ai_cost, 6),
        "total_presentations": user_doc.get('presentations_created', 0),
        "last_conversion": usage_stats.get('last_conversion_date')
    }


def get_upgrade_options(current_tier: str) -> List[Dict[str, Any]]:
    """Get upgrade options for current tier"""
    
    all_tiers = ["free", "basic", "pro", "ai_pro"]
    current_index = all_tiers.index(current_tier) if current_tier in all_tiers else 0
    
    upgrades = []
    for tier in all_tiers[current_index + 1:]:
        tier_config = TIER_CONFIG.get(tier, {})
        plan = {
            "free": SubscriptionPlan.FREE,
            "basic": SubscriptionPlan.BASIC,
            "pro": SubscriptionPlan.PRO,
            "ai_pro": SubscriptionPlan.AI_PRO
        }.get(tier)
        
        if plan:
            config = PLAN_CONFIGS[plan]
            upgrades.append({
                "tier": tier,
                "name": config.get('name'),
                "price": config.get('price'),
                "benefits": config.get('features', [])
            })
    
    return upgrades


@router.post("/test-all-tiers")
async def test_all_tiers_conversion(
    request: Request,
    file: UploadFile = File(...),
    presentation_title: Optional[str] = Form(None),
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Test endpoint: Generate presentations for ALL tiers (FREE, BASIC, PRO, AI_PRO)
    
    This creates 4 separate presentations with different features:
    - FREE: Basic 6-slide presentation
    - BASIC: 7 slides with Executive Summary
    - PRO: 7 slides with enhanced charts and templates
    - AI_PRO: 8 slides with all features including AI insights
    
    Returns a ZIP file containing all 4 presentations
    """
    import zipfile
    
    try:
        # Validate file type
        if not file.filename.lower().endswith(('.xlsx', '.xls', '.csv')):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="File must be an Excel file (.xlsx, .xls) or CSV file"
            )
        
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=Path(file.filename).suffix) as tmp_excel:
            content = await file.read()
            tmp_excel.write(content)
            excel_path = tmp_excel.name
        
        # Create temp directory for all presentations
        temp_dir = tempfile.mkdtemp()
        presentations = []
        
        # Generate presentation for each tier
        tiers_to_test = ["free", "basic", "pro", "ai_pro"]
        
        for tier in tiers_to_test:
            try:
                # Create output path for this tier
                tier_name = tier.upper().replace("_", " ")
                output_filename = f"{Path(file.filename).stem}_{tier.upper()}_Tier.pptx"
                output_path = os.path.join(temp_dir, output_filename)
                
                print(f"\n{'='*80}")
                print(f"📊 Generating {tier_name} Tier Presentation")
                print(f"{'='*80}")
                
                # Convert with tier-specific features
                result = convert_excel_to_ppt(
                    excel_path=excel_path,
                    output_path=output_path,
                    user_tier=tier,
                    template_name=None,  # Use default for each tier
                    presentation_title=presentation_title or Path(file.filename).stem,
                    user_ppt_count=0  # Testing mode, no limits
                )
                
                if result['success']:
                    presentations.append({
                        'tier': tier,
                        'path': output_path,
                        'filename': output_filename,
                        'slides': result.get('slides_created', 0),
                        'template': result.get('template_used'),
                        'ai_features': result.get('ai_features_used', [])
                    })
                    print(f"✅ {tier_name} Tier: {result.get('slides_created')} slides created")
                else:
                    print(f"❌ {tier_name} Tier: Failed - {result.get('error')}")
                    
            except Exception as e:
                print(f"❌ Error generating {tier} tier: {str(e)}")
                continue
        
        # Create ZIP file with all presentations
        zip_path = os.path.join(temp_dir, f"{Path(file.filename).stem}_ALL_TIERS.zip")
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for ppt in presentations:
                zipf.write(ppt['path'], ppt['filename'])
        
        # Clean up input file
        if os.path.exists(excel_path):
            os.unlink(excel_path)
        
        # Prepare summary response
        summary = {
            "total_presentations": len(presentations),
            "presentations": [
                {
                    "tier": p['tier'],
                    "filename": p['filename'],
                    "slides_created": p['slides'],
                    "template_used": p['template'],
                    "ai_features": p['ai_features']
                }
                for p in presentations
            ]
        }
        
        print(f"\n{'='*80}")
        print(f"✅ All Tiers Generated Successfully!")
        print(f"{'='*80}")
        print(f"📦 Total Presentations: {len(presentations)}")
        print(f"📁 ZIP File: {zip_path}")
        
        # Return the ZIP file
        return FileResponse(
            path=zip_path,
            filename=f"{Path(file.filename).stem}_ALL_TIERS.zip",
            media_type='application/zip',
            headers={
                "X-Total-Presentations": str(len(presentations)),
                "X-Summary": str(summary)
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        # Clean up temp files
        if 'excel_path' in locals() and os.path.exists(excel_path):
            os.unlink(excel_path)
        if 'temp_dir' in locals():
            import shutil
            shutil.rmtree(temp_dir, ignore_errors=True)
        
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Test conversion failed: {str(e)}"
        )


@router.get("/available-templates")
async def get_available_templates(
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Get list of available PowerPoint templates based on user's subscription tier
    
    Returns:
    - templates: List of template objects with id, name, description, category
    - tier: User's current tier
    - tier_name: Display name of tier
    - allowed_count: Number of templates allowed (or 'all')
    """
    try:
        # Get user's tier
        tier = get_converter_tier(current_user.subscription_plan)
        tier_config = TIER_CONFIG.get(tier, TIER_CONFIG['free'])
        
        # Define all available templates
        all_templates = [
            {
                "id": "corporate_blue",
                "name": "Corporate Blue",
                "description": "Professional corporate template with blue accents",
                "category": "business"
            },
            {
                "id": "modern_gradient",
                "name": "Modern Gradient",
                "description": "Contemporary design with vibrant gradients",
                "category": "modern"
            },
            {
                "id": "minimal_white",
                "name": "Minimal White",
                "description": "Clean and minimal design with white background",
                "category": "minimal"
            },
            {
                "id": "financial_pro",
                "name": "Financial Pro",
                "description": "Optimized for financial data and charts",
                "category": "finance"
            },
            {
                "id": "executive_suite",
                "name": "Executive Suite",
                "description": "Premium template for executive presentations",
                "category": "executive"
            },
            {
                "id": "tech_blue",
                "name": "Tech Blue",
                "description": "Modern template for technology companies",
                "category": "technology"
            },
            {
                "id": "creative_studio",
                "name": "Creative Studio",
                "description": "Bold and creative design for agencies",
                "category": "creative"
            },
            {
                "id": "luxury_gold",
                "name": "Luxury Gold",
                "description": "Elegant template with gold accents",
                "category": "luxury"
            },
            {
                "id": "startup_pitch",
                "name": "Startup Pitch",
                "description": "Dynamic template for startup presentations",
                "category": "modern"
            },
            {
                "id": "professional_gray",
                "name": "Professional Gray",
                "description": "Timeless gray corporate template",
                "category": "business"
            }
        ]
        
        # Get tier name mapping
        tier_names = {
            "free": "Free",
            "basic": "Basic",
            "pro": "Pro",
            "ai_pro": "AI Pro"
        }
        
        # Filter templates based on tier
        if tier == "free":
            allowed_templates = all_templates[:1]  # Only first template
            allowed_count = 1
        elif tier == "basic":
            allowed_templates = all_templates[:5]  # First 5 templates
            allowed_count = 5
        else:  # pro, ai_pro
            allowed_templates = all_templates  # All templates
            allowed_count = "all"
        
        return {
            "templates": allowed_templates,
            "tier": tier,
            "tier_name": tier_names.get(tier, "Free"),
            "allowed_count": allowed_count,
            "total_available": len(all_templates)
        }
        
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to fetch templates: {str(e)}"
        )
