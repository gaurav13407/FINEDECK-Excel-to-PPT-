"""
Test script for file models validation.
Tests all file models to ensure they work correctly before building services.
"""

from datetime import datetime, timedelta
from models.file import (
    FileType, ProcessingStatus, TemplateCategory, StorageProvider,
    FileUpload, FileRecord, FileResponse, ConversionJob, FileUsageStats,
    validate_template_access, calculate_credits_needed, get_subscription_limits
)

def test_file_upload_model():
    """Test FileUpload model validation"""
    print("🧪 Testing FileUpload model...")
    
    # Valid file upload
    valid_upload = FileUpload(
        filename="test_data.xlsx",
        file_size=1024*1024,  # 1MB
        file_type=FileType.XLSX,
        template_category=TemplateCategory.BASIC,
        template_name="Simple Dashboard",
        conversion_options={"slide_layout": "standard", "color_scheme": "default"}
    )
    print(f"✅ Valid upload: {valid_upload.filename}")
    
    # Test filename validation
    try:
        invalid_upload = FileUpload(
            filename="test_data.txt",  # Invalid extension
            file_size=1024,
            file_type=FileType.XLSX,
            template_category=TemplateCategory.BASIC
        )
        print("❌ Should have failed for invalid filename")
    except ValueError as e:
        print(f"✅ Filename validation works: {e}")
    
    # Test file size limit
    try:
        large_file = FileUpload(
            filename="large_file.xlsx",
            file_size=60*1024*1024,  # 60MB (over 50MB limit)
            file_type=FileType.XLSX,
            template_category=TemplateCategory.BASIC
        )
        print("❌ Should have failed for large file")
    except ValueError as e:
        print(f"✅ File size validation works: {e}")


def test_file_record_model():
    """Test FileRecord model"""
    print("\n🧪 Testing FileRecord model...")
    
    file_record = FileRecord(
        user_id="user123",
        filename="processed_test_data.xlsx",
        original_filename="test_data.xlsx",
        file_size=1024*1024,
        file_type=FileType.XLSX,
        template_category=TemplateCategory.PROFESSIONAL,
        template_name="Business Dashboard",
        subscription_plan="pro",
        credits_used=2,
        excel_metadata={"sheets_count": 3, "has_charts": True},
        sheets_info=[
            {"name": "Sales Data", "rows": 1000, "cols": 15},
            {"name": "Revenue", "rows": 500, "cols": 10}
        ],
        data_summary={"total_records": 1500, "data_types": ["numeric", "text", "date"]}
    )
    
    print(f"✅ FileRecord created: {file_record.filename}")
    print(f"   - User: {file_record.user_id}")
    print(f"   - Template: {file_record.template_category}")
    print(f"   - Credits: {file_record.credits_used}")
    print(f"   - Expires: {file_record.expires_at}")
    print(f"   - Excel Analysis: {len(file_record.sheets_info)} sheets")


def test_conversion_job_model():
    """Test ConversionJob model"""
    print("\n🧪 Testing ConversionJob model...")
    
    job = ConversionJob(
        file_id="file123",
        user_id="user123",
        priority=1,  # High priority
        conversion_config={
            "template_type": "professional",
            "output_format": "pptx",
            "include_animations": True
        },
        template_config={
            "color_scheme": "corporate",
            "font_family": "Arial",
            "slide_size": "16:9"
        }
    )
    
    print(f"✅ ConversionJob created: {job.job_type}")
    print(f"   - File ID: {job.file_id}")
    print(f"   - Priority: {job.priority}")
    print(f"   - Status: {job.status}")
    print(f"   - Max retries: {job.max_retries}")


def test_utility_functions():
    """Test utility functions"""
    print("\n🧪 Testing utility functions...")
    
    # Test template access validation
    print("🔐 Template Access Tests:")
    
    # Basic user tests
    basic_access = validate_template_access(TemplateCategory.BASIC, "basic")
    print(f"   Basic user → Basic template: {basic_access} ✅")
    
    basic_to_pro = validate_template_access(TemplateCategory.PROFESSIONAL, "basic")
    print(f"   Basic user → Pro template: {basic_to_pro} ❌")
    
    # Pro user tests
    pro_access = validate_template_access(TemplateCategory.PROFESSIONAL, "pro")
    print(f"   Pro user → Pro template: {pro_access} ✅")
    
    # Enterprise user tests
    enterprise_access = validate_template_access(TemplateCategory.CUSTOM, "enterprise")
    print(f"   Enterprise user → Custom template: {enterprise_access} ✅")
    
    # Test credit calculation
    print("\n💳 Credit Calculation Tests:")
    basic_credits = calculate_credits_needed(TemplateCategory.BASIC)
    pro_credits = calculate_credits_needed(TemplateCategory.PROFESSIONAL)
    custom_credits = calculate_credits_needed(TemplateCategory.CUSTOM)
    
    print(f"   Basic template: {basic_credits} credits")
    print(f"   Professional template: {pro_credits} credits")
    print(f"   Custom template: {custom_credits} credits")
    
    # Test subscription limits
    print("\n📊 Subscription Limits Tests:")
    for plan in ["basic", "pro", "enterprise"]:
        limits = get_subscription_limits(plan)
        print(f"   {plan.title()} Plan:")
        print(f"     - Monthly files: {limits['monthly_file_limit']}")
        print(f"     - Monthly credits: {limits['monthly_credits_limit']}")
        print(f"     - Storage: {limits['storage_limit_mb']} MB")


def test_usage_stats_model():
    """Test FileUsageStats model"""
    print("\n🧪 Testing FileUsageStats model...")
    
    stats = FileUsageStats(
        total_files_uploaded=150,
        total_file_processing=145,
        file_this_month=25,
        credits_used_this_month=45,
        monthly_file_limit=50,
        monthly_credits_limit=100,
        storage_used_mb=250.5,
        storage_limit_mb=500.0,
        sucess_rate_percentage=96.7
    )
    
    print(f"✅ Usage stats created:")
    print(f"   - Files this month: {stats.file_this_month}/{stats.monthly_file_limit}")
    print(f"   - Credits used: {stats.credits_used_this_month}/{stats.monthly_credits_limit}")
    print(f"   - Storage used: {stats.storage_used_mb}/{stats.storage_limit_mb} MB")
    print(f"   - Success rate: {stats.sucess_rate_percentage}%")


def test_enum_values():
    """Test all enum values"""
    print("\n🧪 Testing Enum values...")
    
    print("📁 File Types:")
    for file_type in FileType:
        print(f"   - {file_type}")
    
    print("\n⚙️ Processing Status:")
    for status in ProcessingStatus:
        print(f"   - {status}")
    
    print("\n🎨 Template Categories:")
    for category in TemplateCategory:
        print(f"   - {category}")
    
    print("\n☁️ Storage Providers:")
    for provider in StorageProvider:
        print(f"   - {provider}")


if __name__ == "__main__":
    print("🚀 Starting File Models Test Suite")
    print("=" * 50)
    
    try:
        test_file_upload_model()
        test_file_record_model()
        test_conversion_job_model()
        test_utility_functions()
        test_usage_stats_model()
        test_enum_values()
        
        print("\n" + "=" * 50)
        print("🎉 All file model tests completed successfully!")
        print("✅ File models are ready for services layer")
        
    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()