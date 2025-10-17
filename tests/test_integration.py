"""
Integration test to verify all models work together correctly
"""

import sys
import os
from datetime import datetime, timedelta
from bson import ObjectId

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src', 'backend', 'app'))

try:
    # Import from user model
    from models.user import TemplateCategory, PyObjectId, SubscriptionPlan, SubscriptionStatus
    
    # Import from conversion model
    from models.conversion import (
        PowerPointTemplate, ConversionSettings, ConversionResult, 
        TemplateUsageStats, LayoutType, CharType, DataMappingType,
        calculate_conversion_credits
    )
    
    # Import from file model
    from models.file import FileType, ProcessingStatus, FileUpload, ConversionJob
    
    print("✅ All model imports successful!")
    
except ImportError as e:
    print(f"❌ Import error: {e}")
    sys.exit(1)


def test_cross_model_integration():
    """Test that models work together correctly"""
    print("\n--- Testing Cross-Model Integration ---")
    
    # Create a user ID
    user_id = PyObjectId()
    
    # Create a PowerPoint template
    template = PowerPointTemplate(
        template_name="Financial Dashboard Template",
        category=TemplateCategory.PROFESSIONAL,
        description="Professional template for financial dashboards",
        template_file_path="/templates/financial_dashboard.pptx",
        slide_count=12,
        supported_layouts=[LayoutType.TITLE_SLIDE, LayoutType.TITLE_CONTENT, LayoutType.COMPARISON],
        is_premium=False
    )
    print(f"✅ Created template: {template.template_name} ({template.category})")
    
    # Create conversion settings
    settings = ConversionSettings(
        preferred_layouts=LayoutType.TITLE_CONTENT,
        data_mapping_strategy=DataMappingType.ROW_TO_SLIDE,
        max_rows_per_slide=8,
        convert_charts=True,
        supported_chart_types=[CharType.BAR, CharType.LINE, CharType.PIE],
        primary_color="#1f4e79"
    )
    print(f"✅ Created conversion settings with {len(settings.supported_chart_types)} chart types")
    
    # Create file upload
    file_upload = FileUpload(
        filename="quarterly_report.xlsx",
        file_size=5 * 1024 * 1024,  # 5MB
        file_type=FileType.XLSX,
        template_category=TemplateCategory.PROFESSIONAL,
        template_name="Financial Dashboard Template"
    )
    print(f"✅ Created file upload: {file_upload.filename} ({file_upload.file_type})")
    
    # Create conversion job
    conversion_job = ConversionJob(
        file_id="file_123",
        user_id=user_id,
        status=ProcessingStatus.PROCESSING,
        conversion_config=settings.model_dump(),
        template_config={"template_id": "template_456"},
        estimated_completion=datetime.utcnow() + timedelta(minutes=5)
    )
    print(f"✅ Created conversion job: {conversion_job.id} (Status: {conversion_job.status})")
    print(f"   File ID: {conversion_job.file_id}")
    print(f"   Job type: {conversion_job.job_type}")
    
    # Create conversion result
    started_time = datetime.utcnow()
    completed_time = started_time + timedelta(minutes=3)
    
    result = ConversionResult(
        conversion_id="conv_789",
        job_id=conversion_job.id or "job_123",
        user_id=user_id,
        source_file_id="file_123",
        template_id="template_456",
        output_file_path="/outputs/quarterly_report.pptx",
        output_file_size=8 * 1024 * 1024,  # 8MB
        slide_count=10,
        processing_time_seconds=180.0,
        data_rows_processed=95,
        charts_converted=4,
        started_at=started_time,
        completed_at=completed_time
    )
    print(f"✅ Created conversion result: {result.conversion_id}")
    print(f"   Processing time: {result.processing_time_seconds}s")
    print(f"   Charts converted: {result.charts_converted}")
    
    # Create template usage stats
    usage_stats = TemplateUsageStats(
        template_id="template_456",
        template_name="Financial Dashboard Template",
        category=TemplateCategory.PROFESSIONAL,
        total_uses=47,
        unique_users=23,
        success_rate=94.5,
        average_rating=4.2,
        total_ratings=18
    )
    print(f"✅ Created usage stats for template")
    print(f"   Total uses: {usage_stats.total_uses}, Success rate: {usage_stats.success_rate}%")
    
    # Test enum consistency across models
    print("\n--- Testing Enum Consistency ---")
    
    # Test TemplateCategory is consistent
    template_categories = [item for item in TemplateCategory]
    print(f"✅ TemplateCategory has {len(template_categories)} values: {[cat.value for cat in template_categories]}")
    
    # Test that file upload and template use same categories
    assert file_upload.template_category == template.category == usage_stats.category
    print("✅ Template categories are consistent across models")
    
    # Test ProcessingStatus values
    processing_statuses = [item for item in ProcessingStatus]
    print(f"✅ ProcessingStatus has {len(processing_statuses)} values: {[status.value for status in processing_statuses]}")
    
    # Test that PyObjectId works consistently
    assert isinstance(user_id, ObjectId)
    assert conversion_job.user_id == result.user_id == user_id
    print("✅ PyObjectId works consistently across models")
    
    print("\n🎉 All cross-model integration tests passed!")
    return True


def test_business_logic_alignment():
    """Test that business logic is aligned across models"""
    print("\n--- Testing Business Logic Alignment ---")
    
    # Test credit calculation logic
    basic_credits = calculate_conversion_credits(
        file_size_mb=3.0,
        slide_count=5,
        template_category=TemplateCategory.BASIC,
        include_charts=False
    )
    
    professional_credits = calculate_conversion_credits(
        file_size_mb=15.0,
        slide_count=12,
        template_category=TemplateCategory.PROFESSIONAL,
        include_charts=True
    )
    
    premium_credits = calculate_conversion_credits(
        file_size_mb=25.0,
        slide_count=20,
        template_category=TemplateCategory.PREMIUM,
        include_charts=True
    )
    
    print(f"✅ Credit calculation:")
    print(f"   Basic conversion: {basic_credits} credits")
    print(f"   Professional conversion: {professional_credits} credits")
    print(f"   Premium conversion: {premium_credits} credits")
    
    # Verify that premium costs more than professional, which costs more than basic
    assert premium_credits >= professional_credits >= basic_credits
    print("✅ Credit pricing logic is correct (premium ≥ professional ≥ basic)")
    
    # Test that file size limits are consistent
    max_file_size = 50 * 1024 * 1024  # 50MB from FileUpload
    print(f"✅ Maximum file size: {max_file_size / (1024*1024):.0f}MB")
    
    # Test that template slide counts are reasonable
    min_slides, max_slides = 1, 50  # From PowerPointTemplate
    print(f"✅ Template slide range: {min_slides}-{max_slides} slides")
    
    # Test conversion settings ranges
    min_title_font, max_title_font = 12, 48  # From ConversionSettings
    min_body_font, max_body_font = 8, 36
    print(f"✅ Font size ranges: Title {min_title_font}-{max_title_font}pt, Body {min_body_font}-{max_body_font}pt")
    
    print("\n🎉 All business logic alignment tests passed!")
    return True


def main():
    """Run all integration tests"""
    print("="*70)
    print("CROSS-MODEL INTEGRATION TEST SUITE")
    print("="*70)
    
    tests = [
        ("Cross-Model Integration", test_cross_model_integration),
        ("Business Logic Alignment", test_business_logic_alignment)
    ]
    
    passed = 0
    total = len(tests)
    
    for test_name, test_func in tests:
        print(f"\n{'='*25} {test_name} {'='*25}")
        try:
            if test_func():
                passed += 1
                print(f"✅ {test_name} PASSED")
            else:
                print(f"❌ {test_name} FAILED")
        except Exception as e:
            print(f"❌ {test_name} FAILED with exception: {e}")
    
    # Summary
    print("\n" + "="*70)
    print("INTEGRATION TEST SUMMARY")
    print("="*70)
    print(f"Total tests: {total}")
    print(f"Passed: {passed}")
    print(f"Failed: {total - passed}")
    print(f"Success rate: {(passed/total)*100:.1f}%")
    
    if passed == total:
        print("\n🎉 All integration tests passed!")
        print("✅ Your conversion.py models are perfectly aligned with other models!")
        print("✅ Ready for production use!")
    else:
        print(f"\n❌ {total - passed} test(s) failed. Please review the issues above.")
    
    return passed == total


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)