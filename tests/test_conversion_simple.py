"""
Simple test script for conversion.py models
Run this to verify the models work correctly without external dependencies
"""

import sys
import os
from datetime import datetime, timedelta
from bson import ObjectId

# Add the src directory to the path so we can import our modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src', 'backend', 'app'))

try:
    from models.conversion import (
        PowerPointTemplate,
        ConversionSettings, 
        ConversionResult,
        TemplateUsageStats,
        LayoutType,
        CharType,
        DataMappingType,
        calculate_conversion_credits,
        validate_conversion_settings,
        estimate_conversion_time
    )
    from models.user import TemplateCategory, PyObjectId
    print("✅ All imports successful!")
except ImportError as e:
    print(f"❌ Import error: {e}")
    sys.exit(1)


def test_basic_model_creation():
    """Test basic model creation and validation"""
    print("\n--- Testing Basic Model Creation ---")
    
    # Test PowerPointTemplate
    try:
        template = PowerPointTemplate(
            template_name="Corporate Financial Report",
            category=TemplateCategory.PROFESSIONAL,
            description="Professional template for financial presentations",
            template_file_path="/templates/corporate_financial.pptx",
            slide_count=15,
            supported_layouts=[LayoutType.TITLE_SLIDE, LayoutType.TITLE_CONTENT],
            color_scheme=["#1f4e79", "#70ad47", "#ffffff"],
            font_families=["Calibri", "Arial"]
        )
        print("✅ PowerPointTemplate creation successful")
        print(f"   Template name: {template.template_name}")
        print(f"   Category: {template.category}")
        print(f"   Slide count: {template.slide_count}")
        print(f"   Is active: {template.is_active}")
    except Exception as e:
        print(f"❌ PowerPointTemplate creation failed: {e}")
        return False
    
    # Test ConversionSettings
    try:
        settings = ConversionSettings(
            preferred_layouts=LayoutType.TITLE_CONTENT,
            data_mapping_strategy=DataMappingType.ROW_TO_SLIDE,
            max_rows_per_slide=8,
            convert_charts=True,
            font_size_title=28,
            font_size_body=20,
            primary_color="#1f4e79",
            secondary_color="#70ad47"
        )
        print("✅ ConversionSettings creation successful")
        print(f"   Layout: {settings.preferred_layouts}")
        print(f"   Mapping strategy: {settings.data_mapping_strategy}")
        print(f"   Max rows per slide: {settings.max_rows_per_slide}")
        print(f"   Primary color: {settings.primary_color}")
    except Exception as e:
        print(f"❌ ConversionSettings creation failed: {e}")
        return False
    
    # Test ConversionResult
    try:
        user_id = PyObjectId()
        started_time = datetime.utcnow()
        completed_time = started_time + timedelta(minutes=5)
        
        result = ConversionResult(
            conversion_id="conv_123",
            job_id="job_456",
            user_id=user_id,
            source_file_id="file_789",
            template_id="template_101",
            output_file_path="/outputs/result.pptx",
            output_file_size=2048576,
            slide_count=12,
            processing_time_seconds=300.5,
            data_rows_processed=150,
            charts_converted=3,
            started_at=started_time,
            completed_at=completed_time
        )
        print("✅ ConversionResult creation successful")
        print(f"   Conversion ID: {result.conversion_id}")
        print(f"   Slide count: {result.slide_count}")
        print(f"   Processing time: {result.processing_time_seconds}s")
        print(f"   Success rate: {result.success_rate}%")
    except Exception as e:
        print(f"❌ ConversionResult creation failed: {e}")
        return False
    
    # Test TemplateUsageStats
    try:
        stats = TemplateUsageStats(
            template_id="template_123",
            template_name="Corporate Template",
            category=TemplateCategory.PROFESSIONAL,
            total_uses=150,
            unique_users=75,
            success_rate=95.5
        )
        print("✅ TemplateUsageStats creation successful")
        print(f"   Template ID: {stats.template_id}")
        print(f"   Total uses: {stats.total_uses}")
        print(f"   Success rate: {stats.success_rate}%")
    except Exception as e:
        print(f"❌ TemplateUsageStats creation failed: {e}")
        return False
    
    return True


def test_validation_rules():
    """Test model validation rules"""
    print("\n--- Testing Validation Rules ---")
    
    # Test invalid template name (too short)
    try:
        PowerPointTemplate(
            template_name="",
            category=TemplateCategory.BASIC,
            template_file_path="/test.pptx",
            slide_count=5
        )
        print("❌ Template name validation failed - empty name should be rejected")
        return False
    except ValueError:
        print("✅ Template name validation works - empty name rejected")
    except Exception as e:
        print(f"❌ Unexpected error in template name validation: {e}")
        return False
    
    # Test invalid slide count
    try:
        PowerPointTemplate(
            template_name="Test Template",
            category=TemplateCategory.BASIC,
            template_file_path="/test.pptx",
            slide_count=0
        )
        print("❌ Slide count validation failed - zero slides should be rejected")
        return False
    except ValueError:
        print("✅ Slide count validation works - zero slides rejected")
    except Exception as e:
        print(f"❌ Unexpected error in slide count validation: {e}")
        return False
    
    # Test invalid color format
    try:
        ConversionSettings(primary_color="invalid_color")
        print("❌ Color validation failed - invalid color should be rejected")
        return False
    except ValueError:
        print("✅ Color validation works - invalid color rejected")
    except Exception as e:
        print(f"❌ Unexpected error in color validation: {e}")
        return False
    
    # Test invalid font size
    try:
        ConversionSettings(font_size_title=10)  # Too small
        print("❌ Font size validation failed - small font should be rejected")
        return False
    except ValueError:
        print("✅ Font size validation works - small font rejected")
    except Exception as e:
        print(f"❌ Unexpected error in font size validation: {e}")
        return False
    
    return True


def test_utility_functions():
    """Test utility functions"""
    print("\n--- Testing Utility Functions ---")
    
    # Test credit calculation
    try:
        # Basic conversion
        basic_credits = calculate_conversion_credits(
            file_size_mb=5.0,
            slide_count=5,
            template_category=TemplateCategory.BASIC,
            include_charts=False
        )
        print(f"✅ Basic conversion credits: {basic_credits}")
        
        # Premium conversion with charts
        premium_credits = calculate_conversion_credits(
            file_size_mb=15.0,
            slide_count=15,
            template_category=TemplateCategory.PREMIUM,
            include_charts=True
        )
        print(f"✅ Premium conversion credits: {premium_credits}")
        
        # Verify premium costs more than basic
        if premium_credits > basic_credits:
            print("✅ Credit calculation logic works - premium costs more")
        else:
            print("❌ Credit calculation logic issue - premium should cost more")
            return False
            
    except Exception as e:
        print(f"❌ Credit calculation failed: {e}")
        return False
    
    # Test settings validation
    try:
        settings = ConversionSettings()
        result = validate_conversion_settings(settings)
        
        if isinstance(result, dict) and "is_valid" in result:
            print("✅ Settings validation function works")
            print(f"   Validation result: {result['is_valid']}")
        else:
            print("❌ Settings validation function returned invalid format")
            return False
    except Exception as e:
        print(f"❌ Settings validation failed: {e}")
        return False
    
    # Test time estimation
    try:
        time_estimate = estimate_conversion_time(
            file_size_mb=10.0,
            row_count=100,
            chart_count=5,
            template_complexity="medium"
        )
        
        if isinstance(time_estimate, (int, float)) and time_estimate > 0:
            print(f"✅ Time estimation works: {time_estimate:.1f} seconds")
        else:
            print("❌ Time estimation returned invalid value")
            return False
    except Exception as e:
        print(f"❌ Time estimation failed: {e}")
        return False
    
    return True


def test_enum_values():
    """Test enum values and accessibility"""
    print("\n--- Testing Enum Values ---")
    
    # Test LayoutType enum
    try:
        layout_values = [item.value for item in LayoutType]
        expected_layouts = [
            "title_slide", "title_content", "section_header", 
            "two_content", "comparison", "title_only", "blank",
            "content_with_caption", "picture_with_caption"
        ]
        
        if set(layout_values) == set(expected_layouts):
            print(f"✅ LayoutType enum has all expected values ({len(layout_values)} items)")
        else:
            print("❌ LayoutType enum missing or has extra values")
            print(f"   Expected: {expected_layouts}")
            print(f"   Found: {layout_values}")
            return False
    except Exception as e:
        print(f"❌ LayoutType enum test failed: {e}")
        return False
    
    # Test CharType enum
    try:
        chart_values = [item.value for item in CharType]
        expected_charts = ["bar", "line", "pie", "column", "area", "scatter", "donut", "combo"]
        
        if set(chart_values) == set(expected_charts):
            print(f"✅ CharType enum has all expected values ({len(chart_values)} items)")
        else:
            print("❌ CharType enum missing or has extra values")
            return False
    except Exception as e:
        print(f"❌ CharType enum test failed: {e}")
        return False
    
    # Test DataMappingType enum
    try:
        mapping_values = [item.value for item in DataMappingType]
        expected_mappings = ["row_to_slide", "sheet_to_slide", "chart_to_slide", "table_to_slide", "summary_to_slide"]
        
        if set(mapping_values) == set(expected_mappings):
            print(f"✅ DataMappingType enum has all expected values ({len(mapping_values)} items)")
        else:
            print("❌ DataMappingType enum missing or has extra values")
            return False
    except Exception as e:
        print(f"❌ DataMappingType enum test failed: {e}")
        return False
    
    return True


def test_integration_with_user_models():
    """Test integration with user models"""
    print("\n--- Testing Integration with User Models ---")
    
    # Test TemplateCategory integration
    try:
        for category in TemplateCategory:
            template = PowerPointTemplate(
                template_name=f"Test {category.value} Template",
                category=category,
                template_file_path=f"/templates/{category.value}.pptx",
                slide_count=10
            )
            print(f"✅ Template created with {category.value} category")
        
        print("✅ All TemplateCategory values work with PowerPointTemplate")
    except Exception as e:
        print(f"❌ TemplateCategory integration failed: {e}")
        return False
    
    # Test PyObjectId integration
    try:
        user_id = PyObjectId()
        
        result = ConversionResult(
            conversion_id="test",
            job_id="test",
            user_id=user_id,
            source_file_id="test",
            output_file_path="/test.pptx",
            output_file_size=1024,
            slide_count=1,
            processing_time_seconds=10.0,
            started_at=datetime.utcnow(),
            completed_at=datetime.utcnow()
        )
        
        if result.user_id == user_id:
            print("✅ PyObjectId integration works correctly")
        else:
            print("❌ PyObjectId integration failed - ID mismatch")
            return False
    except Exception as e:
        print(f"❌ PyObjectId integration failed: {e}")
        return False
    
    return True


def main():
    """Run all tests and provide summary"""
    print("="*60)
    print("CONVERSION MODELS TEST SUITE")
    print("="*60)
    
    tests = [
        ("Basic Model Creation", test_basic_model_creation),
        ("Validation Rules", test_validation_rules), 
        ("Utility Functions", test_utility_functions),
        ("Enum Values", test_enum_values),
        ("Integration with User Models", test_integration_with_user_models)
    ]
    
    passed = 0
    total = len(tests)
    failed_tests = []
    
    for test_name, test_func in tests:
        print(f"\n{'='*20} {test_name} {'='*20}")
        try:
            if test_func():
                passed += 1
                print(f"✅ {test_name} PASSED")
            else:
                failed_tests.append(test_name)
                print(f"❌ {test_name} FAILED")
        except Exception as e:
            failed_tests.append(test_name)
            print(f"❌ {test_name} FAILED with exception: {e}")
    
    # Summary
    print("\n" + "="*60)
    print("TEST SUMMARY")
    print("="*60)
    print(f"Total tests: {total}")
    print(f"Passed: {passed}")
    print(f"Failed: {len(failed_tests)}")
    print(f"Success rate: {(passed/total)*100:.1f}%")
    
    if failed_tests:
        print(f"\nFailed tests:")
        for test in failed_tests:
            print(f"  - {test}")
    else:
        print("\n🎉 All tests passed! Your conversion models are working correctly.")
    
    return len(failed_tests) == 0


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)