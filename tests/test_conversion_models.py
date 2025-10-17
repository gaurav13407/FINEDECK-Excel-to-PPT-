"""
Test cases for conversion.py models
Tests all model validation, relationships, and utility functions
"""

import pytest
import sys
import os
from datetime import datetime, timedelta
from bson import ObjectId

# Add the src directory to the path so we can import our modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src', 'backend', 'app'))

from models.conversion import (
    PowerPointTemplate,
    ConversionSettings, 
    ConversionResult,
    TemplateUsageStats,
    LayoutType,
    CharType,
    DataMappingType,
    get_template_by_category,
    calculate_conversion_credits,
    validate_conversion_settings,
    estimate_conversion_time
)
from models.user import TemplateCategory, PyObjectId


class TestPowerPointTemplate:
    """Test PowerPointTemplate model validation and functionality"""
    
    def test_valid_template_creation(self):
        """Test creating a valid PowerPoint template"""
        template_data = {
            "template_name": "Corporate Financial Report",
            "category": TemplateCategory.PROFESSIONAL,
            "description": "Professional template for financial presentations",
            "template_file_path": "/templates/corporate_financial.pptx",
            "thumbnail_path": "/thumbnails/corporate_financial.jpg",
            "slide_count": 15,
            "supported_layouts": [LayoutType.TITLE_SLIDE, LayoutType.TITLE_CONTENT],
            "color_scheme": ["#1f4e79", "#70ad47", "#ffffff"],
            "font_families": ["Calibri", "Arial"],
            "is_premium": False,
            "required_subscription": ["free", "pro", "enterprise"]
        }
        
        template = PowerPointTemplate(**template_data)
        assert template.template_name == "Corporate Financial Report"
        assert template.category == TemplateCategory.PROFESSIONAL
        assert template.slide_count == 15
        assert template.is_active == True  # Default value
        assert isinstance(template.created_at, datetime)

    def test_template_name_validation(self):
        """Test template name length validation"""
        # Test too short name
        with pytest.raises(ValueError):
            PowerPointTemplate(
                template_name="",
                category=TemplateCategory.BASIC,
                template_file_path="/test.pptx",
                slide_count=5
            )
        
        # Test too long name
        with pytest.raises(ValueError):
            PowerPointTemplate(
                template_name="x" * 101,
                category=TemplateCategory.BASIC,
                template_file_path="/test.pptx",
                slide_count=5
            )

    def test_slide_count_validation(self):
        """Test slide count validation"""
        # Test minimum slides
        with pytest.raises(ValueError):
            PowerPointTemplate(
                template_name="Test Template",
                category=TemplateCategory.BASIC,
                template_file_path="/test.pptx",
                slide_count=0
            )
        
        # Test maximum slides
        with pytest.raises(ValueError):
            PowerPointTemplate(
                template_name="Test Template",
                category=TemplateCategory.BASIC,
                template_file_path="/test.pptx",
                slide_count=51
            )


class TestConversionSettings:
    """Test ConversionSettings model validation"""
    
    def test_default_settings(self):
        """Test default conversion settings"""
        settings = ConversionSettings()
        assert settings.preferred_layouts == LayoutType.TITLE_CONTENT
        assert settings.slide_orientation == "landscape"
        assert settings.data_mapping_strategy == DataMappingType.ROW_TO_SLIDE
        assert settings.max_rows_per_slide == 10
        assert settings.convert_charts == True
        assert settings.font_size_title == 24
        assert settings.font_size_body == 18
        assert settings.primary_color == "#1f4e79"
        assert settings.quality_level == "high"

    def test_color_validation(self):
        """Test color format validation"""
        # Valid colors
        valid_settings = ConversionSettings(
            primary_color="#123456",
            secondary_color="#abcdef",
            background_color="#FFFFFF"
        )
        assert valid_settings.primary_color == "#123456"
        
        # Invalid color format
        with pytest.raises(ValueError):
            ConversionSettings(primary_color="invalid_color")
        
        with pytest.raises(ValueError):
            ConversionSettings(primary_color="#12345")  # Too short
        
        with pytest.raises(ValueError):
            ConversionSettings(primary_color="#1234567")  # Too long

    def test_font_size_validation(self):
        """Test font size validation"""
        # Valid font sizes
        settings = ConversionSettings(
            font_size_title=30,
            font_size_body=16
        )
        assert settings.font_size_title == 30
        assert settings.font_size_body == 16
        
        # Invalid font sizes
        with pytest.raises(ValueError):
            ConversionSettings(font_size_title=10)  # Too small
        
        with pytest.raises(ValueError):
            ConversionSettings(font_size_body=40)  # Too large

    def test_slide_orientation_validation(self):
        """Test slide orientation validation"""
        # Valid orientations
        landscape_settings = ConversionSettings(slide_orientation="landscape")
        portrait_settings = ConversionSettings(slide_orientation="portrait")
        
        assert landscape_settings.slide_orientation == "landscape"
        assert portrait_settings.slide_orientation == "portrait"
        
        # Invalid orientation
        with pytest.raises(ValueError):
            ConversionSettings(slide_orientation="invalid")


class TestConversionResult:
    """Test ConversionResult model"""
    
    def test_valid_conversion_result(self):
        """Test creating a valid conversion result"""
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
            output_file_size=2048576,  # 2MB
            slide_count=12,
            processing_time_seconds=300.5,
            data_rows_processed=150,
            charts_converted=3,
            images_processed=5,
            started_at=started_time,
            completed_at=completed_time
        )
        
        assert result.conversion_id == "conv_123"
        assert result.slide_count == 12
        assert result.processing_time_seconds == 300.5
        assert result.success_rate == 100.0  # Default value
        assert result.error_count == 0  # Default value

    def test_success_rate_validation(self):
        """Test success rate validation"""
        # Valid success rates
        for rate in [0.0, 50.5, 100.0]:
            result = ConversionResult(
                conversion_id="test",
                job_id="test",
                user_id=PyObjectId(),
                source_file_id="test",
                output_file_path="/test.pptx",
                output_file_size=1024,
                slide_count=1,
                processing_time_seconds=10.0,
                started_at=datetime.utcnow(),
                completed_at=datetime.utcnow(),
                success_rate=rate
            )
            assert result.success_rate == rate
        
        # Invalid success rates
        with pytest.raises(ValueError):
            ConversionResult(
                conversion_id="test",
                job_id="test",
                user_id=PyObjectId(),
                source_file_id="test",
                output_file_path="/test.pptx",
                output_file_size=1024,
                slide_count=1,
                processing_time_seconds=10.0,
                started_at=datetime.utcnow(),
                completed_at=datetime.utcnow(),
                success_rate=-1.0  # Negative rate
            )

    def test_user_rating_validation(self):
        """Test user rating validation"""
        result_data = {
            "conversion_id": "test",
            "job_id": "test",
            "user_id": PyObjectId(),
            "source_file_id": "test",
            "output_file_path": "/test.pptx",
            "output_file_size": 1024,
            "slide_count": 1,
            "processing_time_seconds": 10.0,
            "started_at": datetime.utcnow(),
            "completed_at": datetime.utcnow()
        }
        
        # Valid ratings
        for rating in [1, 2, 3, 4, 5]:
            result = ConversionResult(**result_data, user_rating=rating)
            assert result.user_rating == rating
        
        # Invalid ratings
        with pytest.raises(ValueError):
            ConversionResult(**result_data, user_rating=0)  # Too low
        
        with pytest.raises(ValueError):
            ConversionResult(**result_data, user_rating=6)  # Too high


class TestTemplateUsageStats:
    """Test TemplateUsageStats model"""
    
    def test_valid_usage_stats(self):
        """Test creating valid usage statistics"""
        stats = TemplateUsageStats(
            template_id="template_123",
            template_name="Corporate Template",
            category=TemplateCategory.PROFESSIONAL,
            total_uses=150,
            unique_users=75,
            success_rate=95.5,
            average_rating=4.2,
            total_ratings=45,
            usage_trend=12.5,
            rating_trend=-0.3
        )
        
        assert stats.template_id == "template_123"
        assert stats.total_uses == 150
        assert stats.unique_users == 75
        assert stats.success_rate == 95.5
        assert stats.average_rating == 4.2

    def test_default_values(self):
        """Test default values for usage stats"""
        stats = TemplateUsageStats(
            template_id="template_123",
            template_name="Test Template",
            category=TemplateCategory.BASIC
        )
        
        assert stats.total_uses == 0
        assert stats.unique_users == 0
        assert stats.success_rate == 0
        assert stats.average_rating is None
        assert stats.total_ratings == 0
        assert stats.usage_trend == 0.0
        assert stats.rating_trend == 0.0


class TestUtilityFunctions:
    """Test utility functions"""
    
    def test_calculate_conversion_credits(self):
        """Test credit calculation logic"""
        # Basic conversion (small file, few slides, basic template)
        credits = calculate_conversion_credits(
            file_size_mb=5.0,
            slide_count=5,
            template_category=TemplateCategory.BASIC,
            include_charts=False
        )
        assert credits == 1  # Base credits only
        
        # Large file conversion
        credits = calculate_conversion_credits(
            file_size_mb=15.0,
            slide_count=8,
            template_category=TemplateCategory.BASIC,
            include_charts=False
        )
        assert credits == 2  # Base + file size
        
        # Premium template with charts
        credits = calculate_conversion_credits(
            file_size_mb=8.0,
            slide_count=15,
            template_category=TemplateCategory.PREMIUM,
            include_charts=True
        )
        assert credits == 5  # Base + slides + premium + charts
        
        # Maximum credits test
        credits = calculate_conversion_credits(
            file_size_mb=100.0,
            slide_count=50,
            template_category=TemplateCategory.CUSTOM,
            include_charts=True
        )
        assert credits <= 10  # Should be capped at 10

    def test_validate_conversion_settings(self):
        """Test conversion settings validation"""
        valid_settings = ConversionSettings()
        result = validate_conversion_settings(valid_settings)
        
        assert result["is_valid"] == True
        assert isinstance(result["errors"], list)
        assert isinstance(result["warnings"], list)
        assert "normalized_settings" in result

    def test_estimate_conversion_time(self):
        """Test conversion time estimation"""
        # Small conversion
        time_estimate = estimate_conversion_time(
            file_size_mb=2.0,
            row_count=50,
            chart_count=1,
            template_complexity="simple"
        )
        assert time_estimate > 0
        assert isinstance(time_estimate, float)
        
        # Large conversion
        large_estimate = estimate_conversion_time(
            file_size_mb=20.0,
            row_count=500,
            chart_count=10,
            template_complexity="premium"
        )
        assert large_estimate > time_estimate  # Should take longer
        
        # Test different complexity levels
        simple_time = estimate_conversion_time(5.0, 100, 2, "simple")
        complex_time = estimate_conversion_time(5.0, 100, 2, "complex")
        assert complex_time > simple_time


class TestEnumValidation:
    """Test enum value validation"""
    
    def test_layout_type_enum(self):
        """Test LayoutType enum values"""
        assert LayoutType.TITLE_SLIDE == "title_slide"
        assert LayoutType.TITLE_CONTENT == "title_content"
        assert LayoutType.COMPARISON == "comparison"
        
        # Test all values are accessible
        all_layouts = [item.value for item in LayoutType]
        expected_layouts = [
            "title_slide", "title_content", "section_header", 
            "two_content", "comparison", "title_only", "blank",
            "content_with_caption", "picture_with_caption"
        ]
        assert set(all_layouts) == set(expected_layouts)

    def test_chart_type_enum(self):
        """Test CharType enum values"""
        assert CharType.BAR == "bar"
        assert CharType.LINE == "line"
        assert CharType.PIE == "pie"
        
        # Test all chart types
        all_charts = [item.value for item in CharType]
        expected_charts = ["bar", "line", "pie", "column", "area", "scatter", "donut", "combo"]
        assert set(all_charts) == set(expected_charts)

    def test_data_mapping_type_enum(self):
        """Test DataMappingType enum values"""
        assert DataMappingType.ROW_TO_SLIDE == "row_to_slide"
        assert DataMappingType.SHEET_TO_SLIDE == "sheet_to_slide"
        assert DataMappingType.CHART_TO_SLIDE == "chart_to_slide"


class TestModelIntegration:
    """Test integration between models and other modules"""
    
    def test_template_category_import(self):
        """Test that TemplateCategory is properly imported from user module"""
        template = PowerPointTemplate(
            template_name="Test Template",
            category=TemplateCategory.PROFESSIONAL,
            template_file_path="/test.pptx",
            slide_count=5
        )
        assert template.category == TemplateCategory.PROFESSIONAL
        
        stats = TemplateUsageStats(
            template_id="test",
            template_name="Test",
            category=TemplateCategory.PREMIUM
        )
        assert stats.category == TemplateCategory.PREMIUM

    def test_pyobjectid_integration(self):
        """Test PyObjectId integration"""
        user_id = PyObjectId()
        assert isinstance(user_id, ObjectId)
        
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
        assert result.user_id == user_id


# Test runner function
def run_tests():
    """Run all tests and print results"""
    print("Starting conversion.py model tests...")
    
    test_classes = [
        TestPowerPointTemplate,
        TestConversionSettings,
        TestConversionResult,
        TestTemplateUsageStats,
        TestUtilityFunctions,
        TestEnumValidation,
        TestModelIntegration
    ]
    
    total_tests = 0
    passed_tests = 0
    failed_tests = []
    
    for test_class in test_classes:
        print(f"\n--- Running {test_class.__name__} ---")
        test_instance = test_class()
        
        # Get all test methods
        test_methods = [method for method in dir(test_instance) if method.startswith('test_')]
        
        for test_method in test_methods:
            total_tests += 1
            try:
                getattr(test_instance, test_method)()
                print(f"✅ {test_method}")
                passed_tests += 1
            except Exception as e:
                print(f"❌ {test_method}: {str(e)}")
                failed_tests.append(f"{test_class.__name__}.{test_method}: {str(e)}")
    
    print(f"\n{'='*60}")
    print(f"TEST SUMMARY")
    print(f"{'='*60}")
    print(f"Total tests: {total_tests}")
    print(f"Passed: {passed_tests}")
    print(f"Failed: {len(failed_tests)}")
    print(f"Success rate: {(passed_tests/total_tests)*100:.1f}%")
    
    if failed_tests:
        print(f"\nFAILED TESTS:")
        for failure in failed_tests:
            print(f"  - {failure}")
    else:
        print(f"\n🎉 All tests passed!")
    
    return len(failed_tests) == 0


if __name__ == "__main__":
    success = run_tests()
    exit(0 if success else 1)