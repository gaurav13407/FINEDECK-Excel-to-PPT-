"""
Quick Start Guide - Using Templates in Your Application
"""

import sys
import os
sys.path.insert(0, 'src')

# ==============================================================================
# 1. BASIC USAGE - Create Presentation with Template
# ==============================================================================

from src.templates.template_manager import TemplateManager
from src.templates.template_charts import create_templated_chart_slide
from pptx import Presentation
from src.converter.excel_reader import excel_reader
from src.converter.chart_detector import detect_chart_type

def quick_convert(excel_file, output_file, template_name="corporate_blue"):
    """Simple Excel to PPT with template"""
    
    # Load template
    manager = TemplateManager()
    template = manager.load_template(template_name)
    
    # Create presentation
    prs = Presentation()
    
    # Read Excel data
    df = excel_reader(excel_file)
    
    # Detect chart type
    chart_type, config = detect_chart_type(df)
    
    # Create slide with template styling
    create_templated_chart_slide(
        prs, df, chart_type, config,
        title="Financial Report",
        template=template
    )
    
    # Save
    prs.save(output_file)
    print(f"✅ Created: {output_file}")


# ==============================================================================
# 2. LIST ALL AVAILABLE TEMPLATES
# ==============================================================================

def list_all_templates():
    """Show all available templates"""
    manager = TemplateManager()
    templates = manager.list_templates()
    
    print("\n📋 Available Templates:\n")
    for t in templates:
        print(f"  • {t['id']:25} - {t['name']:30} [{t['category']}]")


# ==============================================================================
# 3. CREATE CUSTOM TEMPLATE FROM SCRATCH
# ==============================================================================

def create_my_company_template():
    """Create a custom branded template"""
    
    custom_template = {
        "name": "My Company Brand",
        "description": "Custom template with company colors",
        "category": "custom",
        
        "colors": {
            "primary": "#FF0000",        # Your brand primary
            "secondary": "#CC0000",      # Your brand secondary
            "accent": "#FF6666",         # Your brand accent
            "background": "#FFFFFF",
            "text": "#333333",
            "text_light": "#666666",
            "chart_colors": [
                "#FF0000", "#CC0000", "#FF6666", 
                "#FF9999", "#990000", "#FFCCCC"
            ]
        },
        
        "fonts": {
            "title": {
                "name": "Arial",
                "size": 44,
                "bold": True,
                "color": "#FF0000"
            },
            "subtitle": {
                "name": "Arial",
                "size": 28,
                "bold": False,
                "color": "#666666"
            },
            "heading": {
                "name": "Arial",
                "size": 32,
                "bold": True,
                "color": "#CC0000"
            },
            "body": {
                "name": "Arial",
                "size": 18,
                "bold": False,
                "color": "#333333"
            },
            "table_header": {
                "name": "Arial",
                "size": 11,
                "bold": True,
                "color": "#FFFFFF"
            },
            "table_data": {
                "name": "Arial",
                "size": 10,
                "bold": False,
                "color": "#333333"
            }
        },
        
        "layout": {
            "title_slide": {
                "title_position": {"left": 0.5, "top": 2.5, "width": 9, "height": 1.5},
                "subtitle_position": {"left": 0.5, "top": 4.0, "width": 9, "height": 1.0}
            },
            "content_slide": {
                "title_position": {"left": 0.5, "top": 0.3, "width": 9, "height": 0.6},
                "chart_position": {"left": 0.5, "top": 1.2, "width": 5.5, "height": 5.0},
                "table_position": {"left": 6.2, "top": 1.2, "width": 3.5, "height": 5.0}
            }
        },
        
        "styling": {
            "table_header_color": "#FF0000",
            "table_alt_row_color": "#FFF0F0",
            "border_color": "#CC0000",
            "border_width": 2.0,
            "shadow": True,
            "rounded_corners": False
        }
    }
    
    # Save template
    manager = TemplateManager()
    manager.save_custom_template(custom_template, "my_company_brand")
    print("✅ Custom template created: my_company_brand")


# ==============================================================================
# 4. MODIFY EXISTING TEMPLATE
# ==============================================================================

def customize_existing_template():
    """Duplicate and modify an existing template"""
    
    manager = TemplateManager()
    
    # Duplicate corporate_blue
    manager.duplicate_template("corporate_blue", "my_modified_blue")
    
    # Load the duplicate
    template = manager.load_template("my_modified_blue")
    
    # Modify colors
    template['colors']['primary'] = "#003366"  # Darker blue
    template['colors']['chart_colors'] = [
        "#003366", "#0066CC", "#3399FF", 
        "#66B2FF", "#99CCFF"
    ]
    
    # Save changes
    manager.save_custom_template(template, "my_modified_blue", overwrite=True)
    print("✅ Modified template saved: my_modified_blue")


# ==============================================================================
# 5. USE TEMPLATE WITH CONVERSION API
# ==============================================================================

async def convert_with_template_api(file_id: str, template_name: str):
    """
    Example of using templates in your FastAPI conversion endpoint
    """
    from fastapi import HTTPException
    
    # Load template
    manager = TemplateManager()
    template = manager.load_template(template_name)
    
    if not template:
        raise HTTPException(
            status_code=404,
            detail=f"Template '{template_name}' not found"
        )
    
    # Get template colors for charts
    chart_colors = template['colors']['chart_colors']
    
    # Your existing conversion logic here...
    # But now apply template styling
    
    return {
        "message": "Conversion complete",
        "template_used": template_name,
        "colors_applied": chart_colors
    }


# ==============================================================================
# 6. FRONTEND INTEGRATION EXAMPLES
# ==============================================================================

"""
// React/Vue Frontend Example

async function loadTemplates() {
    const response = await fetch('/api/v1/templates');
    const templates = await response.json();
    
    // Display in dropdown
    return templates.map(t => ({
        value: t.id,
        label: t.name,
        category: t.category
    }));
}

async function convertWithTemplate(fileId, templateId) {
    const formData = new FormData();
    formData.append('file_id', fileId);
    formData.append('template_name', templateId);
    
    const response = await fetch('/api/v1/conversions/convert', {
        method: 'POST',
        body: formData,
        headers: {
            'Authorization': `Bearer ${token}`
        }
    });
    
    return response.blob(); // Download PPT
}

async function uploadCustomTemplate(file) {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('name', 'my_custom');
    
    const response = await fetch('/api/v1/templates/custom', {
        method: 'POST',
        body: formData,
        headers: {
            'Authorization': `Bearer ${token}`
        }
    });
    
    return response.json();
}
"""


# ==============================================================================
# 7. TEMPLATE VALIDATION
# ==============================================================================

def validate_custom_template(template_dict):
    """Validate a custom template before saving"""
    
    manager = TemplateManager()
    
    if manager.validate_template(template_dict):
        print("✅ Template is valid!")
        return True
    else:
        print("❌ Template validation failed")
        return False


# ==============================================================================
# 8. GET TEMPLATE COLORS FOR PREVIEW
# ==============================================================================

def get_template_preview_colors(template_name):
    """Get color palette for UI preview"""
    
    manager = TemplateManager()
    template = manager.load_template(template_name)
    
    if template:
        return {
            'primary': template['colors']['primary'],
            'secondary': template['colors']['secondary'],
            'accent': template['colors']['accent'],
            'chart_colors': template['colors']['chart_colors']
        }


# ==============================================================================
# USAGE EXAMPLES
# ==============================================================================

if __name__ == "__main__":
    
    # Example 1: List templates
    print("\n" + "="*70)
    print("Example 1: List All Templates")
    print("="*70)
    list_all_templates()
    
    # Example 2: Quick conversion
    print("\n" + "="*70)
    print("Example 2: Quick Conversion with Template")
    print("="*70)
    try:
        quick_convert(
            "examples/Portfolio Allocation Data.xlsx",
            "output_with_template.pptx",
            "financial_green"
        )
    except Exception as e:
        print(f"Note: {e}")
    
    # Example 3: Create custom template
    print("\n" + "="*70)
    print("Example 3: Create Custom Template")
    print("="*70)
    create_my_company_template()
    
    # Example 4: Get colors for preview
    print("\n" + "="*70)
    print("Example 4: Get Template Colors")
    print("="*70)
    colors = get_template_preview_colors("corporate_blue")
    if colors:
        print(f"Primary: {colors['primary']}")
        print(f"Chart Colors: {', '.join(colors['chart_colors'][:3])}...")
    
    print("\n" + "="*70)
    print("✅ All examples complete!")
    print("="*70 + "\n")
