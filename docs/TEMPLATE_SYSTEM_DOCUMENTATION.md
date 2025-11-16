# 🎨 Professional Template System - Complete Documentation

## Overview

The FinDeck template system provides **10 professional built-in templates** and **full custom template support** for creating beautiful, branded PowerPoint presentations.

---

## ✨ Features

### 1. **10 Built-in Professional Templates**

| Template | Category | Best For | Primary Color |
|----------|----------|----------|---------------|
| **Corporate Blue** | Business | Executive presentations, board meetings | Navy Blue (#1F4788) |
| **Financial Green** | Finance | Investment reports, financial analysis | Emerald Green (#0F5132) |
| **Executive Dark** | Executive | C-suite presentations, high-level strategy | Dark Gray + Gold (#2C3E50) |
| **Minimal White** | Minimal | Data-focused presentations, reports | Clean Gray (#495057) |
| **Vibrant Orange** | Creative | Marketing, startup pitches | Energetic Orange (#D84315) |
| **Professional Purple** | Business | Tech companies, innovation | Professional Purple (#6A1B9A) |
| **Tech Gradient** | Technology | Software, digital innovation | Blue-Purple Gradient (#3F51B5) |
| **Modern Teal** | Modern | Healthcare, consulting | Contemporary Teal (#00796B) |
| **Elegant Gold** | Luxury | High-end brands, premium products | Black + Gold (#C9A961) |
| **Classic Red** | Business | Sales pitches, important announcements | Bold Red (#C62828) |

### 2. **Custom Template Support**
- ✅ Upload custom templates (JSON format)
- ✅ Duplicate & modify existing templates
- ✅ Export templates to share
- ✅ Delete custom templates
- ✅ Validate template structure

### 3. **Template Components**
Each template includes:
- **Color Palette** - Primary, secondary, accent, background, text colors
- **Chart Colors** - 7 coordinated colors for data visualization
- **Font Styles** - Title, subtitle, heading, body, table fonts
- **Layout Positions** - Precise positioning for all elements
- **Styling Options** - Table colors, borders, shadows

---

## 📁 Project Structure

```
FinDeck/
├── templates/
│   ├── built_in/                    # 10 professional templates
│   │   ├── corporate_blue.json
│   │   ├── financial_green.json
│   │   ├── executive_dark.json
│   │   ├── minimal_white.json
│   │   ├── vibrant_orange.json
│   │   ├── professional_purple.json
│   │   ├── tech_gradient.json
│   │   ├── modern_teal.json
│   │   ├── elegant_gold.json
│   │   └── classic_red.json
│   └── custom/                      # User custom templates
│       └── my_custom_theme.json
│
├── src/
│   ├── templates/
│   │   ├── __init__.py
│   │   ├── template_manager.py      # Core template management
│   │   └── template_charts.py       # Template-aware chart creation
│   │
│   └── backend/app/api/v1/endpoints/
│       └── templates.py             # REST API endpoints
│
└── demo_templates.py                # Demo & testing script
```

---

## 🚀 Usage

### Python API

#### 1. **List Available Templates**
```python
from templates.template_manager import TemplateManager

manager = TemplateManager()
templates = manager.list_templates()

for template in templates:
    print(f"{template['name']} - {template['category']}")
```

#### 2. **Load a Template**
```python
template = manager.load_template("corporate_blue")
print(f"Primary Color: {template['colors']['primary']}")
print(f"Chart Colors: {template['colors']['chart_colors']}")
```

#### 3. **Create Presentation with Template**
```python
from templates.template_charts import create_templated_chart_slide
from pptx import Presentation

# Load template
template = manager.load_template("financial_green")

# Create presentation
prs = Presentation()

# Create chart with template styling
slide = create_templated_chart_slide(
    prs, df, 
    chart_type='pie', 
    config={'category_col': 'Asset', 'value_col': 'Value'},
    title="Portfolio Allocation",
    template=template
)

prs.save("output.pptx")
```

#### 4. **Create Custom Template**
```python
# Duplicate existing template
manager.duplicate_template("corporate_blue", "my_company_theme")

# Load and modify
custom = manager.load_template("my_company_theme")
custom['name'] = "My Company Theme"
custom['colors']['primary'] = "#FF0000"  # Company red
custom['colors']['chart_colors'] = ["#FF0000", "#CC0000", "#990000"]

# Save
manager.save_custom_template(custom, "my_company_theme", overwrite=True)
```

#### 5. **Export/Import Templates**
```python
# Export for sharing
manager.export_template("my_company_theme", "my_theme.json")

# Import on another system
manager.import_template("my_theme.json", name="imported_theme")
```

---

### REST API Endpoints

Base URL: `http://localhost:8000/api/v1`

#### **GET /templates**
List all available templates
```bash
curl http://localhost:8000/api/v1/templates
```

Response:
```json
[
  {
    "id": "corporate_blue",
    "name": "Corporate Blue",
    "description": "Classic professional business theme",
    "category": "business",
    "source": "built-in"
  },
  ...
]
```

#### **GET /templates/{template_name}**
Get detailed template configuration
```bash
curl http://localhost:8000/api/v1/templates/corporate_blue
```

#### **GET /templates/{template_name}/colors**
Get template color palette
```bash
curl http://localhost:8000/api/v1/templates/financial_green/colors
```

Response:
```json
{
  "template": "financial_green",
  "colors": {
    "primary": "#0F5132",
    "secondary": "#198754",
    "accent": "#20C997",
    "chart_colors": ["#0F5132", "#198754", "#20C997", ...]
  }
}
```

#### **POST /templates/custom**
Upload custom template
```bash
curl -X POST http://localhost:8000/api/v1/templates/custom \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -F "file=@my_template.json" \
  -F "name=my_custom_theme"
```

#### **DELETE /templates/custom/{template_name}**
Delete custom template (built-in cannot be deleted)
```bash
curl -X DELETE http://localhost:8000/api/v1/templates/custom/my_custom_theme \
  -H "Authorization: Bearer YOUR_TOKEN"
```

#### **POST /templates/{template_name}/duplicate**
Duplicate template for customization
```bash
curl -X POST http://localhost:8000/api/v1/templates/corporate_blue/duplicate \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -d "new_name=my_company_theme"
```

#### **GET /templates/{template_name}/export**
Export template as JSON file
```bash
curl -H "Authorization: Bearer YOUR_TOKEN" \
  http://localhost:8000/api/v1/templates/my_custom_theme/export \
  -o my_theme.json
```

---

## 🎨 Template JSON Format

```json
{
  "name": "Template Name",
  "description": "Template description",
  "category": "business",
  
  "colors": {
    "primary": "#1F4788",
    "secondary": "#4A90E2",
    "accent": "#7CB9E8",
    "background": "#FFFFFF",
    "text": "#333333",
    "text_light": "#666666",
    "chart_colors": ["#1F4788", "#4A90E2", "#7CB9E8", ...]
  },
  
  "fonts": {
    "title": {
      "name": "Calibri",
      "size": 44,
      "bold": true,
      "color": "#1F4788"
    },
    "subtitle": { ... },
    "heading": { ... },
    "body": { ... },
    "table_header": { ... },
    "table_data": { ... }
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
    "table_header_color": "#1F4788",
    "table_alt_row_color": "#F2F7FC",
    "border_color": "#1F4788",
    "border_width": 1.5,
    "shadow": true,
    "rounded_corners": false
  }
}
```

---

## 🧪 Testing & Demos

### Run Template Manager Demo
```bash
# Test template loading
python src/templates/template_manager.py
```

### Create Presentations with Templates
```bash
# Show all template colors
python demo_templates.py --mode colors

# Create presentation with single template
python demo_templates.py --mode single --template corporate_blue

# Create presentations with ALL 10 templates
python demo_templates.py --mode all

# Test custom template creation
python demo_templates.py --mode custom
```

---

## 📊 Demo Results

### Generated Files:
- ✅ **10 Built-in Templates** - All loaded and validated
- ✅ **Custom Template** - Created and tested (my_custom_theme)
- ✅ **Presentations** - Generated with template styling
- ✅ **Color Palettes** - All 10 templates have coordinated colors
- ✅ **API Endpoints** - 9 endpoints for template management

### Verified Features:
- ✅ Template loading from JSON
- ✅ Template validation
- ✅ Custom template creation
- ✅ Template duplication
- ✅ Color palette application
- ✅ Font styling
- ✅ Layout positioning
- ✅ Chart color customization
- ✅ Table styling
- ✅ Export/Import functionality

---

## 🎯 Use Cases

### 1. **Corporate Presentations**
Use `corporate_blue` or `professional_purple` for business meetings

### 2. **Financial Reports**
Use `financial_green` for investment analysis and banking

### 3. **Executive Briefings**
Use `executive_dark` for C-suite presentations

### 4. **Marketing Pitches**
Use `vibrant_orange` for creative, energetic presentations

### 5. **Technology Demos**
Use `tech_gradient` for software and digital products

### 6. **Healthcare/Consulting**
Use `modern_teal` for modern professional services

### 7. **Luxury Brands**
Use `elegant_gold` for premium, high-end presentations

### 8. **Data Analytics**
Use `minimal_white` for clean, data-focused reports

### 9. **Sales Presentations**
Use `classic_red` for bold, impactful pitches

### 10. **Custom Branding**
Create custom templates matching your company colors

---

## 🔧 Advanced Customization

### Programmatic Template Modification
```python
# Load template
template = manager.load_template("corporate_blue")

# Modify colors
template['colors']['primary'] = "#YOUR_BRAND_COLOR"
template['colors']['chart_colors'] = ["#COLOR1", "#COLOR2", "#COLOR3"]

# Modify fonts
template['fonts']['title']['size'] = 48
template['fonts']['title']['name'] = "Arial"

# Modify layout
template['layout']['content_slide']['chart_position']['width'] = 6.0

# Save as new template
manager.save_custom_template(template, "my_modified_theme")
```

---

## 📝 Template Validation

Templates are automatically validated for:
- ✅ Required sections (name, colors, fonts, layout, styling)
- ✅ Valid hex color codes
- ✅ Font configurations (name, size)
- ✅ Layout positions (left, top, width, height)
- ✅ Styling properties

Invalid templates will be rejected with detailed error messages.

---

## 🚀 Integration with Conversion API

Templates automatically integrate with the Excel-to-PPT conversion:

```python
# In conversions endpoint
@router.post("/convert")
async def convert_excel_to_ppt(
    file_id: str,
    template_name: str = "corporate_blue",  # NEW parameter
    ...
):
    # Load template
    template = template_manager.load_template(template_name)
    
    # Create presentation with template
    prs = create_presentation(
        title="Report",
        subtitle="Generated with FinDeck",
        template_name=template_name  # Apply template
    )
    
    # Charts automatically use template colors
    ...
```

---

## 🎉 Summary

The template system provides **enterprise-grade presentation theming** with:
- ✅ 10 professional built-in templates
- ✅ Full custom template support
- ✅ REST API integration
- ✅ Programmatic access
- ✅ Validation & error handling
- ✅ Export/Import capabilities
- ✅ Easy integration with conversion pipeline

**Ready for production use!** 🚀
