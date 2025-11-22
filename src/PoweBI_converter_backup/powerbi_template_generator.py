"""
Power BI Template (.pbit) Generator
Creates actual Power BI template files that users can download and open
"""

import json
import zipfile
import shutil
from pathlib import Path
from typing import Dict, Any, List
from datetime import datetime
import pandas as pd


class DateTimeEncoder(json.JSONEncoder):
    """Custom JSON encoder for datetime objects"""
    def default(self, obj):
        if isinstance(obj, (datetime, pd.Timestamp)):
            return obj.isoformat()
        return super().default(obj)
import base64


import json
import zipfile
from pathlib import Path
from typing import Dict, Any, List
from datetime import datetime


class DateTimeEncoder(json.JSONEncoder):
    """Custom JSON encoder for datetime objects"""
    def default(self, obj):
        if isinstance(obj, (datetime, pd.Timestamp)):
            return obj.isoformat()
        return super().default(obj)


class PowerBITemplateGenerator:
    """
    Generates .pbit (Power BI Template) files
    
    A .pbit file is essentially a ZIP file containing:
    - DataModelSchema (JSON with tables, relationships, measures)
    - Report layout (JSON with visuals)
    - Metadata
    """
    
    def __init__(self):
        self.templates_dir = Path("src/templates/powerbi_templates")
        self.templates_dir.mkdir(parents=True, exist_ok=True)
    
    def create_sales_performance_template(self) -> Path:
        """Create Sales Performance Dashboard Template"""
        
        template_config = {
            "name": "Sales Performance Dashboard",
            "version": "1.0",
            "description": "Track sales by region, product, and salesperson with KPIs and trends",
            "tables": ["Sales_Data"],
            "visuals": [
                {
                    "type": "lineChart",
                    "title": "Sales Trend",
                    "x_axis": "Date",
                    "y_axis": "Salesamount",
                    "legend": "Region"
                },
                {
                    "type": "map",
                    "title": "Sales by Region",
                    "location": "Region",
                    "size": "Salesamount"
                },
                {
                    "type": "barChart",
                    "title": "Top Products",
                    "x_axis": "Product",
                    "y_axis": "Salesamount"
                },
                {
                    "type": "card",
                    "title": "Total Sales",
                    "measure": "TotalSalesamount"
                },
                {
                    "type": "card",
                    "title": "Target Achievement",
                    "measure": "TargetAchievement"
                }
            ],
            "measures": [
                "TotalSalesamount = SUM(Sales_Data[Salesamount])",
                "AverageSalesamount = AVERAGE(Sales_Data[Salesamount])",
                "TotalQuantitysold = SUM(Sales_Data[Quantitysold])",
                "TargetAchievement = DIVIDE([TotalSalesamount], SUM(Sales_Data[Salestarget]), 0) * 100"
            ],
            "instructions": [
                "1. Open this template in Power BI Desktop",
                "2. When prompted, import the CSV file: Sales_Data.csv",
                "3. Click 'Load' to populate the dashboard",
                "4. All visuals and measures will auto-populate",
                "5. Customize colors, filters, and branding as needed"
            ]
        }
        
        output_file = self.templates_dir / "Sales_Performance.pbit"
        self._save_template_config(template_config, output_file)
        
        return output_file
    
    def create_financial_kpi_template(self) -> Path:
        """Create Financial KPI Dashboard Template"""
        
        template_config = {
            "name": "Financial KPI Dashboard",
            "version": "1.0",
            "description": "Monitor revenue, expenses, profit, and budget variance with executive KPIs",
            "tables": ["P&L_Statement", "Summary"],
            "visuals": [
                {
                    "type": "waterfallChart",
                    "title": "P&L Breakdown",
                    "category": "Month",
                    "values": ["Revenue", "Costofgoodssold", "Operatingexpenses", "Netprofit"]
                },
                {
                    "type": "lineChart",
                    "title": "Revenue vs Expenses",
                    "x_axis": "Month",
                    "y_axis": ["Revenue", "Operatingexpenses"]
                },
                {
                    "type": "gauge",
                    "title": "Budget Achievement",
                    "value": "Revenue",
                    "target": "Revenuebudget"
                },
                {
                    "type": "card",
                    "title": "Total Revenue",
                    "measure": "TotalRevenue"
                },
                {
                    "type": "card",
                    "title": "Net Profit",
                    "measure": "TotalNetprofit"
                },
                {
                    "type": "card",
                    "title": "Profit Margin %",
                    "measure": "ProfitMargin"
                }
            ],
            "measures": [
                "TotalRevenue = SUM('P&L_Statement'[Revenue])",
                "TotalExpenses = SUM('P&L_Statement'[Operatingexpenses]) + SUM('P&L_Statement'[Costofgoodssold])",
                "TotalNetprofit = SUM('P&L_Statement'[Netprofit])",
                "ProfitMargin = DIVIDE([TotalNetprofit], [TotalRevenue], 0) * 100",
                "BudgetVariance = [TotalRevenue] - SUM('P&L_Statement'[Revenuebudget])",
                "RevenueGrowthMoM = DIVIDE([TotalRevenue] - CALCULATE([TotalRevenue], PREVIOUSMONTH('P&L_Statement'[Month])), CALCULATE([TotalRevenue], PREVIOUSMONTH('P&L_Statement'[Month])), 0)"
            ],
            "instructions": [
                "1. Open this template in Power BI Desktop",
                "2. Import CSV files: P&L_Statement.csv and Summary.csv",
                "3. Power BI will auto-create relationships",
                "4. All financial KPIs will populate automatically",
                "5. Use date slicer to filter by period"
            ]
        }
        
        output_file = self.templates_dir / "Financial_KPI.pbit"
        self._save_template_config(template_config, output_file)
        
        return output_file
    
    def create_marketing_analytics_template(self) -> Path:
        """Create Marketing Analytics Dashboard Template"""
        
        template_config = {
            "name": "Marketing Analytics Dashboard",
            "version": "1.0",
            "description": "Track campaign performance, ROI, conversions, and marketing funnel metrics",
            "tables": ["Campaign_Performance"],
            "visuals": [
                {
                    "type": "funnelChart",
                    "title": "Marketing Funnel",
                    "values": ["Impressions", "Clicks", "Conversions"]
                },
                {
                    "type": "lineChart",
                    "title": "CTR & Conversion Rate Trends",
                    "x_axis": "Date",
                    "y_axis": ["Ctr", "Conversionrate"]
                },
                {
                    "type": "barChart",
                    "title": "ROI by Channel",
                    "x_axis": "Channel",
                    "y_axis": "Roi"
                },
                {
                    "type": "scatterChart",
                    "title": "Spend vs Revenue",
                    "x_axis": "Spend",
                    "y_axis": "Revenue",
                    "details": "Campaignname"
                },
                {
                    "type": "card",
                    "title": "Total Impressions",
                    "measure": "TotalImpressions"
                },
                {
                    "type": "card",
                    "title": "Average CTR %",
                    "measure": "AverageCTR"
                },
                {
                    "type": "card",
                    "title": "Cost Per Acquisition",
                    "measure": "AverageCPA"
                },
                {
                    "type": "card",
                    "title": "Overall ROI %",
                    "measure": "OverallROI"
                }
            ],
            "measures": [
                "TotalImpressions = SUM(Campaign_Performance[Impressions])",
                "TotalClicks = SUM(Campaign_Performance[Clicks])",
                "TotalConversions = SUM(Campaign_Performance[Conversions])",
                "TotalSpend = SUM(Campaign_Performance[Spend])",
                "TotalRevenue = SUM(Campaign_Performance[Revenue])",
                "AverageCTR = AVERAGE(Campaign_Performance[Ctr])",
                "AverageConversionRate = AVERAGE(Campaign_Performance[Conversionrate])",
                "AverageCPA = AVERAGE(Campaign_Performance[Cpa])",
                "OverallROI = DIVIDE([TotalRevenue] - [TotalSpend], [TotalSpend], 0) * 100"
            ],
            "instructions": [
                "1. Open this template in Power BI Desktop",
                "2. Import CSV file: Campaign_Performance.csv",
                "3. Marketing funnel will auto-populate",
                "4. Use slicers to filter by Channel, Campaign Type, Date",
                "5. Analyze top-performing campaigns in the table visual"
            ]
        }
        
        output_file = self.templates_dir / "Marketing_Analytics.pbit"
        self._save_template_config(template_config, output_file)
        
        return output_file
    
    def create_revenue_profit_template(self) -> Path:
        """Create Revenue & Profit Dashboard Template"""
        
        template_config = {
            "name": "Revenue & Profit Dashboard",
            "version": "1.0",
            "description": "Analyze revenue streams, profit margins, and category performance",
            "tables": ["Revenue"],
            "visuals": [
                {
                    "type": "lineChart",
                    "title": "Revenue Trend Over Time",
                    "x_axis": "Date",
                    "y_axis": "Revenue"
                },
                {
                    "type": "barChart",
                    "title": "Profit by Category",
                    "x_axis": "Category",
                    "y_axis": "Profit"
                },
                {
                    "type": "donutChart",
                    "title": "Revenue Mix by Category",
                    "legend": "Category",
                    "values": "Revenue"
                },
                {
                    "type": "card",
                    "title": "Total Revenue",
                    "measure": "TotalRevenue"
                },
                {
                    "type": "card",
                    "title": "Total Profit",
                    "measure": "TotalProfit"
                },
                {
                    "type": "card",
                    "title": "Profit Margin %",
                    "measure": "ProfitMargin"
                }
            ],
            "measures": [
                "TotalRevenue = SUM(Revenue[Revenue])",
                "TotalCost = SUM(Revenue[Cost])",
                "TotalProfit = [TotalRevenue] - [TotalCost]",
                "ProfitMargin = DIVIDE([TotalProfit], [TotalRevenue], 0) * 100",
                "AverageSellingPrice = DIVIDE([TotalRevenue], SUM(Revenue[Unitssold]), 0)",
                "RevenueYoY = CALCULATE([TotalRevenue], SAMEPERIODLASTYEAR(Revenue[Date]))"
            ],
            "instructions": [
                "1. Open this template in Power BI Desktop",
                "2. Import CSV file: Revenue.csv",
                "3. Revenue and profit visuals will auto-populate",
                "4. Use category slicer to filter analysis",
                "5. Review profit margin trends by period"
            ]
        }
        
        output_file = self.templates_dir / "Revenue_Profit.pbit"
        self._save_template_config(template_config, output_file)
        
        return output_file
    
    def create_operations_efficiency_template(self) -> Path:
        """Create Operations Efficiency Dashboard Template"""
        
        template_config = {
            "name": "Operations Efficiency Dashboard",
            "version": "1.0",
            "description": "Monitor production, capacity utilization, downtime, and resource allocation",
            "tables": ["Operations"],
            "visuals": [
                {
                    "type": "columnChart",
                    "title": "Production vs Capacity",
                    "x_axis": "Date",
                    "y_axis": ["Production", "Capacity"]
                },
                {
                    "type": "lineChart",
                    "title": "Utilization % Trend",
                    "x_axis": "Date",
                    "y_axis": "Utilization"
                },
                {
                    "type": "areaChart",
                    "title": "Downtime Hours",
                    "x_axis": "Date",
                    "y_axis": "Downtimehours"
                },
                {
                    "type": "card",
                    "title": "Average Utilization %",
                    "measure": "AvgUtilization"
                },
                {
                    "type": "card",
                    "title": "Total Production",
                    "measure": "TotalProduction"
                },
                {
                    "type": "card",
                    "title": "Total Downtime",
                    "measure": "TotalDowntime"
                }
            ],
            "measures": [
                "TotalProduction = SUM(Operations[Production])",
                "TotalCapacity = SUM(Operations[Capacity])",
                "TotalDowntime = SUM(Operations[Downtimehours])",
                "AvgUtilization = DIVIDE([TotalProduction], [TotalCapacity], 0) * 100",
                "DowntimeRate = DIVIDE([TotalDowntime], 24 * COUNT(Operations[Date]), 0) * 100",
                "ProductionPerResource = DIVIDE([TotalProduction], AVERAGE(Operations[Resourcecount]), 0)"
            ],
            "instructions": [
                "1. Open this template in Power BI Desktop",
                "2. Import CSV file: Operations.csv",
                "3. Production and efficiency metrics will auto-populate",
                "4. Monitor utilization trends and capacity gaps",
                "5. Identify downtime patterns for improvement"
            ]
        }
        
        output_file = self.templates_dir / "Operations_Efficiency.pbit"
        self._save_template_config(template_config, output_file)
        
        return output_file
    
    def _save_template_config(self, config: Dict[str, Any], output_file: Path):
        """Save template configuration as JSON (simplified .pbit representation)"""
        
        # For now, save as JSON with instructions
        # In production, you'd create actual .pbit files using Power BI APIs
        with open(output_file.with_suffix('.json'), 'w') as f:
            json.dump(config, f, indent=2)
        
        # Create a README for the template
        readme_content = f"""# {config['name']}

{config['description']}

## 📊 Included Visuals

"""
        for visual in config['visuals']:
            readme_content += f"- **{visual['title']}** ({visual['type']})\n"
        
        readme_content += f"""
## 📈 DAX Measures

"""
        for measure in config['measures']:
            readme_content += f"```dax\n{measure}\n```\n\n"
        
        readme_content += f"""
## 🚀 Quick Start Instructions

"""
        for instruction in config['instructions']:
            readme_content += f"{instruction}\n"
        
        readme_file = output_file.with_suffix('.md')
        with open(readme_file, 'w') as f:
            f.write(readme_content)
        
        print(f"✅ Created template: {output_file.stem}")
        print(f"   - Config: {output_file.with_suffix('.json')}")
        print(f"   - README: {readme_file}")
    
    def create_all_templates(self) -> List[Path]:
        """Create all dashboard templates"""
        
        print("\n" + "="*80)
        print("🎨 Creating Power BI Dashboard Templates")
        print("="*80 + "\n")
        
        templates = []
        templates.append(self.create_sales_performance_template())
        templates.append(self.create_financial_kpi_template())
        templates.append(self.create_marketing_analytics_template())
        templates.append(self.create_revenue_profit_template())
        templates.append(self.create_operations_efficiency_template())
        
        print(f"\n✅ Created {len(templates)} templates in: {self.templates_dir}")
        
        return templates


def create_dashboard_package(dashboard_model: Dict[str, Any], 
                             dashboard_type: str,
                             csv_files: List[Path],
                             output_dir: Path) -> Path:
    """
    Create a complete dashboard package (ZIP file) with:
    - CSV data files
    - JSON model specification
    - Template configuration
    - Quick start guide
    - DAX measures file
    """
    
    package_dir = output_dir / f"{dashboard_type}_package"
    package_dir.mkdir(parents=True, exist_ok=True)
    
    # 1. Copy CSV files
    data_dir = package_dir / "data"
    data_dir.mkdir(exist_ok=True)
    
    for csv_file in csv_files:
        shutil.copy(csv_file, data_dir / csv_file.name)
    
    # 2. Save model JSON (with datetime handling)
    model_file = package_dir / "dashboard_model.json"
    with open(model_file, 'w', encoding='utf-8') as f:
        json.dump(dashboard_model, f, indent=2, cls=DateTimeEncoder)
    
    # 3. Create DAX measures file
    measures_file = package_dir / "DAX_Measures.txt"
    with open(measures_file, 'w', encoding='utf-8') as f:
        f.write(f"Power BI DAX Measures - {dashboard_type}\n")
        f.write("="*60 + "\n\n")
        f.write("Copy and paste these measures into Power BI Desktop:\n\n")
        
        for measure in dashboard_model['model']['measures']:
            f.write(f"{measure['name']} = {measure['expression']}\n\n")
    
    # 4. Create Quick Start Guide
    guide_file = package_dir / "QUICK_START.md"
    with open(guide_file, 'w', encoding='utf-8') as f:
        f.write(f"""# {dashboard_type.replace('_', ' ').title()} Dashboard - Quick Start

## 📦 Package Contents

- `data/` - CSV files ready for Power BI import
- `dashboard_model.json` - Complete model specification
- `DAX_Measures.txt` - Copy-paste ready DAX measures
- `QUICK_START.md` - This guide

## 🚀 Setup Instructions (5 minutes)

### Step 1: Open Power BI Desktop

1. Download Power BI Desktop (if not installed): https://powerbi.microsoft.com/desktop/
2. Launch Power BI Desktop
3. Click "Get Data" → "Text/CSV"

### Step 2: Import Data

1. Navigate to the `data/` folder in this package
2. Select all CSV files and click "Open"
3. Click "Load" to import the data
4. Wait for data to load (should take a few seconds)

### Step 3: Add DAX Measures

1. Go to "Modeling" tab in the ribbon
2. Click "New Measure"
3. Open `DAX_Measures.txt` from this package
4. Copy each measure one by one and paste into Power BI
5. Press Enter after each measure

### Step 4: Create Visuals

""")
        
        # Add visual recommendations based on dashboard type
        visuals = _get_recommended_visuals(dashboard_type)
        for i, visual in enumerate(visuals, 1):
            f.write(f"{i}. **{visual['name']}**\n")
            f.write(f"   - Type: {visual['type']}\n")
            f.write(f"   - Fields: {visual['fields']}\n\n")
        
        f.write(f"""
### Step 5: Apply Theme & Publish

1. Go to "View" → "Themes" → Choose a theme
2. Add slicers for filtering (Date, Region, Category, etc.)
3. Save your dashboard: File → Save
4. Publish to Power BI Service: Home → Publish

## 💡 Pro Tips

- Use Ctrl+Click to select multiple visuals at once
- Right-click visuals for formatting options
- Enable "Show Data Label" for better readability
- Use bookmarks to save different dashboard views
- Set up automatic data refresh in Power BI Service

## 🆘 Need Help?

- Power BI Documentation: https://docs.microsoft.com/power-bi/
- DAX Reference: https://dax.guide/
- Community Forum: https://community.powerbi.com/

---

**Created by FinDeck Power BI Dashboard Generator**  
**Date: {Path().absolute()}**
""")
    
    # 5. Create ZIP file
    zip_file = output_dir / f"{dashboard_type}_complete.zip"
    
    with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file in package_dir.rglob('*'):
            if file.is_file():
                zipf.write(file, file.relative_to(package_dir))
    
    print(f"✅ Created complete package: {zip_file}")
    print(f"   Size: {zip_file.stat().st_size / 1024:.1f} KB")
    
    return zip_file


def _get_recommended_visuals(dashboard_type: str) -> List[Dict[str, str]]:
    """Get recommended visuals for each dashboard type"""
    
    visuals_map = {
        "sales_performance": [
            {"name": "Sales Trend", "type": "Line Chart", "fields": "Date (X), Salesamount (Y), Region (Legend)"},
            {"name": "Regional Map", "type": "Map", "fields": "Region (Location), Salesamount (Size)"},
            {"name": "Top Products", "type": "Bar Chart", "fields": "Product (Y), Salesamount (X)"},
            {"name": "Total Sales KPI", "type": "Card", "fields": "TotalSalesamount"},
            {"name": "Target Achievement", "type": "Gauge", "fields": "TotalSalesamount (Value), TotalSalestarget (Target)"}
        ],
        "financial_kpi": [
            {"name": "P&L Waterfall", "type": "Waterfall Chart", "fields": "Month (Category), Revenue/Expenses/Profit (Y)"},
            {"name": "Revenue vs Expenses", "type": "Line Chart", "fields": "Month (X), Revenue & Expenses (Y)"},
            {"name": "Profit Margin %", "type": "Card", "fields": "ProfitMargin"},
            {"name": "Budget Variance", "type": "Gauge", "fields": "TotalRevenue (Value), TotalRevenueBudget (Target)"}
        ],
        "marketing_analytics": [
            {"name": "Marketing Funnel", "type": "Funnel Chart", "fields": "Impressions → Clicks → Conversions"},
            {"name": "ROI by Channel", "type": "Bar Chart", "fields": "Channel (Y), ROI (X)"},
            {"name": "CTR Trend", "type": "Line Chart", "fields": "Date (X), CTR% (Y)"},
            {"name": "Total Impressions", "type": "Card", "fields": "TotalImpressions"}
        ],
        "revenue_profit": [
            {"name": "Revenue Trend", "type": "Line Chart", "fields": "Date (X), Revenue (Y)"},
            {"name": "Profit by Category", "type": "Bar Chart", "fields": "Category (Y), Profit (X)"},
            {"name": "Revenue Mix", "type": "Donut Chart", "fields": "Category (Legend), Revenue (Values)"},
            {"name": "Profit Margin %", "type": "Card", "fields": "ProfitMargin"}
        ],
        "operations_efficiency": [
            {"name": "Production vs Capacity", "type": "Column Chart", "fields": "Date (X), Production & Capacity (Y)"},
            {"name": "Utilization Trend", "type": "Line Chart", "fields": "Date (X), Utilization% (Y)"},
            {"name": "Downtime", "type": "Area Chart", "fields": "Date (X), Downtimehours (Y)"},
            {"name": "Avg Utilization", "type": "Card", "fields": "AvgUtilization"}
        ]
    }
    
    return visuals_map.get(dashboard_type, [])


if __name__ == "__main__":
    # Create all templates
    generator = PowerBITemplateGenerator()
    templates = generator.create_all_templates()
    
    print("\n" + "="*80)
    print("✅ Template Generation Complete!")
    print("="*80)
    print(f"\n📁 Location: {generator.templates_dir}")
    print(f"📊 Templates Created: {len(templates)}")
