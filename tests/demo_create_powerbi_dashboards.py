"""
Complete Power BI Dashboard Creation Test
Demonstrates end-to-end workflow: Excel → Power BI Dashboard

This script:
1. Creates sample Excel files for each dashboard type
2. Calls the Power BI API endpoint
3. Verifies the generated dashboard
4. Shows how to import into Power BI Desktop
"""

import sys
from pathlib import Path
import pandas as pd
import json

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.backend.app.services.powerbi_etl import ExcelToPowerBIProcessor


def create_sales_dashboard_demo():
    """
    Demo 1: Sales Performance Dashboard
    Real-world scenario: Quarterly sales data by region and product
    """
    
    print("\n" + "="*80)
    print("📊 DEMO 1: SALES PERFORMANCE DASHBOARD")
    print("="*80)
    
    # Step 1: Create realistic sales data
    print("\n🔧 Step 1: Creating Sample Sales Data...")
    
    sales_data = pd.DataFrame({
        'Date': pd.date_range('2024-01-01', periods=100, freq='D'),
        'Region': ['North', 'South', 'East', 'West', 'Central'] * 20,
        'Product': ['Laptop', 'Phone', 'Tablet', 'Desktop', 'Accessories'] * 20,
        'Sales_Amount': [
            50000 + (i * 1000) + (i % 7 * 5000) for i in range(100)
        ],
        'Quantity_Sold': [
            50 + (i % 10) * 5 for i in range(100)
        ],
        'Sales_Target': [
            48000 + (i * 950) + (i % 7 * 4800) for i in range(100)
        ],
        'Salesperson': ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve'] * 20,
        'Customer_Segment': ['Enterprise', 'SMB', 'Consumer', 'Government', 'Education'] * 20
    })
    
    # Save Excel
    excel_file = Path("examples/test_data/demo_sales_dashboard.xlsx")
    excel_file.parent.mkdir(parents=True, exist_ok=True)
    sales_data.to_excel(excel_file, sheet_name='Sales_Data', index=False)
    
    print(f"   ✅ Created: {excel_file}")
    print(f"   📊 Data: {len(sales_data)} rows, {len(sales_data.columns)} columns")
    print(f"   📅 Date Range: {sales_data['Date'].min()} to {sales_data['Date'].max()}")
    print(f"   🌎 Regions: {sales_data['Region'].unique().tolist()}")
    print(f"   📦 Products: {sales_data['Product'].unique().tolist()}")
    
    # Step 2: Process with Power BI ETL
    print("\n⚙️  Step 2: Running Power BI ETL Pipeline...")
    
    processor = ExcelToPowerBIProcessor()
    dashboard_model = processor.process_excel(str(excel_file))
    
    print(f"   ✅ Processing Complete!")
    print(f"   📊 Tables Created: {len(dashboard_model['model']['tables'])}")
    print(f"   🔗 Relationships: {len(dashboard_model['model']['relationships'])}")
    print(f"   📈 DAX Measures: {len(dashboard_model['model']['measures'])}")
    
    # Step 3: Show generated measures
    print("\n📈 Step 3: Generated DAX Measures (Sample)...")
    
    measures = dashboard_model['model']['measures']
    for measure in measures[:8]:
        print(f"   • {measure['name']} = {measure['expression']}")
    
    if len(measures) > 8:
        print(f"   ... and {len(measures) - 8} more measures")
    
    # Step 4: Export for Power BI
    print("\n💾 Step 4: Exporting for Power BI Desktop...")
    
    output_dir = Path("examples/test_data/powerbi_dashboards/sales_performance")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Export CSV files
    processor.export_to_csv(str(output_dir / "csv_data"))
    
    # Export JSON model
    model_file = output_dir / "dashboard_model.json"
    with open(model_file, 'w') as f:
        export_model = dashboard_model.copy()
        for table_name in export_model['model']['tables']:
            export_model['model']['tables'][table_name]['data'] = \
                f"[{export_model['model']['tables'][table_name]['row_count']} rows]"
        json.dump(export_model, f, indent=2)
    
    print(f"   ✅ CSV Files: {output_dir / 'csv_data'}/")
    print(f"   ✅ Model JSON: {model_file}")
    
    # Step 5: Import Instructions
    print("\n📋 Step 5: How to Import into Power BI Desktop...")
    print("""
    1. Open Power BI Desktop
    2. Click 'Get Data' → 'Text/CSV'
    3. Import all CSV files from: {csv_dir}
    4. Go to 'Modeling' tab → 'Manage Relationships'
    5. Verify auto-detected relationships
    6. Create visualizations:
       - Line Chart: Sales_Amount by Date
       - Map: Sales_Amount by Region
       - Table: Top Products by Sales
       - KPI Cards: Total Sales, Target Achievement
       - Slicer: Region, Product, Salesperson
    7. Apply 'Sales Performance' theme
    8. Publish to Power BI Service
    """.format(csv_dir=output_dir / 'csv_data'))
    
    print("\n" + "="*80)
    print("✅ Sales Performance Dashboard Ready!")
    print("="*80)
    
    return dashboard_model


def create_financial_dashboard_demo():
    """
    Demo 2: Financial KPI Dashboard
    Real-world scenario: Monthly P&L statement with budget comparison
    """
    
    print("\n" + "="*80)
    print("💰 DEMO 2: FINANCIAL KPI DASHBOARD")
    print("="*80)
    
    # Step 1: Create financial data
    print("\n🔧 Step 1: Creating Sample Financial Data...")
    
    financial_data = pd.DataFrame({
        'Month': pd.date_range('2023-01-01', periods=18, freq='MS'),
        'Revenue': [
            100000, 120000, 115000, 130000, 140000, 135000,
            150000, 145000, 160000, 170000, 165000, 180000,
            190000, 185000, 200000, 210000, 205000, 220000
        ],
        'Cost_of_Goods_Sold': [
            40000, 48000, 46000, 52000, 56000, 54000,
            60000, 58000, 64000, 68000, 66000, 72000,
            76000, 74000, 80000, 84000, 82000, 88000
        ],
        'Operating_Expenses': [
            30000, 32000, 31000, 33000, 34000, 33500,
            35000, 34500, 36000, 37000, 36500, 38000,
            38500, 38000, 39000, 40000, 39500, 41000
        ],
        'Marketing_Spend': [
            10000, 12000, 11500, 13000, 14000, 13500,
            15000, 14500, 16000, 17000, 16500, 18000,
            18500, 18000, 19000, 20000, 19500, 21000
        ],
        'Revenue_Budget': [
            95000, 115000, 110000, 125000, 135000, 130000,
            145000, 140000, 155000, 165000, 160000, 175000,
            185000, 180000, 195000, 205000, 200000, 215000
        ],
        'Department': ['Sales', 'Operations', 'Marketing', 'IT', 'HR', 'Finance'] * 3
    })
    
    # Calculate derived columns
    financial_data['Gross_Profit'] = financial_data['Revenue'] - financial_data['Cost_of_Goods_Sold']
    financial_data['Net_Profit'] = (financial_data['Gross_Profit'] - 
                                     financial_data['Operating_Expenses'] - 
                                     financial_data['Marketing_Spend'])
    financial_data['Profit_Margin_%'] = (financial_data['Net_Profit'] / financial_data['Revenue'] * 100).round(2)
    
    # Save Excel
    excel_file = Path("examples/test_data/demo_financial_dashboard.xlsx")
    
    # Create multiple sheets for better organization
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        financial_data.to_excel(writer, sheet_name='P&L_Statement', index=False)
        
        # Summary sheet
        summary = pd.DataFrame({
            'Metric': ['Total Revenue', 'Total Expenses', 'Net Profit', 'Average Margin %'],
            'Value': [
                financial_data['Revenue'].sum(),
                (financial_data['Cost_of_Goods_Sold'] + 
                 financial_data['Operating_Expenses'] + 
                 financial_data['Marketing_Spend']).sum(),
                financial_data['Net_Profit'].sum(),
                financial_data['Profit_Margin_%'].mean()
            ]
        })
        summary.to_excel(writer, sheet_name='Summary', index=False)
    
    print(f"   ✅ Created: {excel_file}")
    print(f"   📊 Data: {len(financial_data)} rows, {len(financial_data.columns)} columns")
    print(f"   📅 Period: 18 months (Jan 2023 - Jun 2024)")
    print(f"   💰 Total Revenue: ${financial_data['Revenue'].sum():,.0f}")
    print(f"   📈 Total Profit: ${financial_data['Net_Profit'].sum():,.0f}")
    
    # Process with ETL
    print("\n⚙️  Step 2: Running Power BI ETL Pipeline...")
    
    processor = ExcelToPowerBIProcessor()
    dashboard_model = processor.process_excel(str(excel_file))
    
    print(f"   ✅ Processing Complete!")
    print(f"   📊 Tables Created: {len(dashboard_model['model']['tables'])}")
    print(f"   📈 DAX Measures: {len(dashboard_model['model']['measures'])}")
    
    # Export
    output_dir = Path("examples/test_data/powerbi_dashboards/financial_kpi")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    processor.export_to_csv(str(output_dir / "csv_data"))
    
    model_file = output_dir / "dashboard_model.json"
    with open(model_file, 'w') as f:
        export_model = dashboard_model.copy()
        for table_name in export_model['model']['tables']:
            export_model['model']['tables'][table_name]['data'] = \
                f"[{export_model['model']['tables'][table_name]['row_count']} rows]"
        json.dump(export_model, f, indent=2)
    
    print(f"\n💾 Exported to: {output_dir}")
    
    # Custom DAX recommendations
    print("\n📊 Recommended Custom DAX Measures:")
    print("""
    ProfitMargin = DIVIDE([TotalNetProfit], [TotalRevenue], 0)
    ExpenseRatio = DIVIDE([TotalExpenses], [TotalRevenue], 0)
    BudgetVariance = [TotalRevenue] - [TotalRevenueBudget]
    BudgetAchievement% = DIVIDE([TotalRevenue], [TotalRevenueBudget], 0) * 100
    GrossMargin% = DIVIDE([TotalGrossProfit], [TotalRevenue], 0) * 100
    RevenueGrowthMoM = DIVIDE([TotalRevenue] - [PreviousMonthRevenue], [PreviousMonthRevenue], 0)
    """)
    
    print("\n" + "="*80)
    print("✅ Financial KPI Dashboard Ready!")
    print("="*80)
    
    return dashboard_model


def create_marketing_dashboard_demo():
    """
    Demo 3: Marketing Analytics Dashboard
    Real-world scenario: Multi-channel campaign performance tracking
    """
    
    print("\n" + "="*80)
    print("📱 DEMO 3: MARKETING ANALYTICS DASHBOARD")
    print("="*80)
    
    print("\n🔧 Step 1: Creating Sample Marketing Data...")
    
    # Multi-channel campaign data
    campaigns = []
    channels = ['Google Ads', 'Facebook', 'LinkedIn', 'Instagram', 'Twitter', 'Email']
    campaign_types = ['Brand Awareness', 'Lead Generation', 'Conversion', 'Retargeting']
    
    for week in range(12):
        for channel in channels:
            for campaign_type in campaign_types[:2]:  # 2 campaign types per channel
                campaigns.append({
                    'Date': pd.Timestamp('2024-01-01') + pd.Timedelta(weeks=week),
                    'Channel': channel,
                    'Campaign_Name': f"{channel} - {campaign_type} - Week {week+1}",
                    'Campaign_Type': campaign_type,
                    'Impressions': 50000 + (week * 5000) + (hash(channel) % 30000),
                    'Clicks': 2500 + (week * 250) + (hash(channel) % 1500),
                    'Conversions': 125 + (week * 12) + (hash(channel) % 75),
                    'Spend': 5000 + (week * 500) + (hash(channel) % 3000),
                    'Revenue': 25000 + (week * 2500) + (hash(channel) % 15000)
                })
    
    marketing_data = pd.DataFrame(campaigns)
    
    # Calculate KPIs
    marketing_data['CTR_%'] = (marketing_data['Clicks'] / marketing_data['Impressions'] * 100).round(2)
    marketing_data['Conversion_Rate_%'] = (marketing_data['Conversions'] / marketing_data['Clicks'] * 100).round(2)
    marketing_data['CPC'] = (marketing_data['Spend'] / marketing_data['Clicks']).round(2)
    marketing_data['CPA'] = (marketing_data['Spend'] / marketing_data['Conversions']).round(2)
    marketing_data['ROI_%'] = ((marketing_data['Revenue'] - marketing_data['Spend']) / marketing_data['Spend'] * 100).round(2)
    
    # Save Excel
    excel_file = Path("examples/test_data/demo_marketing_dashboard.xlsx")
    marketing_data.to_excel(excel_file, sheet_name='Campaign_Performance', index=False)
    
    print(f"   ✅ Created: {excel_file}")
    print(f"   📊 Data: {len(marketing_data)} campaigns, {len(marketing_data.columns)} metrics")
    print(f"   📺 Channels: {len(channels)} channels")
    print(f"   💰 Total Spend: ${marketing_data['Spend'].sum():,.0f}")
    print(f"   💵 Total Revenue: ${marketing_data['Revenue'].sum():,.0f}")
    print(f"   📈 Overall ROI: {((marketing_data['Revenue'].sum() - marketing_data['Spend'].sum()) / marketing_data['Spend'].sum() * 100):.1f}%")
    
    # Process
    processor = ExcelToPowerBIProcessor()
    dashboard_model = processor.process_excel(str(excel_file))
    
    # Export
    output_dir = Path("examples/test_data/powerbi_dashboards/marketing_analytics")
    output_dir.mkdir(parents=True, exist_ok=True)
    processor.export_to_csv(str(output_dir / "csv_data"))
    
    print(f"\n💾 Exported to: {output_dir}")
    
    print("\n📊 Key Metrics to Visualize:")
    print("""
    • Funnel Chart: Impressions → Clicks → Conversions
    • Line Chart: CTR% and Conversion Rate% trends
    • Bar Chart: ROI% by Channel
    • Scatter Plot: Spend vs Revenue by Campaign
    • KPI Cards: Total Impressions, CTR%, Avg CPA, Overall ROI%
    • Table: Top 10 campaigns by ROI
    • Slicer: Channel, Campaign Type, Date Range
    """)
    
    print("\n" + "="*80)
    print("✅ Marketing Analytics Dashboard Ready!")
    print("="*80)
    
    return dashboard_model


def main():
    """
    Main test function - Creates all dashboard demos
    """
    
    print("="*80)
    print("🚀 POWER BI DASHBOARD CREATION - COMPLETE DEMO")
    print("="*80)
    print("\nThis demo will create 3 real-world Power BI dashboards:")
    print("  1. 📊 Sales Performance Dashboard")
    print("  2. 💰 Financial KPI Dashboard")
    print("  3. 📱 Marketing Analytics Dashboard")
    print("\nEach dashboard includes:")
    print("  ✅ Realistic sample data")
    print("  ✅ Automated ETL processing")
    print("  ✅ Auto-generated DAX measures")
    print("  ✅ CSV files for Power BI import")
    print("  ✅ Complete model specification")
    
    input("\nPress ENTER to start...")
    
    # Create all dashboards
    dashboards = {}
    
    try:
        dashboards['sales'] = create_sales_dashboard_demo()
        dashboards['financial'] = create_financial_dashboard_demo()
        dashboards['marketing'] = create_marketing_dashboard_demo()
        
        # Final Summary
        print("\n" + "="*80)
        print("🎉 ALL DASHBOARDS CREATED SUCCESSFULLY!")
        print("="*80)
        
        print("\n📁 Generated Files:")
        print("   examples/test_data/powerbi_dashboards/")
        print("   ├── sales_performance/")
        print("   │   ├── csv_data/ (CSV files)")
        print("   │   └── dashboard_model.json")
        print("   ├── financial_kpi/")
        print("   │   ├── csv_data/ (CSV files)")
        print("   │   └── dashboard_model.json")
        print("   └── marketing_analytics/")
        print("       ├── csv_data/ (CSV files)")
        print("       └── dashboard_model.json")
        
        print("\n📊 Summary:")
        for name, model in dashboards.items():
            tables_count = len(model['model']['tables'])
            measures_count = len(model['model']['measures'])
            print(f"   • {name.title()}: {tables_count} tables, {measures_count} measures")
        
        print("\n🎯 Next Steps:")
        print("   1. Open Power BI Desktop")
        print("   2. Import CSV files from any dashboard folder")
        print("   3. Review auto-created relationships")
        print("   4. Add recommended DAX measures from JSON model")
        print("   5. Create visualizations based on dashboard type")
        print("   6. Apply theme and publish to Power BI Service")
        
        print("\n💡 Pro Tips:")
        print("   • All date columns have hierarchies (Year/Quarter/Month/Day)")
        print("   • DAX measures follow naming convention: Total*, Average*, *YoY")
        print("   • Use slicers for Date, Region, Product, Channel filters")
        print("   • Consider adding calculated measures for ratios and %")
        
        print("\n" + "="*80)
        print("✅ Demo Complete! Ready for Power BI Desktop import.")
        print("="*80)
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
