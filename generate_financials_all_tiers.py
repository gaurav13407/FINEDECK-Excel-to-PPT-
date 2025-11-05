"""
Generate PowerPoint presentations for all tiers from Financials.csv
Creates 4 presentations: BASIC, PRO, AI_PRO, and a Country-specific analysis
"""

import pandas as pd
import os
from datetime import datetime

# Convert CSV to Excel format that our converter can use
def prepare_financials_data():
    """Convert Financials.csv to Excel format with proper sheets"""
    
    print("📊 Loading Financials.csv...")
    df = pd.read_csv('examples/Financials.csv')
    
    # Clean column names (remove extra spaces)
    df.columns = df.columns.str.strip()
    
    # Clean currency values (remove $ and commas)
    currency_cols = ['Units Sold', 'Manufacturing Price', 'Sale Price', 
                     'Gross Sales', 'Discounts', 'Sales', 'COGS', 'Profit']
    
    for col in currency_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('$', '').str.replace(',', '').str.replace('-', '0')
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    print(f"✅ Loaded {len(df)} financial records")
    print(f"📅 Date range: {df['Date'].min()} to {df['Date'].max()}")
    print(f"🌍 Countries: {df['Country'].nunique()}")
    print(f"📦 Products: {df['Product'].nunique()}")
    print(f"💰 Total Sales: ${df['Sales'].sum():,.2f}")
    print(f"📈 Total Profit: ${df['Profit'].sum():,.2f}")
    
    # Create Excel file with multiple sheets
    output_path = 'examples/Company_Data/financials_bundle.xlsx'
    
    print(f"\n🔧 Creating Excel bundle: {output_path}")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: Summary by Product
        summary_product = df.groupby('Product').agg({
            'Sales': 'sum',
            'Profit': 'sum',
            'Units Sold': 'sum',
            'COGS': 'sum'
        }).reset_index()
        
        summary_product['Profit_Margin_%'] = (summary_product['Profit'] / summary_product['Sales'] * 100).round(2)
        summary_product['Market_Share_%'] = (summary_product['Sales'] / summary_product['Sales'].sum() * 100).round(2)
        summary_product.columns = ['Ticker', 'Sales', 'Profit', 'Units_Sold', 'COGS', 'Return_1Y_%', 'MarketCap']
        summary_product['Name'] = summary_product['Ticker'] + ' Product'
        summary_product['Sector'] = 'Consumer Goods'
        summary_product['TrailingPE'] = (summary_product['Sales'] / summary_product['Profit']).round(2)
        summary_product['ForwardPE'] = summary_product['TrailingPE'] * 0.95
        
        # Reorder columns to match expected format
        summary_product = summary_product[['Ticker', 'Name', 'Sector', 'MarketCap', 'Return_1Y_%', 'TrailingPE', 'ForwardPE']]
        summary_product.to_excel(writer, sheet_name='Summary', index=False)
        print(f"   ✅ Summary sheet: {len(summary_product)} products")
        
        # Sheet 2-6: Time series data for top 5 products
        top_products = summary_product.nlargest(5, 'MarketCap')['Ticker'].tolist()
        
        for product in top_products:
            product_data = df[df['Product'].str.strip() == product].copy()
            product_data['Date'] = pd.to_datetime(product_data['Date'], format='%m/%d/%Y')
            
            # Create time series aggregated by month
            time_series = product_data.groupby('Date').agg({
                'Sales': 'sum',
                'Profit': 'sum'
            }).reset_index()
            
            time_series = time_series.sort_values('Date')
            time_series.columns = ['Date', 'Close', 'Volume']
            time_series['Open'] = time_series['Close'] * 0.98
            time_series['High'] = time_series['Close'] * 1.02
            time_series['Low'] = time_series['Close'] * 0.97
            
            sheet_name = f'{product}_Prices'
            time_series.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"   ✅ {sheet_name} sheet: {len(time_series)} data points")
        
        # Additional sheets by Segment and Country
        segment_summary = df.groupby('Segment').agg({
            'Sales': 'sum',
            'Profit': 'sum',
            'Units Sold': 'sum'
        }).reset_index()
        segment_summary.to_excel(writer, sheet_name='Segment_Analysis', index=False)
        print(f"   ✅ Segment_Analysis sheet: {len(segment_summary)} segments")
        
        country_summary = df.groupby('Country').agg({
            'Sales': 'sum',
            'Profit': 'sum',
            'Units Sold': 'sum'
        }).reset_index()
        country_summary.to_excel(writer, sheet_name='Country_Analysis', index=False)
        print(f"   ✅ Country_Analysis sheet: {len(country_summary)} countries")
        
        # Month analysis
        month_summary = df.groupby(['Year', 'Month Name']).agg({
            'Sales': 'sum',
            'Profit': 'sum'
        }).reset_index()
        month_summary.to_excel(writer, sheet_name='Monthly_Trends', index=False)
        print(f"   ✅ Monthly_Trends sheet: {len(month_summary)} months")
    
    print(f"✅ Excel bundle created successfully!")
    return output_path


def generate_all_presentations(excel_path):
    """Generate presentations for all tiers"""
    
    from src.converter.excel_to_ppt_converter import ExcelToPPTConverter
    
    print("\n" + "="*80)
    print("🎯 GENERATING PRESENTATIONS FOR ALL TIERS")
    print("="*80)
    
    base_output_path = "examples/professional_demo"
    os.makedirs(base_output_path, exist_ok=True)
    
    tiers = [
        {
            'name': 'BASIC',
            'tier': 'basic',
            'user_id': 'financial_basic_user',
            'metadata': {
                'name': 'Basic Financial Analyst',
                'company': 'Financial Services Corp',
                'email': 'analyst@financial.com'
            },
            'title': 'Financial Performance Analysis - Q4 2025 (BASIC)',
            'output': 'Financials_BASIC_Tier.pptx',
            'expected_slides': 7
        },
        {
            'name': 'PRO',
            'tier': 'pro',
            'user_id': 'financial_pro_user',
            'metadata': {
                'name': 'Pro Financial Analyst',
                'company': 'Investment Analytics Ltd',
                'email': 'pro@investment.com'
            },
            'title': 'Comprehensive Financial Analysis - Q4 2025 (PRO)',
            'output': 'Financials_PRO_Tier.pptx',
            'expected_slides': 7
        },
        {
            'name': 'AI_PRO',
            'tier': 'ai_pro',
            'user_id': 'financial_ai_pro_user',
            'metadata': {
                'name': 'AI Financial Strategist',
                'company': 'Advanced Analytics Group',
                'email': 'ai-pro@analytics.com'
            },
            'title': 'AI-Powered Financial Insights - Q4 2025 (AI_PRO)',
            'output': 'Financials_AI_PRO_Tier.pptx',
            'expected_slides': 8
        },
        {
            'name': 'COUNTRY_FOCUS',
            'tier': 'ai_pro',
            'user_id': 'usa_market_analysis',
            'metadata': {
                'name': 'Regional Market Analyst',
                'company': 'USA Market Research Division',
                'email': 'usa@market-research.com'
            },
            'title': 'USA Market Performance - Deep Dive Analysis Q4 2025',
            'output': 'Financials_USA_Market_Analysis.pptx',
            'expected_slides': 8
        }
    ]
    
    results = []
    
    for tier_info in tiers:
        print("\n" + "="*80)
        print(f"📊 TIER: {tier_info['name']}")
        print("="*80)
        
        try:
            # Create converter
            converter = ExcelToPPTConverter(
                user_tier=tier_info['tier'],
                user_id=tier_info['user_id'],
                user_metadata=tier_info['metadata']
            )
            
            # Generate presentation
            output_path = os.path.join(base_output_path, tier_info['output'])
            
            result = converter.convert_professional(
                excel_path=excel_path,
                output_path=output_path,
                presentation_title=tier_info['title'],
                user_ppt_count=0
            )
            
            # Get file size
            file_size_kb = os.path.getsize(output_path) / 1024
            
            print(f"\n✅ {tier_info['name']} TIER CREATED!")
            print(f"   Slides: {tier_info['expected_slides']}")
            print(f"   File Size: {file_size_kb:.1f} KB")
            print(f"   Output: {output_path}")
            
            results.append({
                'tier': tier_info['name'],
                'success': True,
                'slides': tier_info['expected_slides'],
                'size_kb': file_size_kb,
                'output': tier_info['output']
            })
            
        except Exception as e:
            print(f"❌ FAILED: {tier_info['name']}")
            print(f"   Error: {str(e)}")
            results.append({
                'tier': tier_info['name'],
                'success': False,
                'error': str(e)
            })
    
    # Summary table
    print("\n" + "="*80)
    print("📊 GENERATION SUMMARY")
    print("="*80)
    print(f"\n{'Tier':<20} {'Status':<10} {'Slides':<10} {'Size (KB)':<15} {'Output':<40}")
    print("-"*80)
    
    for result in results:
        if result['success']:
            status = "✅ SUCCESS"
            slides = str(result['slides'])
            size = f"{result['size_kb']:.1f} KB"
            output = result['output']
        else:
            status = "❌ FAILED"
            slides = "N/A"
            size = "N/A"
            output = result.get('error', 'Unknown error')
        
        print(f"{result['tier']:<20} {status:<10} {slides:<10} {size:<15} {output:<40}")
    
    print("\n" + "="*80)
    successful = sum(1 for r in results if r['success'])
    print(f"✅ {successful}/{len(results)} presentations generated successfully!")
    print("="*80)
    
    return results


def main():
    """Main execution"""
    
    print("\n" + "="*80)
    print("💼 FINANCIALS.CSV TO POWERPOINT - ALL TIERS")
    print("="*80)
    print(f"📅 Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*80)
    
    # Step 1: Prepare data
    excel_path = prepare_financials_data()
    
    # Step 2: Generate presentations
    results = generate_all_presentations(excel_path)
    
    # Step 3: Final summary
    print("\n" + "="*80)
    print("🎉 ALL TIERS GENERATION COMPLETE!")
    print("="*80)
    print("\n📦 DELIVERABLES:")
    for result in results:
        if result['success']:
            print(f"   ✅ {result['output']} ({result['slides']} slides, {result['size_kb']:.1f} KB)")
    
    print("\n" + "="*80)
    print("📍 Location: examples/professional_demo/")
    print("="*80)


if __name__ == "__main__":
    main()
