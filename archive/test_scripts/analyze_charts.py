"""
Detailed Chart Analysis - Extract chart data to prove they exist
"""

from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE

def analyze_chart_details(ppt_path):
    """Extract detailed chart information"""
    
    print(f"\n{'='*80}")
    print(f"🔍 DETAILED CHART ANALYSIS: {ppt_path}")
    print(f"{'='*80}\n")
    
    prs = Presentation(ppt_path)
    
    chart_types = {
        5: "PIE",
        51: "COLUMN_CLUSTERED",
        57: "BAR_CLUSTERED",
        65: "LINE_MARKERS",
        73: "LINE",
        69: "DOUGHNUT"
    }
    
    total_charts = 0
    
    for slide_num, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if shape.shape_type == 3:  # Chart
                total_charts += 1
                chart = shape.chart
                chart_type_num = int(chart.chart_type)
                chart_type_name = chart_types.get(chart_type_num, f"Type_{chart_type_num}")
                
                print(f"📊 SLIDE {slide_num} - Chart #{total_charts}")
                print(f"   Chart Type: {chart_type_name} ({chart_type_num})")
                print(f"   Has Legend: {chart.has_legend}")
                print(f"   Series Count: {len(chart.series)}")
                
                # Get series data
                for i, series in enumerate(chart.series, 1):
                    try:
                        series_name = series.name
                        print(f"   Series {i}: {series_name}")
                        
                        # Try to get colors
                        try:
                            fill = series.format.fill
                            if fill.type == 1:  # Solid fill
                                rgb = fill.fore_color.rgb
                                print(f"      Color: RGB{rgb}")
                        except:
                            pass
                    except:
                        pass
                
                print()
    
    if total_charts == 0:
        print("❌ NO CHARTS FOUND IN THIS FILE\n")
    else:
        print(f"✅ TOTAL CHARTS: {total_charts}\n")
    
    return total_charts

if __name__ == "__main__":
    # Analyze the comparison demo
    analyze_chart_details("test_output/chart_comparison_demo.pptx")
    analyze_chart_details("test_output/portfolio_with_pie_charts.pptx")
