"""
Enhanced Chart Builder with Multiple Chart Types (v2.0 - NaN/Inf Protection)
=================================================
Supports:
- Pie Charts (sector distribution, portfolio allocation)
- Bar Charts (top performers, comparisons)
- Line Charts (trends, time-series)
- Auto-detection of best chart type
- NaN/Inf value protection for all chart types

VERSION: 2024-11-08-FIXED
"""

print("🔥 LOADING enhanced_charts.py - VERSION 2024-11-08-FIXED")

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import pandas as pd
import numpy as np


class EnhancedChartBuilder:
    """Build multiple chart types with smart auto-detection"""
    
    def __init__(self, template_colors):
        self.template_colors = template_colors
        self.chart_colors = template_colors.get('chart_colors', [
            (79, 129, 189),
            (194, 57, 52),
            (155, 187, 89),
            (128, 100, 162),
            (75, 172, 198),
            (247, 150, 70)
        ])
    
    def detect_chart_type(self, df):
        """
        Auto-detect best chart type for FINANCE data
        
        Finance-optimized detection with priority order:
        1. Time-series data → LINE chart (highest priority)
        2. Portfolio/Allocation → PIE chart
        3. Performance/Comparison → COLUMN chart
        4. Default → COLUMN chart
        
        Returns: 'pie', 'bar', 'line', or 'column'
        """
        # Check column names for finance-specific patterns
        col_names_lower = ' '.join([str(col).lower() for col in df.columns])
        
        # PRIORITY 1: Time-series/Trend data → LINE chart (check FIRST)
        line_keywords = ['date', 'time', 'quarter', 'month', 'year', 'q1', 'q2', 'q3', 'q4',
                        'trend', 'growth', 'ytd', 'mtd', 'jan', 'feb', 'mar', 'apr', 'may', 
                        'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
        if any(keyword in col_names_lower for keyword in line_keywords):
            return 'line'
        
        # PRIORITY 2: Portfolio/Allocation/Distribution → PIE chart
        pie_keywords = ['allocation', 'portfolio', 'sector', 'distribution', 
                        'breakdown', 'composition', 'share', 'weight']
        if any(keyword in col_names_lower for keyword in pie_keywords):
            return 'pie'
        
        # PRIORITY 3: Check for "%" but NOT with time/ranking keywords
        if ('percent' in col_names_lower or '%' in col_names_lower):
            # If it has % but also has ranking/performance keywords, use COLUMN
            ranking_keywords = ['top', 'bottom', 'rank', 'performance', 'return']
            if any(keyword in col_names_lower for keyword in ranking_keywords):
                return 'column'
            else:
                return 'pie'
        
        # PRIORITY 4: Top performers/Rankings/Comparisons → COLUMN chart
        comparison_keywords = ['top', 'bottom', 'rank', 'performance', 'revenue', 'sales', 
                              'profit', 'earnings', 'return', 'yield', 'stock', 'company']
        if any(keyword in col_names_lower for keyword in comparison_keywords):
            return 'column'
        
        # DEFAULT: COLUMN charts (vertical) for all other finance comparisons
        return 'column'
    
    def create_pie_chart(self, slide, left, top, width, height, data_dict, title="Distribution"):
        """
        Create pie chart for distribution/allocation data (FINANCE OPTIMIZED)
        
        Args:
            slide: PowerPoint slide
            left, top, width, height: Position and size in inches
            data_dict: Dictionary {category: value}
            title: Chart title
        """
        # Clean NaN/Inf values from data
        cleaned_dict = {}
        for key, value in data_dict.items():
            try:
                # Convert to Python float first
                if pd.isna(value) or np.isnan(float(value)) or np.isinf(float(value)):
                    cleaned_dict[key] = float(0)
                else:
                    cleaned_dict[key] = float(value)
            except (ValueError, TypeError):
                cleaned_dict[key] = float(0)
        
        # Prepare chart data
        chart_data = CategoryChartData()
        chart_data.categories = list(cleaned_dict.keys())
        chart_data.add_series('Values', list(cleaned_dict.values()))
        
        # Add chart
        chart_placeholder = slide.shapes.add_chart(
            XL_CHART_TYPE.PIE,
            left, top, width, height,  # Already in EMU units, don't wrap in Inches()
            chart_data
        )
        
        chart = chart_placeholder.chart
        
        # Styling for finance presentations
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.font.size = Pt(10)
        chart.legend.font.bold = True
        
        # Apply template colors to pie slices
        for i, point in enumerate(chart.series[0].points):
            color_idx = i % len(self.chart_colors)
            fill = point.format.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(*self.chart_colors[color_idx])
        
        # Data labels with percentages AND values (finance standard)
        chart.plots[0].has_data_labels = True
        data_labels = chart.plots[0].data_labels
        data_labels.show_percentage = True
        data_labels.show_value = True
        data_labels.show_category_name = False
        data_labels.font.size = Pt(10)
        data_labels.font.bold = True
        
        # Format: "Value (Percentage%)"
        data_labels.number_format = '#,##0 (0%)'
        
        return chart
    
    def create_bar_chart(self, slide, left, top, width, height, categories, values, title="Comparison", series_name="Values"):
        """
        Create horizontal bar chart for comparisons
        
        Args:
            slide: PowerPoint slide
            left, top, width, height: Position and size in inches
            categories: List of category names
            values: List of values
            title: Chart title
            series_name: Series name for legend
        """
        # Clean NaN/Inf values
        cleaned_values = []
        for val in values:
            try:
                # Convert to Python float first
                if pd.isna(val) or np.isnan(float(val)) or np.isinf(float(val)):
                    cleaned_values.append(float(0))
                else:
                    cleaned_values.append(float(val))
            except (ValueError, TypeError):
                cleaned_values.append(float(0))
        
        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series(series_name, cleaned_values)
        
        chart_placeholder = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_CLUSTERED,
            left, top, width, height,  # Already in EMU units, don't wrap in Inches()
            chart_data
        )
        
        chart = chart_placeholder.chart
        
        # Styling
        chart.has_legend = False
        
        # Apply template color
        fill = chart.series[0].format.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*self.template_colors['light_blue'])
        
        # Data labels
        chart.plots[0].has_data_labels = True
        data_labels = chart.plots[0].data_labels
        data_labels.show_value = True
        data_labels.font.size = Pt(10)
        
        return chart
    
    def create_line_chart(self, slide, left, top, width, height, categories, series_dict, title="Trend Analysis"):
        """
        Create line chart for trend/time-series data
        
        Args:
            slide: PowerPoint slide
            left, top, width, height: Position and size in inches
            categories: List of time periods (x-axis)
            series_dict: Dictionary {series_name: [values]}
            title: Chart title
        """
        chart_data = CategoryChartData()
        chart_data.categories = categories
        
        # Clean NaN/Inf values from series data
        for series_name, values in series_dict.items():
            # Replace NaN and Inf with 0, convert to Python float
            cleaned_values = []
            for val in values:
                try:
                    # Convert to Python float first
                    if pd.isna(val) or np.isnan(float(val)) or np.isinf(float(val)):
                        cleaned_values.append(float(0))
                    else:
                        cleaned_values.append(float(val))
                except (ValueError, TypeError):
                    cleaned_values.append(float(0))
            
            chart_data.add_series(series_name, cleaned_values)
        
        print(f"   🎯 Chart position: left={left}, top={top}, width={width}, height={height}")
        
        chart_placeholder = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE_MARKERS,
            left, top, width, height,  # Already in EMU units, don't wrap in Inches()
            chart_data
        )
        
        chart = chart_placeholder.chart
        
        # Styling
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(10)
        
        # Apply template colors to lines
        for i, series in enumerate(chart.series):
            color_idx = i % len(self.chart_colors)
            line = series.format.line
            line.color.rgb = RGBColor(*self.chart_colors[color_idx])
            line.width = Pt(2.5)
            
            # Marker style
            series.marker.style = 8  # Circle
            series.marker.size = 7
        
        return chart
    
    def create_column_chart(self, slide, left, top, width, height, categories, series_dict, title="Analysis"):
        """
        Create column chart for categorical comparisons (FINANCE OPTIMIZED)
        
        Args:
            slide: PowerPoint slide
            left, top, width, height: Position and size in inches
            categories: List of categories (x-axis)
            series_dict: Dictionary {series_name: [values]}
            title: Chart title
        """
        chart_data = CategoryChartData()
        chart_data.categories = categories
        
        # Clean NaN/Inf values from series data
        for series_name, values in series_dict.items():
            cleaned_values = []
            for val in values:
                try:
                    # Convert to Python float first
                    if pd.isna(val) or np.isnan(float(val)) or np.isinf(float(val)):
                        cleaned_values.append(float(0))
                    else:
                        cleaned_values.append(float(val))
                except (ValueError, TypeError):
                    cleaned_values.append(float(0))
            
            chart_data.add_series(series_name, cleaned_values)
        
        chart_placeholder = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            left, top, width, height,  # Already in EMU units, don't wrap in Inches()
            chart_data
        )
        
        chart = chart_placeholder.chart
        
        # Styling
        chart.has_legend = len(series_dict) > 1
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.font.size = Pt(10)
        
        # Apply template colors to columns
        for i, series in enumerate(chart.series):
            color_idx = i % len(self.chart_colors)
            for point in series.points:
                fill = point.format.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(*self.chart_colors[color_idx])
        
        # Add data labels for better readability in finance presentations
        for series in chart.series:
            series.has_data_labels = True
            data_labels = series.data_labels
            data_labels.show_value = True
            data_labels.show_category_name = False
            data_labels.position = 2  # Outside end
            data_labels.font.size = Pt(9)
            data_labels.font.bold = True
        
        return chart
        
        return chart
    
    def auto_create_chart(self, slide, df, left=1, top=2, width=8, height=4, title=None):
        """
        Automatically detect and create best chart type for DataFrame
        
        Args:
            slide: PowerPoint slide
            df: Pandas DataFrame
            left, top, width, height: Chart position and size
            title: Optional chart title
        
        Returns:
            chart object or None if failed
        """
        try:
            # Validate DataFrame
            if df is None or df.empty:
                print("   ⚠️ DataFrame is empty, cannot create chart")
                return None
            
            # Clean DataFrame: Replace NaN and Inf values
            df = df.replace([np.inf, -np.inf], np.nan)  # Convert Inf to NaN first
            df = df.fillna(0)  # Replace all NaN with 0
            
            text_cols = df.select_dtypes(include=['object']).columns
            numeric_cols = df.select_dtypes(include=['number']).columns
            
            if len(numeric_cols) == 0:
                print("   ⚠️ No numeric columns found, cannot create chart")
                return None
            
            chart_type = self.detect_chart_type(df)
            print(f"   📊 Detected chart type: {chart_type.upper()}")
            
            if chart_type == 'pie' and len(text_cols) > 0:
                # For pie chart, use first text column as categories, first numeric as values
                text_col = text_cols[0]
                numeric_col = numeric_cols[0]
                
                # Limit to top 6 categories
                top_data = df.nlargest(6, numeric_col) if len(df) > 6 else df
                data_dict = dict(zip(top_data[text_col], top_data[numeric_col]))
                
                print(f"   ✓ Creating PIE chart with {len(data_dict)} slices")
                return self.create_pie_chart(slide, left, top, width, height, data_dict, title or "Distribution")
            
            elif chart_type == 'bar' and len(text_cols) > 0:
                text_col = text_cols[0]
                numeric_col = numeric_cols[0]
                
                # Limit to top 8 items
                top_data = df.nlargest(8, numeric_col) if len(df) > 8 else df
                categories = list(top_data[text_col])
                values = list(top_data[numeric_col])
                
                print(f"   ✓ Creating BAR chart with {len(categories)} bars")
                return self.create_bar_chart(slide, left, top, width, height, categories, values, title or "Comparison")
            
            elif chart_type == 'line':
                # Time column + one or more numeric columns
                time_cols = [col for col in df.columns if any(kw in str(col).lower() 
                             for kw in ['date', 'time', 'quarter', 'month', 'year', 'q1', 'q2', 'q3', 'q4'])]
                
                if time_cols:
                    time_col = time_cols[0]
                    categories = list(df[time_col].astype(str))
                else:
                    # Use row numbers or first text column
                    if len(text_cols) > 0:
                        categories = list(df[text_cols[0]].astype(str))
                    else:
                        categories = [f"Period {i+1}" for i in range(len(df))]
                
                # Limit data points for clarity
                limit = min(20, len(df))
                categories = categories[:limit]
                
                # Get up to 3 numeric series
                numeric_cols_limited = numeric_cols[:3]
                series_dict = {col: list(df[col][:limit]) for col in numeric_cols_limited}
                
                print(f"   ✓ Creating LINE chart with {len(series_dict)} series and {limit} points")
                return self.create_line_chart(slide, left, top, width, height, categories, series_dict, title or "Trend")
            
            else:  # column chart (default)
                # For column chart, need categories
                if len(text_cols) > 0:
                    text_col = text_cols[0]
                    categories = list(df[text_col][:10])  # Top 10
                else:
                    # Use row index if no text columns
                    categories = [f"Item {i+1}" for i in range(min(10, len(df)))]
                
                # Get up to 2 numeric series
                numeric_cols_limited = numeric_cols[:2]
                series_dict = {col: list(df[col][:len(categories)]) for col in numeric_cols_limited}
                
                print(f"   ✓ Creating COLUMN chart with {len(series_dict)} series and {len(categories)} categories")
                return self.create_column_chart(slide, left, top, width, height, categories, series_dict, title or "Analysis")
        
        except Exception as e:
            print(f"   ❌ Error in auto_create_chart: {e}")
            import traceback
            traceback.print_exc()
            return None
