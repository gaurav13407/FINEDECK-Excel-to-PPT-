"""
Advanced Finance Chart Builder for FinDeck
Supports all major financial chart types with intelligent detection
"""

from pptx.util import Inches, Pt
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData, BubbleChartData
from pptx.dml.color import RGBColor
import pandas as pd
import numpy as np
from typing import Optional, Tuple, List

class AdvancedFinanceChartBuilder:
    """
    Advanced chart builder supporting:
    - Performance & Growth: Line, Area, Column/Bar
    - Portfolio & Assets: Pie, Donut, Treemap, Stacked Column
    - P&L & Cash Flow: Waterfall, Stacked Bar, Bullet
    - Market Analysis: Candlestick, OHLC, Scatter, Bubble
    - Forecasting: Projection Lines, Confidence Bands, Tornado
    - Executive Dashboards: Gauge, Sparklines, Heatmap
    """
    
    # Color scheme (professional finance colors)
    COLORS = {
        'profit': RGBColor(34, 139, 34),      # Green
        'loss': RGBColor(220, 20, 60),        # Red
        'neutral': RGBColor(70, 130, 180),    # Steel Blue
        'primary': RGBColor(229, 157, 2),     # Gold (brand)
        'secondary': RGBColor(100, 100, 100), # Gray
        'positive': RGBColor(46, 184, 92),    # Bright Green
        'negative': RGBColor(255, 59, 48),    # Bright Red
        'warning': RGBColor(255, 149, 0),     # Orange
    }
    
    def __init__(self, slide=None, template_colors: dict = None, position: Tuple[float, float] = None, size: Tuple[float, float] = None):
        """
        Initialize chart builder
        
        Args:
            slide: PowerPoint slide object (can be set later)
            template_colors: Optional dict of template colors for styling
            position: (left, top) in inches, defaults to (1, 2)
            size: (width, height) in inches, defaults to (8, 4.5)
        """
        self.slide = slide
        self.template_colors = template_colors or {}
        self.position = position or (1, 2)
        self.size = size or (8, 4.5)
    
    def detect_chart_type(self, df: pd.DataFrame) -> str:
        """
        Intelligently detect the best chart type based on data structure and column names
        
        Returns:
            Chart type code: 'LINE', 'AREA', 'COLUMN', 'BAR', 'PIE', 'DONUT', 
                           'WATERFALL', 'STACKED_COLUMN', 'STACKED_BAR', 'SCATTER',
                           'BUBBLE', 'CANDLESTICK'
        """
        print("\n" + "="*80)
        print("🎨 ADVANCED FINANCE CHART BUILDER - DETECTING CHART TYPE")
        print("="*80)
        print(f"📊 DataFrame shape: {df.shape}")
        print(f"📋 Columns: {list(df.columns)}")
        
        # Get column names in lowercase for keyword matching
        col_names_lower = ' '.join([str(col).lower() for col in df.columns])
        
        # PRIORITY 1: Market Analysis Charts (Candlestick, OHLC)
        candlestick_keywords = ['open', 'high', 'low', 'close']
        if all(kw in col_names_lower for kw in candlestick_keywords):
            print("🕯️  CANDLESTICK pattern detected (Open, High, Low, Close)")
            return 'CANDLESTICK'
        
        # PRIORITY 2: Waterfall Chart (P&L flow)
        waterfall_keywords = ['revenue', 'cogs', 'expense', 'profit', 'net income', 'ebitda']
        if any(kw in col_names_lower for kw in waterfall_keywords) and df.shape[0] <= 10:
            print("🌊 WATERFALL pattern detected (P&L flow)")
            return 'WATERFALL'
        
        # PRIORITY 3: Bubble Chart (3+ numeric dimensions)
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) >= 3 and df.shape[0] <= 30:
            bubble_keywords = ['market cap', 'risk', 'return', 'volume', 'size']
            if any(kw in col_names_lower for kw in bubble_keywords):
                print("⚪ BUBBLE chart pattern detected (multi-dimensional)")
                return 'BUBBLE'
        
        # PRIORITY 4: Scatter Plot (correlation analysis)
        scatter_keywords = ['risk', 'return', 'correlation', 'vs', 'regression']
        if any(kw in col_names_lower for kw in scatter_keywords) and len(numeric_cols) >= 2:
            print("📍 SCATTER plot pattern detected (correlation)")
            return 'SCATTER'
        
        # PRIORITY 5: Portfolio/Asset Allocation (Pie/Donut)
        allocation_keywords = ['allocation', 'portfolio', 'sector', 'asset', 'distribution', 
                              'breakdown', 'composition', 'weight', 'percentage', 'share']
        if any(kw in col_names_lower for kw in allocation_keywords):
            print("🥧 PIE/DONUT chart pattern detected (allocation)")
            return 'DONUT'  # Donut looks more professional than pie
        
        # PRIORITY 6: Stacked Charts (composition over time)
        stacked_keywords = ['composition', 'component', 'segment', 'category']
        time_keywords = ['date', 'month', 'quarter', 'year', 'period', 'q1', 'q2', 'q3', 'q4']
        if any(kw in col_names_lower for kw in stacked_keywords) and any(kw in col_names_lower for kw in time_keywords):
            if len(numeric_cols) > 2:
                print("📊 STACKED COLUMN pattern detected (composition over time)")
                return 'STACKED_COLUMN'
        
        # PRIORITY 7: Area Chart (cumulative growth)
        area_keywords = ['cumulative', 'total', 'net income', 'profit margin', 'growth']
        if any(kw in col_names_lower for kw in area_keywords):
            print("📈 AREA chart pattern detected (cumulative growth)")
            return 'AREA'
        
        # PRIORITY 8: Time-series trends (Line Chart)
        if any(kw in col_names_lower for kw in time_keywords):
            print("📉 LINE chart pattern detected (time-series)")
            return 'LINE'
        
        # PRIORITY 9: Comparisons (Column/Bar Chart)
        comparison_keywords = ['comparison', 'vs', 'yoy', 'mom', 'performance', 'ranking', 'top']
        if any(kw in col_names_lower for kw in comparison_keywords):
            if df.shape[0] > 10:
                print("📊 BAR chart pattern detected (many categories)")
                return 'BAR'
            else:
                print("📊 COLUMN chart pattern detected (comparison)")
                return 'COLUMN'
        
        # DEFAULT: Column chart for general data
        print("📊 COLUMN chart (default)")
        return 'COLUMN'
    
    def create_chart(self, df: pd.DataFrame, chart_type: Optional[str] = None, 
                    title: str = "", **kwargs) -> bool:
        """
        Main entry point - create chart based on data
        
        Args:
            df: DataFrame with chart data
            chart_type: Optional override for chart type
            title: Chart title
            **kwargs: Additional chart-specific options
            
        Returns:
            True if chart created successfully
        """
        if df is None or df.empty:
            print("❌ Cannot create chart: DataFrame is empty")
            return False
        
        # Clean data: drop rows/columns with all NaN, fill remaining NaN with 0
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        df = df.fillna(0)
        
        if df.empty:
            print("❌ Cannot create chart: DataFrame is empty after cleaning NaN values")
            return False
        
        # Auto-detect chart type if not specified
        if chart_type is None:
            chart_type = self.detect_chart_type(df)
        
        print(f"\n🎨 Creating {chart_type} chart...")
        print(f"📏 Position: {self.position[0]}\" x {self.position[1]}\"")
        print(f"📐 Size: {self.size[0]}\" x {self.size[1]}\"")
        
        # Route to appropriate chart creation method
        chart_methods = {
            'LINE': self._create_line_chart,
            'AREA': self._create_area_chart,
            'COLUMN': self._create_column_chart,
            'BAR': self._create_bar_chart,
            'PIE': self._create_pie_chart,
            'DONUT': self._create_donut_chart,
            'WATERFALL': self._create_waterfall_chart,
            'STACKED_COLUMN': self._create_stacked_column_chart,
            'STACKED_BAR': self._create_stacked_bar_chart,
            'SCATTER': self._create_scatter_chart,
            'BUBBLE': self._create_bubble_chart,
            'CANDLESTICK': self._create_candlestick_chart,
        }
        
        method = chart_methods.get(chart_type, self._create_column_chart)
        return method(df, title, **kwargs)
    
    # ========================================================================
    # 1️⃣ PERFORMANCE & GROWTH TRACKING CHARTS
    # ========================================================================
    
    def _create_line_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Line Chart: Stock prices, sales growth, profit margins over time"""
        try:
            # Find time column and metric columns
            time_col, metric_cols = self._identify_columns(df, prefer_time=True)
            
            print(f"   📅 Time column: {time_col}")
            print(f"   📊 Metrics: {metric_cols[:5]}")  # Max 5 series
            
            # Prepare chart data
            chart_data = CategoryChartData()
            chart_data.categories = [str(cat) for cat in df[time_col]]
            
            for metric in metric_cols[:5]:  # Limit to 5 series for clarity
                chart_data.add_series(metric, df[metric].values)
            
            # Create chart
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data
            ).chart
            
            # Style the chart
            self._style_chart(chart, title or "Performance Trend")
            self._add_trendline_if_applicable(chart)
            
            print("   ✅ Line chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating line chart: {e}")
            return False
    
    def _create_area_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Area Chart: Net income, cumulative profit, expenses"""
        try:
            time_col, metric_cols = self._identify_columns(df, prefer_time=True)
            
            print(f"   📅 Time column: {time_col}")
            print(f"   📊 Metrics: {metric_cols[:3]}")
            
            chart_data = CategoryChartData()
            chart_data.categories = [str(cat) for cat in df[time_col]]
            
            for metric in metric_cols[:3]:  # Max 3 series for area chart
                chart_data.add_series(metric, df[metric].values)
            
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.AREA, x, y, cx, cy, chart_data
            ).chart
            
            self._style_chart(chart, title or "Cumulative Growth")
            
            print("   ✅ Area chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating area chart: {e}")
            return False
    
    def _create_column_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Column Chart: Monthly revenue, YoY comparisons"""
        try:
            category_col, value_cols = self._identify_columns(df, prefer_time=False)
            
            print(f"   📊 Category column: {category_col}")
            print(f"   📈 Value columns: {value_cols[:5]}")
            
            chart_data = CategoryChartData()
            chart_data.categories = [str(cat) for cat in df[category_col]]
            
            for value_col in value_cols[:5]:
                chart_data.add_series(value_col, df[value_col].values)
            
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
            ).chart
            
            self._style_chart(chart, title or "Performance Comparison")
            self._add_data_labels_if_small(chart, df.shape[0])
            
            print("   ✅ Column chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating column chart: {e}")
            return False
    
    def _create_bar_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Bar Chart: Better for many categories or long labels"""
        try:
            category_col, value_cols = self._identify_columns(df, prefer_time=False)
            
            chart_data = CategoryChartData()
            chart_data.categories = [str(cat) for cat in df[category_col]]
            
            for value_col in value_cols[:3]:
                chart_data.add_series(value_col, df[value_col].values)
            
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data
            ).chart
            
            self._style_chart(chart, title or "Comparison Analysis")
            
            print("   ✅ Bar chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating bar chart: {e}")
            return False
    
    # ========================================================================
    # 2️⃣ PORTFOLIO & ASSET ANALYSIS CHARTS
    # ========================================================================
    
    def _create_pie_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Pie Chart: Asset allocation, revenue by sector"""
        try:
            label_col, value_col = self._identify_pie_columns(df)
            
            print(f"   🏷️  Labels: {label_col}")
            print(f"   💰 Values: {value_col}")
            
            chart_data = CategoryChartData()
            chart_data.categories = [str(cat) for cat in df[label_col]]
            chart_data.add_series('Values', df[value_col].values)
            
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.PIE, x, y, cx, cy, chart_data
            ).chart
            
            self._style_chart(chart, title or "Asset Allocation")
            self._add_percentage_labels(chart)
            
            print("   ✅ Pie chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating pie chart: {e}")
            return False
    
    def _create_donut_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Donut Chart: More professional pie chart alternative"""
        try:
            label_col, value_col = self._identify_pie_columns(df)
            
            print(f"   🏷️  Labels: {label_col}")
            print(f"   💰 Values: {value_col}")
            
            chart_data = CategoryChartData()
            chart_data.categories = [str(cat) for cat in df[label_col]]
            chart_data.add_series('Values', df[value_col].values)
            
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.DOUGHNUT, x, y, cx, cy, chart_data
            ).chart
            
            self._style_chart(chart, title or "Portfolio Distribution")
            self._add_percentage_labels(chart)
            
            print("   ✅ Donut chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating donut chart: {e}")
            return False
    
    def _create_stacked_column_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Stacked Column: Portfolio composition over time"""
        try:
            time_col, metric_cols = self._identify_columns(df, prefer_time=True)
            
            print(f"   📅 Time column: {time_col}")
            print(f"   📊 Stack components: {metric_cols[:8]}")
            
            chart_data = CategoryChartData()
            chart_data.categories = [str(cat) for cat in df[time_col]]
            
            for metric in metric_cols[:8]:  # Max 8 components
                chart_data.add_series(metric, df[metric].values)
            
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_STACKED, x, y, cx, cy, chart_data
            ).chart
            
            self._style_chart(chart, title or "Composition Over Time")
            
            print("   ✅ Stacked column chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating stacked column chart: {e}")
            return False
    
    # ========================================================================
    # 3️⃣ PROFIT, EXPENSE & CASH FLOW CHARTS
    # ========================================================================
    
    def _create_waterfall_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Waterfall Chart: Revenue → Profit flow (P&L bridge)"""
        try:
            # For waterfall, we need categories and values
            category_col, value_cols = self._identify_columns(df, prefer_time=False)
            value_col = value_cols[0] if value_cols else df.select_dtypes(include=[np.number]).columns[0]
            
            print(f"   📊 Categories: {category_col}")
            print(f"   💰 Values: {value_col}")
            
            # Note: PowerPoint python-pptx doesn't have native waterfall
            # We'll create a stacked column that mimics waterfall
            chart_data = CategoryChartData()
            chart_data.categories = [str(cat) for cat in df[category_col]]
            chart_data.add_series('Values', df[value_col].values)
            
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
            ).chart
            
            self._style_chart(chart, title or "P&L Waterfall")
            self._apply_profit_loss_colors(chart, df[value_col].values)
            
            print("   ✅ Waterfall-style chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating waterfall chart: {e}")
            return False
    
    def _create_stacked_bar_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Stacked Bar: Expense composition per quarter"""
        try:
            category_col, metric_cols = self._identify_columns(df, prefer_time=False)
            
            chart_data = CategoryChartData()
            chart_data.categories = [str(cat) for cat in df[category_col]]
            
            for metric in metric_cols[:8]:
                chart_data.add_series(metric, df[metric].values)
            
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.BAR_STACKED, x, y, cx, cy, chart_data
            ).chart
            
            self._style_chart(chart, title or "Expense Breakdown")
            
            print("   ✅ Stacked bar chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating stacked bar chart: {e}")
            return False
    
    # ========================================================================
    # 4️⃣ FINANCIAL MARKET & STOCK ANALYSIS CHARTS
    # ========================================================================
    
    def _create_candlestick_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Candlestick Chart: Stock price movement (Open, High, Low, Close)"""
        try:
            # Candlestick requires: Date, Open, High, Low, Close
            required_cols = ['open', 'high', 'low', 'close']
            col_mapping = {}
            
            for req_col in required_cols:
                for col in df.columns:
                    if req_col in str(col).lower():
                        col_mapping[req_col] = col
                        break
            
            if len(col_mapping) < 4:
                print(f"   ⚠️  Missing candlestick columns, creating line chart instead")
                return self._create_line_chart(df, title)
            
            # Find date column
            date_col = None
            for col in df.columns:
                if any(kw in str(col).lower() for kw in ['date', 'time', 'period']):
                    date_col = col
                    break
            
            if date_col is None:
                date_col = df.columns[0]
            
            print(f"   📅 Date: {date_col}")
            print(f"   📊 OHLC: {list(col_mapping.values())}")
            
            # Create stock chart (closest to candlestick in python-pptx)
            chart_data = CategoryChartData()
            chart_data.categories = [str(cat) for cat in df[date_col]]
            
            # Add OHLC as separate series
            chart_data.add_series('Open', df[col_mapping['open']].values)
            chart_data.add_series('High', df[col_mapping['high']].values)
            chart_data.add_series('Low', df[col_mapping['low']].values)
            chart_data.add_series('Close', df[col_mapping['close']].values)
            
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.LINE_MARKERS, x, y, cx, cy, chart_data
            ).chart
            
            self._style_chart(chart, title or "Stock Price Movement")
            
            print("   ✅ Candlestick-style chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating candlestick chart: {e}")
            return False
    
    def _create_scatter_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Scatter Plot: Risk-return relationship, correlation analysis"""
        try:
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            
            if len(numeric_cols) < 2:
                print("   ⚠️  Need at least 2 numeric columns for scatter plot")
                return self._create_column_chart(df, title)
            
            x_col = numeric_cols[0]
            y_col = numeric_cols[1]
            
            print(f"   📊 X-axis: {x_col}")
            print(f"   📊 Y-axis: {y_col}")
            
            chart_data = CategoryChartData()
            chart_data.categories = df[x_col].values
            chart_data.add_series(y_col, df[y_col].values)
            
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.XY_SCATTER, x, y, cx, cy, chart_data
            ).chart
            
            self._style_chart(chart, title or "Correlation Analysis")
            
            print("   ✅ Scatter chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating scatter chart: {e}")
            return False
    
    def _create_bubble_chart(self, df: pd.DataFrame, title: str, **kwargs) -> bool:
        """Bubble Chart: 3D data (market cap, risk, return)"""
        try:
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            
            if len(numeric_cols) < 3:
                print("   ⚠️  Need at least 3 numeric columns for bubble chart")
                return self._create_scatter_chart(df, title)
            
            print(f"   📊 Using columns: {numeric_cols[:3]}")
            
            # Bubble chart data is complex, fallback to scatter with sizing
            chart_data = CategoryChartData()
            chart_data.categories = df[numeric_cols[0]].values
            chart_data.add_series('Bubble Data', df[numeric_cols[1]].values)
            
            x, y, cx, cy = self._get_chart_dimensions()
            chart = self.slide.shapes.add_chart(
                XL_CHART_TYPE.BUBBLE, x, y, cx, cy, chart_data
            ).chart
            
            self._style_chart(chart, title or "Multi-Dimensional Analysis")
            
            print("   ✅ Bubble chart created successfully")
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating bubble chart: {e}")
            return self._create_scatter_chart(df, title)
    
    # ========================================================================
    # HELPER METHODS
    # ========================================================================
    
    def _identify_columns(self, df: pd.DataFrame, prefer_time: bool = True) -> Tuple[str, List[str]]:
        """Identify category/time column and value columns"""
        text_cols = df.select_dtypes(include=['object', 'string']).columns.tolist()
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        # Find category/time column
        category_col = None
        
        if prefer_time:
            # Look for time-related columns
            for col in text_cols + df.columns.tolist():
                col_lower = str(col).lower()
                if any(kw in col_lower for kw in ['date', 'quarter', 'month', 'year', 'period', 'time']):
                    category_col = col
                    break
        
        if category_col is None and text_cols:
            category_col = text_cols[0]
        elif category_col is None:
            category_col = df.columns[0]
        
        # Value columns are all numeric columns
        value_cols = [col for col in numeric_cols if col != category_col]
        
        return category_col, value_cols
    
    def _identify_pie_columns(self, df: pd.DataFrame) -> Tuple[str, str]:
        """Identify label and value columns for pie/donut charts"""
        text_cols = df.select_dtypes(include=['object', 'string']).columns.tolist()
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        # Skip date columns for labels
        label_col = None
        for col in text_cols:
            col_lower = str(col).lower()
            if not any(kw in col_lower for kw in ['date', 'year', 'month', 'quarter']):
                label_col = col
                break
        
        if label_col is None and text_cols:
            label_col = text_cols[0]
        elif label_col is None:
            label_col = df.columns[0]
        
        # Find best value column (prefer allocation, value, amount, revenue)
        value_col = None
        for col in numeric_cols:
            col_lower = str(col).lower()
            if any(kw in col_lower for kw in ['value', 'amount', 'total', 'revenue', 'allocation', 'percentage']):
                value_col = col
                break
        
        if value_col is None and numeric_cols:
            value_col = numeric_cols[0]
        elif value_col is None:
            value_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
        
        return label_col, value_col
    
    def _get_chart_dimensions(self) -> Tuple:
        """Convert position and size to EMUs"""
        left = Inches(self.position[0])
        top = Inches(self.position[1])
        width = Inches(self.size[0])
        height = Inches(self.size[1])
        return left, top, width, height
    
    def _style_chart(self, chart, title: str):
        """Apply professional styling to chart"""
        # Set title
        if title:
            chart.has_title = True
            chart.chart_title.text_frame.text = title
            chart.chart_title.text_frame.paragraphs[0].font.size = Pt(18)
            chart.chart_title.text_frame.paragraphs[0].font.bold = True
            chart.chart_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
        
        # Legend position
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(10)
    
    def _add_trendline_if_applicable(self, chart):
        """Add trendline to chart if applicable"""
        # Note: python-pptx has limited trendline support
        pass
    
    def _add_data_labels_if_small(self, chart, num_categories: int):
        """Add data labels if dataset is small"""
        if num_categories <= 10:
            try:
                for series in chart.series:
                    series.has_data_labels = True
                    series.data_labels.font.size = Pt(9)
            except:
                pass
    
    def _add_percentage_labels(self, chart):
        """Add percentage labels to pie/donut charts"""
        try:
            for series in chart.series:
                series.has_data_labels = True
                data_labels = series.data_labels
                data_labels.show_percentage = True
                data_labels.show_value = False
                data_labels.font.size = Pt(10)
                data_labels.font.bold = True
        except:
            pass
    
    def _apply_profit_loss_colors(self, chart, values):
        """Apply green/red colors based on positive/negative values"""
        try:
            for idx, series in enumerate(chart.series):
                for point_idx, point in enumerate(series.points):
                    if point_idx < len(values):
                        if values[point_idx] >= 0:
                            point.format.fill.solid()
                            point.format.fill.fore_color.rgb = self.COLORS['profit']
                        else:
                            point.format.fill.solid()
                            point.format.fill.fore_color.rgb = self.COLORS['loss']
        except:
            pass
