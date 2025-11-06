"""
Advanced Chart Templates with AI Integration and Multiple Fallbacks
Provides 10+ chart types with intelligent selection and clean templates
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional, Tuple
from pptx.util import Inches, Pt
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor


# ============================================================================
# CHART TYPE DEFINITIONS
# ============================================================================

CHART_TYPES = {
    # Basic Charts
    'column': {
        'ppt_type': XL_CHART_TYPE.COLUMN_CLUSTERED,
        'use_case': 'Comparing values across categories',
        'data_structure': 'categorical_comparison',
        'max_categories': 15
    },
    'bar': {
        'ppt_type': XL_CHART_TYPE.BAR_CLUSTERED,
        'use_case': 'Horizontal comparison, good for long labels',
        'data_structure': 'categorical_comparison',
        'max_categories': 15
    },
    'line': {
        'ppt_type': XL_CHART_TYPE.LINE,
        'use_case': 'Time series, trends over time',
        'data_structure': 'time_series',
        'max_points': 50
    },
    'line_markers': {
        'ppt_type': XL_CHART_TYPE.LINE_MARKERS,
        'use_case': 'Time series with data points highlighted',
        'data_structure': 'time_series',
        'max_points': 30
    },
    'pie': {
        'ppt_type': XL_CHART_TYPE.PIE,
        'use_case': 'Part-to-whole relationships',
        'data_structure': 'distribution',
        'max_categories': 8
    },
    'doughnut': {
        'ppt_type': XL_CHART_TYPE.DOUGHNUT,
        'use_case': 'Part-to-whole with modern look',
        'data_structure': 'distribution',
        'max_categories': 8
    },
    'area': {
        'ppt_type': XL_CHART_TYPE.AREA,
        'use_case': 'Cumulative trends over time',
        'data_structure': 'time_series',
        'max_points': 40
    },
    'area_stacked': {
        'ppt_type': XL_CHART_TYPE.AREA_STACKED,
        'use_case': 'Multiple series cumulative trends',
        'data_structure': 'multi_series_time',
        'max_series': 4
    },
    'column_stacked': {
        'ppt_type': XL_CHART_TYPE.COLUMN_STACKED,
        'use_case': 'Multiple series comparison',
        'data_structure': 'multi_series_categorical',
        'max_series': 4
    },
    'column_stacked_100': {
        'ppt_type': XL_CHART_TYPE.COLUMN_STACKED_100,
        'use_case': 'Percentage distribution across categories',
        'data_structure': 'multi_series_categorical',
        'max_series': 4
    },
    'scatter': {
        'ppt_type': XL_CHART_TYPE.XY_SCATTER,
        'use_case': 'Correlation between two variables',
        'data_structure': 'correlation',
        'max_points': 100
    },
    'scatter_lines': {
        'ppt_type': XL_CHART_TYPE.XY_SCATTER_LINES,
        'use_case': 'Correlation with trend lines',
        'data_structure': 'correlation',
        'max_points': 100
    },
    'bubble': {
        'ppt_type': XL_CHART_TYPE.BUBBLE,
        'use_case': '3-dimensional comparison',
        'data_structure': 'multi_dimensional',
        'max_points': 50
    }
}


# Professional Color Schemes
COLOR_SCHEMES = {
    'business': [
        (41, 128, 185),   # Blue
        (39, 174, 96),    # Green
        (230, 126, 34),   # Orange
        (231, 76, 60),    # Red
        (142, 68, 173),   # Purple
        (241, 196, 15),   # Yellow
        (52, 152, 219),   # Light Blue
        (26, 188, 156),   # Turquoise
    ],
    'vibrant': [
        (255, 107, 107),  # Coral
        (78, 205, 196),   # Mint
        (255, 195, 18),   # Gold
        (106, 137, 204),  # Periwinkle
        (255, 121, 198),  # Pink
        (116, 185, 255),  # Sky Blue
        (162, 155, 254),  # Lavender
        (223, 230, 233),  # Silver
    ],
    'professional': [
        (25, 42, 86),     # Navy
        (41, 128, 185),   # Blue
        (52, 152, 219),   # Light Blue
        (127, 140, 141),  # Gray
        (39, 174, 96),    # Green
        (230, 126, 34),   # Orange
    ],
    'gradient_blue': [
        (13, 71, 161),    # Dark Blue
        (25, 118, 210),   # Blue
        (66, 165, 245),   # Light Blue
        (144, 202, 249),  # Very Light Blue
    ]
}


class AdvancedChartBuilder:
    """
    Advanced chart builder with AI integration and multiple fallback options
    
    Priority System:
    1. AI Service recommendations (with confidence score)
    2. SmartChartAnalyzer intelligent selection
    3. Data structure analysis
    4. Default chart templates
    """
    
    def __init__(self, ai_service=None, smart_analyzer=None):
        self.ai_service = ai_service
        self.smart_analyzer = smart_analyzer
        self.color_scheme = COLOR_SCHEMES['business']
    
    # ========================================================================
    # MAIN CHART SELECTION METHOD
    # ========================================================================
    
    def select_chart_type(self, data: pd.DataFrame, context: str = 'general') -> Dict[str, Any]:
        """
        Intelligently select chart type using priority system
        
        Args:
            data: DataFrame to visualize
            context: Context hint ('dashboard', 'detailed', 'comparison', 'trend')
        
        Returns:
            Chart configuration dictionary
        """
        if data.empty:
            return self._get_default_config()
        
        # PRIORITY 1: AI Service Recommendations
        if self.ai_service:
            ai_config = self._get_ai_recommendation(data, context)
            if ai_config and ai_config.get('confidence', 0) > 0.7:
                print(f"🤖 AI recommends: {ai_config['type']} (confidence: {ai_config['confidence']:.2f})")
                return ai_config
        
        # PRIORITY 2: SmartChartAnalyzer
        if self.smart_analyzer:
            smart_config = self._get_smart_analyzer_recommendation(data, context)
            if smart_config:
                print(f"📊 SmartChartAnalyzer recommends: {smart_config['type']}")
                return smart_config
        
        # PRIORITY 3: Data Structure Analysis
        data_config = self._analyze_data_structure(data, context)
        if data_config:
            print(f"📈 Data structure analysis: {data_config['type']}")
            return data_config
        
        # PRIORITY 4: Default fallback
        print("⚙️  Using default chart configuration")
        return self._get_default_config()
    
    # ========================================================================
    # PRIORITY 1: AI SERVICE INTEGRATION
    # ========================================================================
    
    def _get_ai_recommendation(self, data: pd.DataFrame, context: str) -> Optional[Dict[str, Any]]:
        """Get chart recommendation from AI service"""
        try:
            if not self.ai_service:
                return None
            
            # Get AI recommendation
            ai_result = self.ai_service.recommend_chart_type(
                data, 
                data.columns.tolist(),
                business_context=context
            )
            
            if not ai_result or 'recommended' not in ai_result:
                return None
            
            recommended = ai_result['recommended']
            chart_type = recommended.get('type', '').lower()
            confidence = recommended.get('confidence', 0)
            reasoning = recommended.get('reasoning', '')
            
            # Map AI recommendations to our chart types
            chart_mapping = {
                'bar': 'bar',
                'column': 'column',
                'line': 'line_markers',
                'pie': 'pie',
                'scatter': 'scatter',
                'area': 'area',
                'stacked_bar': 'column_stacked',
                'stacked_column': 'column_stacked',
                'horizontal_bar': 'bar',
                'trend': 'line_markers',
                'distribution': 'pie',
                'comparison': 'column'
            }
            
            mapped_type = chart_mapping.get(chart_type, 'column')
            
            return {
                'type': mapped_type,
                'confidence': confidence,
                'reasoning': reasoning,
                'source': 'ai_service',
                'ppt_type': CHART_TYPES[mapped_type]['ppt_type']
            }
            
        except Exception as e:
            print(f"⚠️  AI recommendation failed: {e}")
            return None
    
    # ========================================================================
    # PRIORITY 2: SMART CHART ANALYZER
    # ========================================================================
    
    def _get_smart_analyzer_recommendation(self, data: pd.DataFrame, context: str) -> Optional[Dict[str, Any]]:
        """Get recommendation from SmartChartAnalyzer"""
        try:
            if not self.smart_analyzer:
                return None
            
            # Get recommendations from SmartChartAnalyzer
            recommendations = self.smart_analyzer.get_recommended_charts(context)
            
            if not recommendations or len(recommendations) == 0:
                return None
            
            # Use first recommendation
            chart_config = recommendations[0]
            chart_type = chart_config.get('type', 'column')
            
            # Map SmartChartAnalyzer types to our types
            type_mapping = {
                'bar': 'bar',
                'column': 'column',
                'line': 'line_markers',
                'pie': 'pie',
                'scatter': 'scatter',
                'area': 'area'
            }
            
            mapped_type = type_mapping.get(chart_type, 'column')
            
            return {
                'type': mapped_type,
                'source': 'smart_analyzer',
                'original_config': chart_config,
                'ppt_type': CHART_TYPES[mapped_type]['ppt_type']
            }
            
        except Exception as e:
            print(f"⚠️  SmartChartAnalyzer failed: {e}")
            return None
    
    # ========================================================================
    # PRIORITY 3: DATA STRUCTURE ANALYSIS
    # ========================================================================
    
    def _analyze_data_structure(self, data: pd.DataFrame, context: str) -> Dict[str, Any]:
        """Analyze data structure to determine best chart type"""
        numeric_cols = data.select_dtypes(include=[np.number]).columns
        text_cols = data.select_dtypes(include=['object']).columns
        date_cols = data.select_dtypes(include=['datetime64']).columns
        
        n_rows = len(data)
        n_numeric = len(numeric_cols)
        n_text = len(text_cols)
        n_dates = len(date_cols)
        
        # Time series detection
        if n_dates > 0 or self._is_time_series_data(data):
            if n_numeric > 1:
                return {'type': 'area_stacked', 'ppt_type': XL_CHART_TYPE.AREA_STACKED}
            else:
                return {'type': 'line_markers', 'ppt_type': XL_CHART_TYPE.LINE_MARKERS}
        
        # Distribution analysis (few categories)
        if n_text > 0 and n_numeric > 0 and n_rows <= 8:
            return {'type': 'doughnut', 'ppt_type': XL_CHART_TYPE.DOUGHNUT}
        
        # Comparison (many categories)
        if n_text > 0 and n_numeric > 0 and n_rows > 8:
            return {'type': 'bar', 'ppt_type': XL_CHART_TYPE.BAR_CLUSTERED}
        
        # Multiple series comparison
        if n_numeric >= 2 and n_text > 0:
            if context == 'percentage' or 'percent' in str(data.columns).lower():
                return {'type': 'column_stacked_100', 'ppt_type': XL_CHART_TYPE.COLUMN_STACKED_100}
            else:
                return {'type': 'column_stacked', 'ppt_type': XL_CHART_TYPE.COLUMN_STACKED}
        
        # Correlation analysis (two numeric columns)
        if n_numeric >= 2 and n_rows > 10:
            return {'type': 'scatter_lines', 'ppt_type': XL_CHART_TYPE.XY_SCATTER_LINES}
        
        # Default: column chart
        return {'type': 'column', 'ppt_type': XL_CHART_TYPE.COLUMN_CLUSTERED}
    
    def _is_time_series_data(self, data: pd.DataFrame) -> bool:
        """Detect if data represents time series"""
        # Check if index is datetime
        if pd.api.types.is_datetime64_any_dtype(data.index):
            return True
        
        # Check for common time-related column names
        time_keywords = ['date', 'time', 'year', 'month', 'quarter', 'day', 'week']
        for col in data.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in time_keywords):
                return True
        
        return False
    
    # ========================================================================
    # PRIORITY 4: DEFAULT CONFIGURATION
    # ========================================================================
    
    def _get_default_config(self) -> Dict[str, Any]:
        """Get default chart configuration"""
        return {
            'type': 'column',
            'ppt_type': XL_CHART_TYPE.COLUMN_CLUSTERED,
            'source': 'default'
        }
    
    # ========================================================================
    # CHART CREATION METHODS
    # ========================================================================
    
    def create_chart(self, slide, data: pd.DataFrame, chart_config: Dict[str, Any],
                    x: Inches, y: Inches, width: Inches, height: Inches,
                    title: str = None) -> Any:
        """
        Create chart on slide using configuration
        
        Args:
            slide: PowerPoint slide object
            data: Data to visualize
            chart_config: Chart configuration dictionary
            x, y, width, height: Chart position and size
            title: Optional chart title
        
        Returns:
            Chart object
        """
        chart_type = chart_config.get('type', 'column')
        ppt_type = chart_config.get('ppt_type', XL_CHART_TYPE.COLUMN_CLUSTERED)
        
        try:
            # Prepare data based on chart type
            if chart_type in ['scatter', 'scatter_lines', 'bubble']:
                chart_data = self._prepare_xy_data(data, chart_type)
            else:
                chart_data = self._prepare_category_data(data, chart_type)
            
            # Add chart to slide
            chart = slide.shapes.add_chart(
                ppt_type,
                x, y, width, height,
                chart_data
            ).chart
            
            # Apply styling
            self._apply_chart_styling(chart, chart_type, title)
            
            return chart
            
        except Exception as e:
            print(f"⚠️  Chart creation failed: {e}")
            # Fallback to simple column chart
            return self._create_fallback_chart(slide, data, x, y, width, height)
    
    def _prepare_category_data(self, data: pd.DataFrame, chart_type: str) -> CategoryChartData:
        """Prepare data for category-based charts"""
        chart_data = CategoryChartData()
        
        numeric_cols = data.select_dtypes(include=[np.number]).columns
        text_cols = data.select_dtypes(include=['object']).columns
        
        # Determine categories
        if len(text_cols) > 0:
            categories = data[text_cols[0]].head(15).tolist()
            data_subset = data.head(15)
        else:
            categories = [f"Item {i+1}" for i in range(min(15, len(data)))]
            data_subset = data.head(15)
        
        chart_data.categories = categories
        
        # Add series based on chart type
        if chart_type in ['column_stacked', 'column_stacked_100', 'area_stacked']:
            # Multiple series
            for col in numeric_cols[:3]:  # Max 3 series
                values = data_subset[col].fillna(0).tolist()
                chart_data.add_series(col, values)
        else:
            # Single series
            if len(numeric_cols) > 0:
                col = numeric_cols[0]
                values = data_subset[col].fillna(0).tolist()
                chart_data.add_series(col, values)
        
        return chart_data
    
    def _prepare_xy_data(self, data: pd.DataFrame, chart_type: str) -> XyChartData:
        """Prepare data for XY scatter/bubble charts"""
        chart_data = XyChartData()
        
        numeric_cols = data.select_dtypes(include=[np.number]).columns
        
        if len(numeric_cols) >= 2:
            x_col = numeric_cols[0]
            y_col = numeric_cols[1]
            
            # Clean data
            clean_data = data[[x_col, y_col]].dropna().head(50)
            
            series = chart_data.add_series('Data Points')
            for _, row in clean_data.iterrows():
                series.add_data_point(float(row[x_col]), float(row[y_col]))
        
        return chart_data
    
    def _apply_chart_styling(self, chart, chart_type: str, title: str = None):
        """Apply professional styling to chart"""
        # Title
        if title:
            chart.has_title = True
            chart.chart_title.text_frame.text = title
            chart.chart_title.text_frame.paragraphs[0].font.size = Pt(18)
            chart.chart_title.text_frame.paragraphs[0].font.bold = True
        
        # Legend
        if chart_type not in ['pie', 'doughnut']:
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.font.size = Pt(10)
        else:
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.font.size = Pt(10)
        
        # Data labels for pie/doughnut
        if chart_type in ['pie', 'doughnut']:
            plot = chart.plots[0]
            plot.has_data_labels = True
            data_labels = plot.data_labels
            data_labels.font.size = Pt(10)
            data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    
    def _create_fallback_chart(self, slide, data: pd.DataFrame, 
                               x: Inches, y: Inches, width: Inches, height: Inches):
        """Create simple fallback chart if main creation fails"""
        chart_data = CategoryChartData()
        
        numeric_cols = data.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            chart_data.categories = [f"Item {i+1}" for i in range(min(10, len(data)))]
            values = data[numeric_cols[0]].head(10).fillna(0).tolist()
            chart_data.add_series('Values', values)
        else:
            chart_data.categories = ['No Data']
            chart_data.add_series('Values', [0])
        
        return slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            x, y, width, height,
            chart_data
        ).chart


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def format_large_number(value: float) -> str:
    """Format large numbers with K, M, B suffixes"""
    if abs(value) >= 1e9:
        return f"{value/1e9:.1f}B"
    elif abs(value) >= 1e6:
        return f"{value/1e6:.1f}M"
    elif abs(value) >= 1e3:
        return f"{value/1e3:.1f}K"
    else:
        return f"{value:.1f}"


def get_trend_arrow(current: float, previous: float) -> str:
    """Get trend arrow emoji based on change"""
    if current > previous:
        return "↗️"
    elif current < previous:
        return "↘️"
    else:
        return "→"
