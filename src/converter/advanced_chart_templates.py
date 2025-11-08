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
        Intelligently select chart type using priority system.
        
        Supports a finance-only mode when `context` contains 'finance'. In that
        case the selected chart type will be restricted to finance-appropriate
        charts (column, bar, line_markers, doughnut, column_stacked_100, scatter).
        
        Args:
            data: DataFrame to visualize
            context: Context hint ('dashboard', 'detailed', 'comparison', 'trend', 'finance')
        
        Returns:
            Chart configuration dictionary
        """
        if data.empty:
            return self._get_default_config()

        result_config = None

        # PRIORITY 1: AI Service Recommendations
        if self.ai_service:
            ai_config = self._get_ai_recommendation(data, context)
            if ai_config and ai_config.get('confidence', 0) > 0.7:
                print(f"🤖 AI recommends: {ai_config['type']} (confidence: {ai_config['confidence']:.2f})")
                result_config = ai_config

        # PRIORITY 2: SmartChartAnalyzer (only if no strong AI recommendation)
        if result_config is None and self.smart_analyzer:
            smart_config = self._get_smart_analyzer_recommendation(data, context)
            if smart_config:
                print(f"📊 SmartChartAnalyzer recommends: {smart_config['type']}")
                result_config = smart_config

        # PRIORITY 3: Data Structure Analysis
        if result_config is None:
            data_config = self._analyze_data_structure(data, context)
            if data_config:
                print(f"📈 Data structure analysis: {data_config['type']}")
                result_config = data_config

        # PRIORITY 4: Default fallback
        if result_config is None:
            print("⚙️  Using default chart configuration")
            result_config = self._get_default_config()

        # If context asks for finance-only charts, filter the chosen type
        if isinstance(context, str) and 'finance' in context.lower():
            result_config = self._filter_for_finance(result_config)

        return result_config

    def _filter_for_finance(self, chart_config: Dict[str, Any]) -> Dict[str, Any]:
        """Restrict chart types to finance-appropriate ones.

        Allowed finance chart types: column, bar, line_markers, doughnut,
        column_stacked_100, scatter, bubble
        """
        allowed = {'column', 'bar', 'line_markers', 'doughnut', 'column_stacked_100', 'scatter', 'bubble'}
        ctype = chart_config.get('type', 'column')
        if ctype in allowed:
            return chart_config

        # Map unsupported types to nearest finance-appropriate type
        fallback_map = {
            'area': 'line_markers',
            'area_stacked': 'line_markers',
            'column_stacked': 'column',
            'column_stacked_100': 'column_stacked_100',
            'pie': 'doughnut',
            'bubble': 'bubble',
            'scatter_lines': 'scatter',
            'scatter': 'scatter'
        }

        new_type = fallback_map.get(ctype, 'column')
        # Ensure ppt_type is updated
        ppt_type = CHART_TYPES.get(new_type, CHART_TYPES['column'])['ppt_type']
        print(f"   🔒 Finance mode: remapping '{ctype}' → '{new_type}'")
        return {'type': new_type, 'ppt_type': ppt_type, 'source': chart_config.get('source', 'filtered')}
    
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
            
            # Map AI recommendations to ONLY 4 PRACTICAL CHART TYPES
            # Column, Bar, Doughnut, Line - that's it!
            chart_mapping = {
                # Direct mappings
                'bar': 'bar',
                'column': 'column',
                'line': 'line_markers',
                'pie': 'doughnut',
                'doughnut': 'doughnut',
                
                # Everything else → simplify to basic types
                'scatter': 'column',          # Scatter → Column
                'area': 'line_markers',       # Area → Line
                'stacked_bar': 'column',      # Stacked → Column
                'stacked_column': 'column',   # Stacked → Column
                'horizontal_bar': 'bar',      # Keep as Bar
                'bubble': 'column',           # Bubble → Column
                'waterfall': 'column',        # Waterfall → Column
                
                # Intent-based mappings
                'trend': 'line_markers',
                'distribution': 'doughnut',
                'comparison': 'column',
                'ranking': 'bar',
                'breakdown': 'doughnut',
                'composition': 'doughnut'
            }
            
            mapped_type = chart_mapping.get(chart_type.lower(), 'column')
            
            # Quality gate: Override if confidence is too low (< 70%)
            if confidence < 0.7:
                print(f"   ⚠️  Low AI confidence ({confidence:.0%}), using safe column chart")
                mapped_type = 'column'
            
            print(f"   🤖 AI recommends: '{chart_type}' → mapped to '{mapped_type}' (confidence: {confidence:.0%})")
            
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
        """
        Analyze data structure to determine best chart type
        ONLY USES 4 PRACTICAL CHART TYPES: Column, Bar, Doughnut, Line
        """
        numeric_cols = data.select_dtypes(include=[np.number]).columns
        text_cols = data.select_dtypes(include=['object']).columns
        date_cols = data.select_dtypes(include=['datetime64']).columns
        
        n_rows = len(data)
        n_numeric = len(numeric_cols)
        n_text = len(text_cols)
        
        # ============================================================
        # CONTEXT-FIRST APPROACH (Most Important)
        # ============================================================
        
        if context in ['distribution', 'breakdown', 'composition', 'sector']:
            # Distribution context → ALWAYS use doughnut
            print("   📊 Context 'distribution' → Doughnut chart")
            return {'type': 'doughnut', 'ppt_type': XL_CHART_TYPE.DOUGHNUT, 'source': 'context'}
        
        if context in ['trend', 'time', 'timeline', 'growth', 'over_time']:
            # Trend context → ALWAYS use line chart
            print("   📈 Context 'trend' → Line chart")
            return {'type': 'line_markers', 'ppt_type': XL_CHART_TYPE.LINE_MARKERS, 'source': 'context'}
        
        if context in ['comparison', 'ranking', 'top', 'versus']:
            # Comparison context → Column or bar
            if n_rows <= 8:
                print("   📊 Context 'comparison' → Column chart")
                return {'type': 'column', 'ppt_type': XL_CHART_TYPE.COLUMN_CLUSTERED, 'source': 'context'}
            else:
                print("   📊 Context 'comparison' → Bar chart")
                return {'type': 'bar', 'ppt_type': XL_CHART_TYPE.BAR_CLUSTERED, 'source': 'context'}
        
        # ============================================================
        # DATA STRUCTURE ANALYSIS (Fallback)
        # ============================================================
        
        # Few items with categories → Doughnut
        if n_text > 0 and n_numeric == 1 and n_rows <= 6:
            print("   📊 Few items → Doughnut chart")
            return {'type': 'doughnut', 'ppt_type': XL_CHART_TYPE.DOUGHNUT, 'source': 'structure'}
        
        # Many items → Bar chart (horizontal for long labels)
        if n_text > 0 and n_rows > 10:
            print("   📊 Many items → Bar chart")
            return {'type': 'bar', 'ppt_type': XL_CHART_TYPE.BAR_CLUSTERED, 'source': 'structure'}
        
        # Default: Column chart (most universal)
        print("   📊 Default → Column chart")
        return {'type': 'column', 'ppt_type': XL_CHART_TYPE.COLUMN_CLUSTERED, 'source': 'default'}
    
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
        """Prepare data for category-based charts with intelligent limits and value validation"""
        chart_data = CategoryChartData()
        
        numeric_cols = data.select_dtypes(include=[np.number]).columns
        text_cols = data.select_dtypes(include=['object']).columns
        
        # Intelligent data limits based on chart type
        if chart_type in ['pie', 'doughnut']:
            max_items = 6  # Fewer for pie charts (easier to read)
        elif chart_type in ['bar', 'column']:
            max_items = 10  # Good balance for bar/column
        else:
            max_items = 12  # More for other types
        
        # Filter out zero/negative values for pie/doughnut charts (they don't make sense)
        if chart_type in ['pie', 'doughnut'] and len(numeric_cols) > 0:
            value_col = numeric_cols[0]
            data = data[data[value_col] > 0].copy()
        
        # Determine categories and sort by value for better visualization
        if len(text_cols) > 0 and len(numeric_cols) > 0:
            # Sort by first numeric column (descending) to show top performers
            value_col = numeric_cols[0]
            sorted_data = data.sort_values(by=value_col, ascending=False)
            
            # Remove any rows with invalid/null category names
            sorted_data = sorted_data[sorted_data[text_cols[0]].notna()]
            
            categories = sorted_data[text_cols[0]].head(max_items).tolist()
            data_subset = sorted_data.head(max_items)
            
            # Truncate long category names for readability
            categories = [str(cat)[:30] + '...' if len(str(cat)) > 30 else str(cat) 
                         for cat in categories]
            
            # Ensure we have valid data
            if len(categories) == 0:
                categories = ['No Valid Data']
                data_subset = pd.DataFrame({value_col: [0]})
                
        elif len(numeric_cols) > 0:
            # No text columns, use generic labels
            data_subset = data.head(max_items)
            categories = [f"Item {i+1}" for i in range(len(data_subset))]
        else:
            # Fallback
            categories = ['No Data']
            data_subset = pd.DataFrame({'Value': [0]})
        
        chart_data.categories = categories
        
        # Add series based on chart type - SIMPLIFIED
        if chart_type in ['column_stacked', 'column_stacked_100', 'area_stacked']:
            # Multiple series - limit to 2 for clarity (not 3)
            for i, col in enumerate(numeric_cols[:2]):
                # Clean column name
                clean_name = str(col)[:20]
                values = data_subset[col].fillna(0).tolist()
                
                # Smart rounding based on magnitude
                rounded_values = []
                for v in values:
                    v = float(v)
                    if abs(v) < 0.01:
                        rounded_values.append(0)
                    elif abs(v) < 10:
                        rounded_values.append(round(v, 2))
                    elif abs(v) < 1000:
                        rounded_values.append(round(v, 1))
                    else:
                        rounded_values.append(round(v, 0))
                
                chart_data.add_series(clean_name, rounded_values)
        else:
            # Single series - MOST COMMON
            if len(numeric_cols) > 0:
                col = numeric_cols[0]
                clean_name = str(col)[:20]
                values = data_subset[col].fillna(0).tolist()
                
                # Smart rounding for readability
                rounded_values = []
                for v in values:
                    v = float(v)
                    if abs(v) < 0.01:
                        rounded_values.append(0)
                    elif abs(v) < 10:
                        rounded_values.append(round(v, 2))
                    elif abs(v) < 1000:
                        rounded_values.append(round(v, 1))
                    else:
                        rounded_values.append(round(v, 0))
                
                chart_data.add_series(clean_name, rounded_values)
        
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
        """Apply professional styling to chart with enhanced visuals"""
        
        # ============ TITLE STYLING ============
        if title:
            chart.has_title = True
            chart.chart_title.text_frame.text = title
            title_para = chart.chart_title.text_frame.paragraphs[0]
            # Slightly larger title for better presentation
            title_para.font.size = Pt(22)
            title_para.font.bold = True
            title_para.font.color.rgb = RGBColor(25, 42, 86)  # Navy blue
        
        # ============ LEGEND STYLING ============
        chart.has_legend = True
        # Increase legend size and add breathing room
        try:
            if chart_type not in ['pie', 'doughnut']:
                chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart.legend.font.size = Pt(12)
            else:
                chart.legend.position = XL_LEGEND_POSITION.RIGHT
                chart.legend.font.size = Pt(11)

            # Make legend text clearer
            chart.legend.font.bold = False
            # Do not force legend into layout so it doesn't squeeze the plot area
            try:
                chart.legend.include_in_layout = False
            except Exception:
                pass
        except Exception as e:
            print(f"   Note: Legend styling not fully supported: {e}")
        
        # ============ DATA LABELS ============
        try:
            plot = chart.plots[0]

            if chart_type in ['pie', 'doughnut']:
                # Pie/Doughnut: Show percentages
                plot.has_data_labels = True
                data_labels = plot.data_labels
                # Slightly larger labels for readability
                data_labels.font.size = Pt(12)
                data_labels.font.bold = True
                data_labels.font.color.rgb = RGBColor(255, 255, 255)  # White text
                data_labels.position = XL_LABEL_POSITION.INSIDE_END

                # Show percentage
                try:
                    data_labels.number_format = '0%'
                    data_labels.show_percentage = True
                    data_labels.show_value = False
                except:
                    pass

            elif chart_type in ['column', 'bar']:
                # Bar/Column: Show values on top and increase spacing
                plot.has_data_labels = True
                data_labels = plot.data_labels
                data_labels.font.size = Pt(10)
                data_labels.font.bold = True
                data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
                try:
                    data_labels.number_format = '#,##0'
                except:
                    pass

                # Try to increase gap/spacing between bars for readability (best-effort)
                try:
                    for ser in chart.series:
                        # no-op placeholder for series-level spacing adjustments; kept for future expansion
                        _ = getattr(ser, 'format', None)
                except Exception:
                    pass

        except Exception as e:
            print(f"   Note: Data labels not fully supported: {e}")
        
        # ============ AXIS STYLING ============
        try:
            # Category axis (X-axis for column, Y-axis for bar)
            category_axis = chart.category_axis if hasattr(chart, 'category_axis') else None
            if category_axis:
                # Positioning and font improvements for category labels
                try:
                    category_axis.tick_label_position = 'low'
                except Exception:
                    pass
                category_axis.tick_labels.font.size = Pt(11)
                category_axis.tick_labels.font.color.rgb = RGBColor(89, 89, 89)

                # Rotate labels if needed (less extreme angle)
                if chart_type == 'column':
                    try:
                        category_axis.tick_labels.rotation = -30  # Angle labels for readability
                    except:
                        pass
            
            # Value axis (Y-axis for column, X-axis for bar)
            value_axis = chart.value_axis if hasattr(chart, 'value_axis') else None
            if value_axis:
                # Improve value axis readability
                value_axis.tick_labels.font.size = Pt(11)
                value_axis.tick_labels.font.color.rgb = RGBColor(89, 89, 89)
                value_axis.tick_labels.number_format = '#,##0'

                # Show major gridlines for easier reading and subtle style
                value_axis.has_major_gridlines = True
                value_axis.has_minor_gridlines = False
                try:
                    # best-effort: make gridlines lighter if possible
                    mg = value_axis.major_gridlines
                    try:
                        mg.format.line.color.rgb = RGBColor(220, 220, 220)
                    except Exception:
                        pass
                except Exception:
                    pass

                # Set visible min to 0 for bar/column charts
                if chart_type in ['column', 'bar', 'column_stacked']:
                    try:
                        value_axis.minimum_scale = 0
                    except:
                        pass
        
        except Exception as e:
            print(f"   Note: Axis styling not fully supported: {e}")
        
        # ============ SERIES COLORS (Professional palette) ============
        try:
            professional_colors = [
                RGBColor(41, 128, 185),   # Professional Blue
                RGBColor(39, 174, 96),    # Success Green
                RGBColor(230, 126, 34),   # Warning Orange
                RGBColor(231, 76, 60),    # Danger Red
                RGBColor(142, 68, 173),   # Royal Purple
                RGBColor(241, 196, 15),   # Gold
                RGBColor(52, 152, 219),   # Light Blue
                RGBColor(26, 188, 156),   # Turquoise
            ]
            
            for idx, series in enumerate(chart.series):
                color = professional_colors[idx % len(professional_colors)]

                # Apply color to series (safe operations)
                try:
                    fill = series.format.fill
                    fill.solid()
                    fill.fore_color.rgb = color
                except Exception:
                    pass

                # Add smooth lines for line charts (best-effort)
                if chart_type in ['line', 'line_markers']:
                    try:
                        series.smooth = True
                    except:
                        pass

            # Provide additional breathing room by reducing visual clutter
            try:
                # If chart has a plot_area, set white fill so plot area looks cleaner
                chart.plot_area.format.fill.solid()
                chart.plot_area.format.fill.fore_color.rgb = RGBColor(255, 255, 255)
            except Exception:
                pass
        
        except Exception as e:
            print(f"   Note: Series colors not fully applied: {e}")
    
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
# HELPER FUNCTIONS FOR BETTER CHART READABILITY
# ============================================================================

def format_large_number(value: float) -> str:
    """Format large numbers with K, M, B suffixes for cleaner display"""
    try:
        value = float(value)
        if abs(value) >= 1e9:
            return f"{value/1e9:.1f}B"
        elif abs(value) >= 1e6:
            return f"{value/1e6:.1f}M"
        elif abs(value) >= 1e3:
            return f"{value/1e3:.1f}K"
        else:
            return f"{value:.1f}"
    except:
        return str(value)


def get_trend_arrow(current: float, previous: float) -> str:
    """Get trend arrow emoji based on change"""
    try:
        if current > previous:
            return "↗️"
        elif current < previous:
            return "↘️"
        else:
            return "→"
    except:
        return "→"


def calculate_percentage_change(current: float, previous: float) -> str:
    """Calculate and format percentage change"""
    try:
        if previous == 0:
            return "N/A"
        change = ((current - previous) / previous) * 100
        sign = "+" if change > 0 else ""
        return f"{sign}{change:.1f}%"
    except:
        return "N/A"


def clean_chart_title(title: str, max_length: int = 50) -> str:
    """Clean and truncate chart titles for better display"""
    if not title:
        return "Data Visualization"
    
    title = str(title).strip()
    if len(title) > max_length:
        return title[:max_length-3] + "..."
    return title


def round_to_significant(value: float, sig_figs: int = 3) -> float:
    """Round number to significant figures for cleaner display"""
    try:
        if value == 0:
            return 0
        import math
        return round(value, -int(math.floor(math.log10(abs(value)))) + (sig_figs - 1))
    except:
        return round(value, 2)
