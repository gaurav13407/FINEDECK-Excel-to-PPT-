"""
Smart Chart Analyzer - Intelligently determines best charts for the data
"""

import pandas as pd
from typing import Dict, List, Tuple, Any
import numpy as np


class SmartChartAnalyzer:
    """Analyzes data structure and recommends optimal chart types"""
    
    def __init__(self, summary_df: pd.DataFrame, sheets_data: List[Tuple[str, pd.DataFrame]]):
        self.summary_df = summary_df
        self.sheets_data = sheets_data
        self.recommendations = []
        
    def analyze(self) -> Dict[str, Any]:
        """Analyze data and return chart recommendations"""
        
        charts = {
            'dashboard_charts': [],
            'comparison_charts': [],
            'trend_charts': [],
            'distribution_charts': []
        }
        
        # Analyze Summary sheet
        if self.summary_df is not None and not self.summary_df.empty:
            summary_analysis = self._analyze_summary()
            charts['dashboard_charts'].extend(summary_analysis['dashboard'])
            charts['comparison_charts'].extend(summary_analysis['comparison'])
        
        # Analyze time-series sheets
        time_series_analysis = self._analyze_time_series()
        if time_series_analysis:
            charts['trend_charts'].extend(time_series_analysis)
        
        # Analyze categorical distributions
        distribution_analysis = self._analyze_distributions()
        if distribution_analysis:
            charts['distribution_charts'].extend(distribution_analysis)
        
        return charts
    
    def _analyze_summary(self) -> Dict[str, List[Dict]]:
        """Analyze summary data for chart opportunities"""
        df = self.summary_df
        charts = {'dashboard': [], 'comparison': []}
        
        # Check for numeric columns
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
        
        # 1. Bar Chart: Top items by value
        if 'MarketCap' in df.columns or 'Sales' in df.columns or 'Profit' in df.columns:
            value_col = 'MarketCap' if 'MarketCap' in df.columns else ('Sales' if 'Sales' in df.columns else 'Profit')
            label_col = 'Ticker' if 'Ticker' in df.columns else ('Name' if 'Name' in df.columns else df.columns[0])
            
            charts['dashboard'].append({
                'type': 'bar_horizontal',
                'title': f'Top Items by {value_col}',
                'x_col': label_col,
                'y_col': value_col,
                'sort': 'desc',
                'limit': 10,
                'format': 'currency' if value_col in ['Sales', 'Profit', 'MarketCap'] else 'number'
            })
        
        # 2. Comparison Chart: Multiple metrics
        if len(numeric_cols) >= 2:
            charts['dashboard'].append({
                'type': 'grouped_bar',
                'title': 'Multi-Metric Comparison',
                'x_col': categorical_cols[0] if categorical_cols else df.columns[0],
                'y_cols': numeric_cols[:3],  # Up to 3 metrics
                'limit': 8
            })
        
        # 3. Horizontal Bar Chart: Category distribution (better than pie)
        if 'Sector' in df.columns:
            charts['comparison'].append({
                'type': 'bar_horizontal',
                'title': 'Distribution by Sector',
                'x_col': 'Sector',
                'y_col': value_col if value_col in df.columns else 'Sales',
                'sort': 'desc',
                'limit': 10,
                'format': 'currency' if value_col in ['Sales', 'Profit', 'MarketCap'] else 'number'
            })
        
        # 4. Scatter Plot: Correlation analysis
        if 'Return_1Y_%' in df.columns and 'TrailingPE' in df.columns:
            charts['comparison'].append({
                'type': 'scatter',
                'title': 'Return vs P/E Ratio',
                'x_col': 'TrailingPE',
                'y_col': 'Return_1Y_%',
                'label_col': 'Ticker'
            })
        
        return charts
    
    def _analyze_time_series(self) -> List[Dict]:
        """Analyze time-series data for trend charts"""
        charts = []
        
        for sheet_name, df in self.sheets_data:
            # Check if it's a time-series sheet
            if df is None or df.empty:
                continue
                
            date_cols = [col for col in df.columns if 'date' in col.lower()]
            if not date_cols:
                continue
            
            date_col = date_cols[0]
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            
            if len(numeric_cols) >= 1:
                # Line chart for main metric
                charts.append({
                    'type': 'line',
                    'title': f'{sheet_name.replace("_", " ")} Trend',
                    'sheet': sheet_name,
                    'x_col': date_col,
                    'y_col': numeric_cols[0],
                    'limit': 60  # Last 60 data points
                })
        
        return charts
    
    def _analyze_distributions(self) -> List[Dict]:
        """Analyze data for distribution charts"""
        charts = []
        
        # Check for segment/category analysis in sheets
        for sheet_name, df in self.sheets_data:
            if df is None or df.empty:
                continue
            
            # Look for segment or category columns
            if 'Segment' in df.columns:
                numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
                if numeric_cols:
                    charts.append({
                        'type': 'stacked_bar',
                        'title': 'Sales by Segment',
                        'sheet': sheet_name,
                        'category_col': 'Segment',
                        'value_col': numeric_cols[0]
                    })
            
            if 'Country' in df.columns:
                numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
                if numeric_cols:
                    charts.append({
                        'type': 'bar_chart',
                        'title': 'Performance by Country',
                        'sheet': sheet_name,
                        'category_col': 'Country',
                        'value_col': numeric_cols[0],
                        'limit': 10
                    })
        
        return charts
    
    def get_recommended_charts(self, slide_position: str) -> List[Dict]:
        """Get recommended charts for a specific slide position"""
        analysis = self.analyze()
        
        if slide_position == 'dashboard':
            return analysis['dashboard_charts'][:2]  # Top 2 dashboard charts
        elif slide_position == 'trends':
            return analysis['trend_charts'][:2]  # Top 2 trend charts
        elif slide_position == 'comparison':
            return analysis['comparison_charts'][:1]  # 1 comparison chart
        elif slide_position == 'distribution':
            return analysis['distribution_charts'][:1]  # 1 distribution chart
        
        return []


def format_value(value: float, format_type: str = 'number') -> str:
    """Format values for display with smart scaling"""
    if pd.isna(value) or (isinstance(value, float) and np.isnan(value)):
        return "N/A"
    
    # Convert to float if not already
    try:
        value = float(value)
    except:
        return str(value)
    
    if format_type == 'currency':
        # Smart scaling - use appropriate unit
        if abs(value) >= 1e9:
            return f"${value/1e9:.2f}B"
        elif abs(value) >= 1e6:
            return f"${value/1e6:.2f}M"
        elif abs(value) >= 1e3:
            return f"${value/1e3:.2f}K"
        elif abs(value) >= 1:
            return f"${value:.2f}"
        else:
            return f"${value:.4f}"
    
    elif format_type == 'percentage':
        return f"{value:.1f}%"
    
    elif format_type == 'number':
        # Smart scaling for numbers
        if abs(value) >= 1e9:
            return f"{value/1e9:.2f}B"
        elif abs(value) >= 1e6:
            return f"{value/1e6:.2f}M"
        elif abs(value) >= 1e3:
            return f"{value/1e3:.2f}K"
        elif abs(value) >= 1:
            return f"{value:.2f}"
        else:
            return f"{value:.4f}"
    
    return str(value)
