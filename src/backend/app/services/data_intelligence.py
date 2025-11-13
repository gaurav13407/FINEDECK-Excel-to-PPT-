"""
FinDeck Data Intelligence Engine
Automatically analyzes any Excel file and produces structured, accurate insights for PPT generation.
Zero hallucination - 100% based on actual data.
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Tuple
from datetime import datetime
import re
from collections import Counter


class DataIntelligenceEngine:
    """
    Analyzes any uploaded Excel file and produces structured insights.
    Works with any dataset structure - Sales, Marketing, Finance, HR, Health, etc.
    """
    
    # Known identifier column patterns
    IDENTIFIER_PATTERNS = [
        r'.*_?id$', r'^id_?.*', r'order.*id', r'customer.*id', r'product.*id',
        r'row.*id', r'transaction.*id', r'invoice.*id', r'postal.*code',
        r'zip.*code', r'phone', r'email', r'ssn', r'account.*number'
    ]
    
    # Known sales/revenue patterns
    SALES_PATTERNS = [
        r'sales?', r'revenue', r'amount', r'total.*price', r'value',
        r'gmv', r'gross.*merchandise', r'turnover'
    ]
    
    # Known profit patterns
    PROFIT_PATTERNS = [
        r'profit', r'margin', r'net.*income', r'earnings', r'ebitda'
    ]
    
    # Known geo patterns
    GEO_PATTERNS = [
        r'city', r'state', r'country', r'region', r'postal.*code',
        r'zip.*code', r'location', r'territory', r'area', r'zone'
    ]
    
    def __init__(self):
        self.df = None
        self.column_types = {}
        self.primary_metrics = {}
        self.insights = {}
        
    def analyze_file(self, file_path: str) -> Dict[str, Any]:
        """
        Main entry point - analyze any Excel file and return structured insights.
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            Dictionary with complete analysis results in PPT-ready format
        """
        # Load the Excel file
        self.df = pd.read_excel(file_path)
        
        # STEP 1: Classify all columns
        self._classify_columns()
        
        # STEP 2: Extract metrics for each column type
        self._extract_all_metrics()
        
        # STEP 3: Identify and analyze hierarchy (Sales, Profit, Geo)
        self._analyze_hierarchy()
        
        # STEP 4: Generate PPT-ready output
        output = self._generate_ppt_output()
        
        return output
    
    def _classify_columns(self):
        """
        STEP 1: Detect and classify each column automatically.
        """
        for col in self.df.columns:
            col_lower = col.lower().strip()
            
            # Check if identifier (never sum these!)
            if self._is_identifier(col_lower):
                self.column_types[col] = 'identifier'
                continue
            
            # Check data type
            dtype = self.df[col].dtype
            
            if pd.api.types.is_numeric_dtype(dtype):
                # Additional check: if all unique values (might be ID)
                if len(self.df[col].unique()) == len(self.df):
                    self.column_types[col] = 'identifier'
                else:
                    self.column_types[col] = 'numeric'
                    
            elif pd.api.types.is_datetime64_any_dtype(dtype):
                self.column_types[col] = 'date'
                
            elif pd.api.types.is_bool_dtype(dtype):
                self.column_types[col] = 'boolean'
                
            else:
                # Try to convert to datetime
                try:
                    pd.to_datetime(self.df[col], errors='raise')
                    self.column_types[col] = 'date'
                except:
                    # Check if categorical (limited unique values)
                    unique_ratio = len(self.df[col].unique()) / len(self.df)
                    if unique_ratio < 0.5:  # Less than 50% unique = categorical
                        self.column_types[col] = 'categorical'
                    else:
                        self.column_types[col] = 'text'
    
    def _is_identifier(self, col_name: str) -> bool:
        """Check if column is an identifier (should never be summed)."""
        for pattern in self.IDENTIFIER_PATTERNS:
            if re.search(pattern, col_name, re.IGNORECASE):
                return True
        return False
    
    def _extract_all_metrics(self):
        """
        STEP 2: Extract metrics for each column type.
        """
        self.insights['numeric_metrics'] = {}
        self.insights['categorical_metrics'] = {}
        self.insights['date_metrics'] = {}
        self.insights['boolean_metrics'] = {}
        
        for col, col_type in self.column_types.items():
            if col_type == 'numeric':
                self.insights['numeric_metrics'][col] = self._analyze_numeric(col)
            elif col_type == 'categorical':
                self.insights['categorical_metrics'][col] = self._analyze_categorical(col)
            elif col_type == 'date':
                self.insights['date_metrics'][col] = self._analyze_date(col)
            elif col_type == 'boolean':
                self.insights['boolean_metrics'][col] = self._analyze_boolean(col)
    
    def _analyze_numeric(self, col: str) -> Dict[str, Any]:
        """Analyze numeric column - calculate sum, mean, median, min, max, std, top 10."""
        data = self.df[col].dropna()
        
        if len(data) == 0:
            return {}
        
        # Top 10 values (excluding duplicates, sorted)
        top_10 = data.nlargest(10).tolist()
        
        return {
            'sum': float(data.sum()),
            'mean': float(data.mean()),
            'median': float(data.median()),
            'min': float(data.min()),
            'max': float(data.max()),
            'std': float(data.std()) if len(data) > 1 else 0.0,
            'count': int(len(data)),
            'top_10_values': top_10
        }
    
    def _analyze_categorical(self, col: str) -> Dict[str, Any]:
        """Analyze categorical column - unique count, top 10, distribution %."""
        data = self.df[col].dropna()
        
        if len(data) == 0:
            return {}
        
        value_counts = data.value_counts()
        total = len(data)
        
        # Top 10 categories
        top_10 = value_counts.head(10)
        
        # Percentage distribution
        distribution = []
        for category, count in top_10.items():
            distribution.append({
                'category': str(category),
                'count': int(count),
                'percentage': round((count / total) * 100, 2)
            })
        
        return {
            'unique_count': int(data.nunique()),
            'total_count': int(total),
            'top_10_categories': distribution,
            'most_common': str(value_counts.index[0]) if len(value_counts) > 0 else None
        }
    
    def _analyze_date(self, col: str) -> Dict[str, Any]:
        """Analyze date column - monthly/yearly grouping, trend detection."""
        data = pd.to_datetime(self.df[col], errors='coerce').dropna()
        
        if len(data) == 0:
            return {}
        
        # Monthly grouping
        monthly = data.dt.to_period('M').value_counts().sort_index()
        monthly_data = [
            {'month': str(period), 'count': int(count)}
            for period, count in monthly.items()
        ]
        
        # Yearly grouping
        yearly = data.dt.to_period('Y').value_counts().sort_index()
        yearly_data = [
            {'year': str(period), 'count': int(count)}
            for period, count in yearly.items()
        ]
        
        # Trend detection (simple: increasing, decreasing, stable)
        if len(monthly) > 1:
            trend = 'increasing' if monthly.iloc[-1] > monthly.iloc[0] else \
                    'decreasing' if monthly.iloc[-1] < monthly.iloc[0] else 'stable'
        else:
            trend = 'insufficient_data'
        
        return {
            'earliest': str(data.min().date()),
            'latest': str(data.max().date()),
            'monthly_distribution': monthly_data[:12],  # Last 12 months
            'yearly_distribution': yearly_data,
            'trend': trend
        }
    
    def _analyze_boolean(self, col: str) -> Dict[str, Any]:
        """Analyze boolean column - true/false counts and percentages."""
        data = self.df[col].dropna()
        
        if len(data) == 0:
            return {}
        
        value_counts = data.value_counts()
        total = len(data)
        
        return {
            'true_count': int(value_counts.get(True, 0)),
            'false_count': int(value_counts.get(False, 0)),
            'true_percentage': round((value_counts.get(True, 0) / total) * 100, 2),
            'false_percentage': round((value_counts.get(False, 0) / total) * 100, 2)
        }
    
    def _analyze_hierarchy(self):
        """
        STEP 3: Identify and analyze hierarchy (Sales, Profit, Geo).
        """
        self.insights['hierarchy'] = {}
        
        # Find Sales/Revenue column
        sales_col = self._find_column(self.SALES_PATTERNS)
        if sales_col and self.column_types.get(sales_col) == 'numeric':
            self.insights['hierarchy']['sales'] = self._analyze_sales(sales_col)
        
        # Find Profit column
        profit_col = self._find_column(self.PROFIT_PATTERNS)
        if profit_col and self.column_types.get(profit_col) == 'numeric':
            self.insights['hierarchy']['profit'] = self._analyze_profit(profit_col, sales_col)
        
        # Find Geo columns
        geo_cols = [col for col in self.df.columns 
                   if any(re.search(pattern, col.lower(), re.IGNORECASE) 
                         for pattern in self.GEO_PATTERNS)]
        if geo_cols:
            self.insights['hierarchy']['geography'] = self._analyze_geography(geo_cols, sales_col)
    
    def _find_column(self, patterns: List[str]) -> str:
        """Find first column matching any pattern."""
        for col in self.df.columns:
            col_lower = col.lower().strip()
            for pattern in patterns:
                if re.search(pattern, col_lower, re.IGNORECASE):
                    return col
        return None
    
    def _analyze_sales(self, sales_col: str) -> Dict[str, Any]:
        """Analyze sales column comprehensively."""
        data = self.df[sales_col].dropna()
        
        result = {
            'column_name': sales_col,
            'total_sales': float(data.sum()),
            'average_sale': float(data.mean()),
            'min_sale': float(data.min()),
            'max_sale': float(data.max())
        }
        
        # Sales by Category (if categorical column exists)
        categorical_cols = [col for col, ctype in self.column_types.items() 
                           if ctype == 'categorical']
        if categorical_cols:
            cat_col = categorical_cols[0]  # Use first categorical column
            sales_by_cat = self.df.groupby(cat_col)[sales_col].sum().sort_values(ascending=False)
            result['sales_by_category'] = {
                'category_column': cat_col,
                'breakdown': [
                    {'category': str(cat), 'sales': float(sales)}
                    for cat, sales in sales_by_cat.head(10).items()
                ]
            }
        
        # Monthly Sales Trend (if date column exists)
        date_cols = [col for col, ctype in self.column_types.items() if ctype == 'date']
        if date_cols:
            date_col = date_cols[0]
            df_copy = self.df[[date_col, sales_col]].copy()
            df_copy[date_col] = pd.to_datetime(df_copy[date_col], errors='coerce')
            df_copy = df_copy.dropna()
            df_copy['month'] = df_copy[date_col].dt.to_period('M')
            monthly_sales = df_copy.groupby('month')[sales_col].sum().sort_index()
            
            result['monthly_trend'] = [
                {'month': str(month), 'sales': float(sales)}
                for month, sales in monthly_sales.items()
            ]
        
        return result
    
    def _analyze_profit(self, profit_col: str, sales_col: str = None) -> Dict[str, Any]:
        """Analyze profit column."""
        data = self.df[profit_col].dropna()
        
        result = {
            'column_name': profit_col,
            'total_profit': float(data.sum()),
            'average_profit': float(data.mean()),
            'min_profit': float(data.min()),
            'max_profit': float(data.max())
        }
        
        # Profit Margin (if sales column exists)
        if sales_col and sales_col in self.df.columns:
            df_clean = self.df[[sales_col, profit_col]].dropna()
            total_sales = df_clean[sales_col].sum()
            total_profit = df_clean[profit_col].sum()
            if total_sales > 0:
                result['profit_margin_percentage'] = round((total_profit / total_sales) * 100, 2)
        
        # Profit by Category
        categorical_cols = [col for col, ctype in self.column_types.items() 
                           if ctype == 'categorical']
        if categorical_cols:
            cat_col = categorical_cols[0]
            profit_by_cat = self.df.groupby(cat_col)[profit_col].sum().sort_values(ascending=False)
            result['profit_by_category'] = [
                {'category': str(cat), 'profit': float(profit)}
                for cat, profit in profit_by_cat.head(10).items()
            ]
        
        return result
    
    def _analyze_geography(self, geo_cols: List[str], sales_col: str = None) -> Dict[str, Any]:
        """Analyze geographic distribution."""
        result = {}
        
        for geo_col in geo_cols[:3]:  # Analyze up to 3 geo columns
            if self.column_types.get(geo_col) in ['categorical', 'text']:
                data = self.df[geo_col].dropna()
                
                # Top locations by count
                top_locations = data.value_counts().head(10)
                
                geo_analysis = {
                    'column_name': geo_col,
                    'unique_locations': int(data.nunique()),
                    'top_locations_by_count': [
                        {'location': str(loc), 'count': int(count)}
                        for loc, count in top_locations.items()
                    ]
                }
                
                # Top locations by sales (if sales column exists)
                if sales_col and sales_col in self.df.columns:
                    df_clean = self.df[[geo_col, sales_col]].dropna()
                    sales_by_location = df_clean.groupby(geo_col)[sales_col].sum().sort_values(ascending=False)
                    geo_analysis['top_locations_by_sales'] = [
                        {'location': str(loc), 'sales': float(sales)}
                        for loc, sales in sales_by_location.head(10).items()
                    ]
                
                result[geo_col] = geo_analysis
        
        return result
    
    def _generate_ppt_output(self) -> Dict[str, Any]:
        """
        STEP 4: Generate PPT-ready output structure.
        All values 100% based on Excel data - zero hallucination.
        """
        # Count column types
        type_counts = Counter(self.column_types.values())
        
        # Executive Summary
        executive_summary = {
            'total_records': int(len(self.df)),
            'total_columns': int(len(self.df.columns)),
            'numeric_columns': int(type_counts.get('numeric', 0)),
            'categorical_columns': int(type_counts.get('categorical', 0)),
            'date_columns': int(type_counts.get('date', 0)),
            'identifier_columns': int(type_counts.get('identifier', 0)),
            'boolean_columns': int(type_counts.get('boolean', 0)),
            'text_columns': int(type_counts.get('text', 0))
        }
        
        # Key Metrics (top numeric metrics)
        key_metrics = []
        for col, metrics in list(self.insights['numeric_metrics'].items())[:5]:
            key_metrics.append({
                'metric_name': col,
                'total': metrics.get('sum', 0),
                'average': metrics.get('mean', 0),
                'min': metrics.get('min', 0),
                'max': metrics.get('max', 0)
            })
        
        # Top Categories (from categorical columns)
        top_categories = []
        for col, metrics in list(self.insights['categorical_metrics'].items())[:3]:
            top_categories.append({
                'category_type': col,
                'unique_count': metrics.get('unique_count', 0),
                'top_values': metrics.get('top_10_categories', [])[:5]
            })
        
        # Distribution Insights
        distribution_insights = []
        for col, metrics in self.insights['categorical_metrics'].items():
            if metrics.get('top_10_categories'):
                distribution_insights.append({
                    'column': col,
                    'distribution': metrics['top_10_categories']
                })
        
        # Trend Analysis (from date columns)
        trend_analysis = {}
        for col, metrics in self.insights['date_metrics'].items():
            trend_analysis[col] = {
                'date_range': f"{metrics.get('earliest', 'N/A')} to {metrics.get('latest', 'N/A')}",
                'trend': metrics.get('trend', 'unknown'),
                'monthly_data': metrics.get('monthly_distribution', [])
            }
        
        # Recommendations (data-driven)
        recommendations = self._generate_recommendations()
        
        # Final PPT-ready output
        output = {
            'executive_summary': executive_summary,
            'key_metrics': key_metrics,
            'top_categories': top_categories,
            'top_numeric_metrics': list(self.insights['numeric_metrics'].values())[:10],
            'distribution_insights': distribution_insights,
            'trend_analysis': trend_analysis,
            'hierarchy_analysis': self.insights.get('hierarchy', {}),
            'recommendations': recommendations,
            'column_classification': self.column_types
        }
        
        return output
    
    def _generate_recommendations(self) -> List[str]:
        """Generate data-driven recommendations based on actual insights."""
        recommendations = []
        
        # Check for Sales/Revenue
        if 'sales' in self.insights.get('hierarchy', {}):
            sales_data = self.insights['hierarchy']['sales']
            recommendations.append(
                f"Total sales of {sales_data['total_sales']:,.2f} with average "
                f"transaction of {sales_data['average_sale']:,.2f}"
            )
        
        # Check for Profit
        if 'profit' in self.insights.get('hierarchy', {}):
            profit_data = self.insights['hierarchy']['profit']
            if 'profit_margin_percentage' in profit_data:
                recommendations.append(
                    f"Profit margin is {profit_data['profit_margin_percentage']}%"
                )
        
        # Check for trends
        for col, metrics in self.insights.get('date_metrics', {}).items():
            if metrics.get('trend') == 'increasing':
                recommendations.append(f"Positive {col} trend detected")
            elif metrics.get('trend') == 'decreasing':
                recommendations.append(f"Declining {col} trend - needs attention")
        
        # Check categorical concentration
        for col, metrics in self.insights.get('categorical_metrics', {}).items():
            if metrics.get('top_10_categories'):
                top_cat = metrics['top_10_categories'][0]
                if top_cat['percentage'] > 50:
                    recommendations.append(
                        f"High concentration in {col}: {top_cat['category']} "
                        f"represents {top_cat['percentage']}% of data"
                    )
        
        return recommendations if recommendations else ["Dataset analyzed successfully"]


# Usage example
def analyze_excel_file(file_path: str) -> Dict[str, Any]:
    """
    Main function to analyze any Excel file.
    
    Args:
        file_path: Path to Excel file
        
    Returns:
        PPT-ready structured insights
    """
    engine = DataIntelligenceEngine()
    results = engine.analyze_file(file_path)
    return results
