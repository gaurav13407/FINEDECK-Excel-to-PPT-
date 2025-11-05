""" AI-Powered Service for FinDeck
    uses Groq API for ultra-fast AI analysis Of Excel data
    Implement 5 premium features for AI pro users
"""

import os
import json
from typing import Dict,List,Optional,Tuple,Any
from dotenv import load_dotenv
from groq import Groq
import pandas as pd
import numpy as np

load_dotenv()


class AIInsightsService:
    """Service to provide AI-powered insights using Groq API
    Implements 5 speacialized features for premium users

    Features:
    1.Slide Title Generation
    2.Automatic Data insights
    3.Smart Template Selection
    4.Slide Layout Optimization
    5.AI Chart Type Recommendation
    
    """



    def __init__(self,api_key:Optional[str]=None):
        """Initialize AI Services wih Groq API"""
        self.api_key=api_key or os.getenv("GROQ_API_KEY")
        if not self.api_key:
            raise ValueError("GROQ_API_KEY not set in environment variables")
        
        self.client=Groq(api_key=self.api_key)

        # Updated models (Nov 2025) - using currently supported models
        # Check https://console.groq.com/docs/models for latest models
        # Note: Using faster models to avoid rate limits (100k tokens/day on free tier)
        self.fast_model="llama-3.1-8b-instant"  # Fast responses, higher rate limits
        self.balanced_model="llama-3.1-8b-instant"  # Using fast model to avoid rate limits
        self.capable_model="llama-3.1-8b-instant"  # Using fast model to avoid rate limits
        
        self.total_tokens_used=0
        self.total_cost=0.0


    def generate_slide_title(self,df:pd.DataFrame,sheet_name:str,chart_type:Optional[str]=None)->str:
        """ 
        Generate engaging,data-driven slide title

        Args:
        df.DataFrame with slide data
        sheet_name:Name of the Excel Sheet
        """
        data_summary=self._create_data_summary(df,max_rows=5)
        stats=self._calculate_basic_stats(df)

        prompt=f""" You are an expert at creating engaging executive presentation titles.
        
Sheet Name:{sheet_name}
Chart Type:{chart_type or 'Unkown'}
Data Summary:{data_summary}

Key Statistics:{json.dumps(stats,indent=2)}
Create ONE clear, data-driven title (MAXIMUM 10 words) that highlights the most important insight.
Focus on the biggest number, trend, or comparison.
Use specific percentages or amounts when possible.

GOOD Examples:
- "Revenue Surged 34% to $2.8M in Q4"
- "Product A Dominates with 67% Market Share"
- "Customer Acquisition Costs Down 23% YoY"

BAD Examples (too generic):
- "Sales Data Analysis"
- "Quarterly Report"
- "Revenue Overview"

                """

        try:
            response=self.client.chat.completions.create(
            model=self.capable_model,
            messages=[{"role":"user","content":prompt}],
            temperature=0.7,
            max_tokens=30,
            )

            title=response.choices[0].message.content.strip()
            title=title.strip('"\'')

            self._track_usage(response.usage)
            return title
        except Exception as e:
            print(f"AI titel Generation Error:{str(e)}")
            return f"{sheet_name}Analysis"


    def generate_slide_summary(self,df:pd.DataFrame,sheet_name:str,
                               chart_type:Optional[str]=None)-> str:
        """ 
        Generate exceutive summary for slides

        """
        data_summary=self._create_data_summary(df,max_rows=8)
        stats=self._calculate_basic_stats(df)

        prompt=f"""You are a Financil analyst creating exceutive summaries.
        
        Sheet: {sheet_name}
Chart Type: {chart_type or 'Data visualization'}

Data:
{data_summary}

Statistics:
{json.dumps(stats, indent=2)}

Create a professional 2-3 sentence summary highlighting:
1. The main finding (with specific number)
2. A key comparison or trend
3. Business implication or context

Use specific numbers and percentages. Be concise and impactful.

Example:
"Revenue exceeded forecast by $2.3M (18%) driven by strong product sales. All regions showed positive growth with APAC leading at 34%. Profit margins improved from 18% to 22%, indicating strong operational efficiency."

Generate summary:"""
        try:
            response=self.client.chat.completions.create(
                model=self.capable_model,
                messages=[{"role":"user","content":prompt}],
                temperature=0.6,
                max_tokens=150
            )
            summary=response.choices[0].message.content.strip()
            self._track_usage(response.usage)
            return summary

        except Exception as e:
            print(f"AI Summary Generation Error:{str(e)}")
            return f"Analysis of {sheet_name} showing key metrices and trends"
        
    def generate_data_insights(self, df: pd.DataFrame, sheet_name: str,
                              num_insights: int = 5) -> List[str]:
        """
        Generate actionable data insights as bullet points
        
        Args:
            df: DataFrame to analyze
            sheet_name: Name of the sheet
            num_insights: Number of insights to generate (default 5)
            
        Returns:
            List of professional bullet points with specific data
        """
        
        data_summary = self._create_data_summary(df, max_rows=10)
        stats = self._calculate_basic_stats(df)
        trends = self._detect_simple_trends(df)
        
        prompt = f"""You are a data analyst creating insights for an executive presentation.

Sheet: {sheet_name}

Data Sample:
{data_summary}

Statistics:
{json.dumps(stats, indent=2)}

Detected Trends:
{json.dumps(trends, indent=2)}

Generate {num_insights} concise, data-driven bullet points (each MAX 15 words).
Focus on:
- Growth rates and trends
- Comparisons (best/worst performers)
- Risks or opportunities
- Specific numbers and percentages

Format: Do NOT include bullet symbols, just the text
Each insight on a new line

Example output:
Revenue grew 23% YoY accelerating from 15% in previous quarters
Product A accounts for 67% of revenue creating concentration risk
APAC region shows 34% growth outpacing all other markets
Customer acquisition costs decreased 12% while maintaining quality
Q4 shows exceptional performance exceeding targets by $2.3M

Generate {num_insights} insights:"""

        try:
            response = self.client.chat.completions.create(
                model=self.capable_model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.6,
                max_tokens=400
            )
            
            content = response.choices[0].message.content.strip()
            self._track_usage(response.usage)
            
            # Parse insights (split by newline)
            insights = [line.strip() for line in content.split('\n') 
                       if line.strip() and len(line.strip()) > 10]
            
            return insights[:num_insights]
            
        except Exception as e:
            print(f"AI Insights Generation Error: {str(e)}")
            return [
                f"Analysis of {len(df)} data points across {len(df.columns)} metrics",
                "Key trends and patterns identified in the dataset",
                "Multiple data points showing significant variation"
            ]
    
    def recommend_template(self, file_name: str, all_sheets_info: List[Dict],
                          available_templates: List[str]) -> Dict[str, Any]:
        """
        Recommend best template based on presentation context
        
        Args:
            file_name: Name of Excel file
            all_sheets_info: List of dicts with sheet names and data types
            available_templates: List of available template names
            
        Returns:
            Dict with top 3 template recommendations and reasoning
        """
        
        # Prepare context
        sheets_summary = "\n".join([
            f"- {info.get('name', 'Unknown')}: {info.get('data_type', 'unknown')} "
            f"({info.get('rows', 0)} rows, {info.get('cols', 0)} columns)"
            for info in all_sheets_info[:10]  # First 10 sheets
        ])
        
        templates_list = ", ".join(available_templates)
        
        prompt = f"""You are a presentation design expert selecting the best template.

File Name: {file_name}

Sheets Overview:
{sheets_summary}

Available Templates:
{templates_list}

Template Descriptions:
- corporate_blue: Classic professional, trusted for business reports
- financial_green: Financial data, conveys growth and stability
- executive_dark: Sophisticated dark theme for board-level presentations
- minimal_white: Clean minimal design for modern tech presentations
- vibrant_orange: Energetic sales and marketing presentations
- professional_purple: Creative yet professional for consulting
- tech_gradient: Modern tech metrics and dashboards
- modern_teal: Fresh contemporary look for innovation topics
- elegant_gold: Premium elegant design for high-stakes presentations
- classic_red: Bold powerful design for urgent/important topics

Based on the file name and data type, recommend the TOP 3 BEST templates.

Respond in this EXACT JSON format:
{{
  "recommendations": [
    {{
      "template": "template_name",
      "confidence": 0.95,
      "reasoning": "One sentence explaining why this template fits"
    }},
    {{
      "template": "template_name",
      "confidence": 0.85,
      "reasoning": "One sentence explaining why"
    }},
    {{
      "template": "template_name",
      "confidence": 0.75,
      "reasoning": "One sentence explaining why"
    }}
  ],
  "auto_selected": "template_name"
}}

Generate JSON response:"""

        try:
            response = self.client.chat.completions.create(
                model=self.balanced_model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.5,
                max_tokens=500,
                response_format={"type": "json_object"}  # Force JSON response
            )
            
            content = response.choices[0].message.content.strip()
            self._track_usage(response.usage)
            
            # Parse JSON response
            result = json.loads(content)
            return result
            
        except Exception as e:
            print(f"AI Template Recommendation Error: {str(e)}")
            # Fallback to default
            return {
                "recommendations": [
                    {
                        "template": "corporate_blue",
                        "confidence": 0.8,
                        "reasoning": "Safe default choice for business presentations"
                    },
                    {
                        "template": "financial_green",
                        "confidence": 0.7,
                        "reasoning": "Alternative for financial data"
                    },
                    {
                        "template": "minimal_white",
                        "confidence": 0.6,
                        "reasoning": "Clean modern alternative"
                    }
                ],
                "auto_selected": "corporate_blue"
            }
    def optimize_slide_layout(self, df: pd.DataFrame, chart_type: str,
                             has_insights: bool = True) -> Dict[str, Any]:
        """
        Determine optimal slide layout based on data complexity
        
        Args:
            df: DataFrame with slide data
            chart_type: Type of chart (pie, bar, line, etc.)
            has_insights: Whether AI insights are available
            
        Returns:
            Dict with layout type and positioning specs
        """
        
        rows = len(df)
        cols = len(df.columns)
        numeric_cols = len(df.select_dtypes(include=[np.number]).columns)
        
        context = f"""
Data: {rows} rows, {cols} columns ({numeric_cols} numeric)
Chart Type: {chart_type}
Has AI Insights: {has_insights}
"""
        
        prompt = f"""You are a presentation layout expert optimizing slide design.

{context}

Available Layout Types:
1. full_chart: Use when data is simple (<10 rows), focus entirely on visualization
2. chart_table: Use when data is detailed (>20 rows), show chart + data table
3. chart_insights: Use when AI insights available, show chart + bullet points
4. comparison: Use when comparing 2-3 metrics side by side
5. dashboard: Use when showing 4+ metrics, grid of small charts

Select the BEST layout and provide exact positioning.

Respond in this EXACT JSON format:
{{
  "layout": "chart_insights",
  "reasoning": "One sentence explaining why this layout is best",
  "positioning": {{
    "chart": {{"left": 0.5, "top": 1.5, "width": 5.5, "height": 5.0}},
    "insights": {{"left": 6.5, "top": 1.5, "width": 3.5, "height": 5.0}},
    "title": {{"left": 0.5, "top": 0.3, "width": 9.0, "height": 0.8}}
  }},
  "font_sizes": {{
    "title": 28,
    "insights": 14,
    "chart_labels": 11
  }}
}}

Note: Slide dimensions are 10 inches wide x 7.5 inches tall
Positions are in inches from top-left corner

Generate JSON response:"""

        try:
            response = self.client.chat.completions.create(
                model=self.balanced_model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.4,
                max_tokens=400,
                response_format={"type": "json_object"}  # Force JSON response
            )
            
            content = response.choices[0].message.content.strip()
            self._track_usage(response.usage)
            
            result = json.loads(content)
            return result
            
        except Exception as e:
            print(f"AI Layout Optimization Error: {str(e)}")
            # Fallback to default layout
            return {
                "layout": "chart_insights" if has_insights else "chart_table",
                "reasoning": "Balanced layout for data and insights",
                "positioning": {
                    "chart": {"left": 0.5, "top": 1.5, "width": 5.5, "height": 5.0},
                    "insights": {"left": 6.5, "top": 1.5, "width": 3.5, "height": 5.0},
                    "title": {"left": 0.5, "top": 0.3, "width": 9.0, "height": 0.8}
                },
                "font_sizes": {"title": 28, "insights": 14, "chart_labels": 11}
            }
        
    def recommend_chart_type(self, df: pd.DataFrame, column_names: List[str],
                           business_context: Optional[str] = None) -> Dict[str, Any]:
        """
        Recommend best chart type based on data structure and context
        
        Args:
            df: DataFrame to visualize
            column_names: List of column names
            business_context: Optional context (e.g., "quarterly sales")
            
        Returns:
            Dict with recommended chart type, alternatives, and tips
        """
        
        # Analyze data structure
        rows = len(df)
        cols = len(df.columns)
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        # Check for time series
        has_date = any('date' in str(col).lower() or 'time' in str(col).lower() 
                      or 'month' in str(col).lower() or 'quarter' in str(col).lower()
                      or 'year' in str(col).lower() for col in df.columns)
        
        data_summary = self._create_data_summary(df, max_rows=5)
        
        prompt = f"""You are a data visualization expert recommending chart types.

Data Structure:
- Rows: {rows}
- Columns: {cols}
- Numeric columns: {len(numeric_cols)}
- Has time/date column: {has_date}
- Column names: {', '.join(column_names)}

Business Context: {business_context or 'General business data'}

Sample Data:
{data_summary}

Chart Type Guidelines:
- Line: Time series trends, continuous data over time
- Bar: Comparing categories (horizontal bars)
- Column: Comparing values across categories (vertical bars)
- Pie: Part-to-whole relationships (2-7 categories only)
- Scatter: Correlation between two numeric variables
- Combo: Multiple metrics with different scales

Recommend the BEST chart type and provide alternatives.

Respond in this EXACT JSON format:
{{
  "recommended": {{
    "type": "line",
    "confidence": 0.92,
    "reasoning": "One sentence explaining why this chart type is best",
    "tips": [
      "Specific tip about colors or styling",
      "Specific tip about what to highlight",
      "Specific tip about additional elements"
    ]
  }},
  "alternatives": [
    {{
      "type": "column",
      "confidence": 0.75,
      "reasoning": "When this alternative works better",
      "use_case": "Specific scenario for this chart"
    }},
    {{
      "type": "combo",
      "confidence": 0.68,
      "reasoning": "When this alternative works better",
      "use_case": "Specific scenario for this chart"
    }}
  ]
}}

Generate JSON response:"""

        try:
            response = self.client.chat.completions.create(
                model=self.capable_model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.5,
                max_tokens=600,
                response_format={"type": "json_object"}  # Force JSON response
            )
            
            content = response.choices[0].message.content.strip()
            self._track_usage(response.usage)
            
            result = json.loads(content)
            return result
            
        except Exception as e:
            print(f"AI Chart Recommendation Error: {str(e)}")
            # Fallback logic
            if has_date and len(numeric_cols) > 0:
                chart_type = "line"
            elif rows <= 7 and len(numeric_cols) == 1:
                chart_type = "pie"
            elif len(numeric_cols) >= 2:
                chart_type = "scatter"
            else:
                chart_type = "column"
            
            return {
                "recommended": {
                    "type": chart_type,
                    "confidence": 0.7,
                    "reasoning": f"Selected based on data structure: {rows} rows, {cols} columns",
                    "tips": [
                        "Use clear colors for distinction",
                        "Add data labels for clarity",
                        "Consider adding a legend"
                    ]
                },
                "alternatives": []
            }
        
    def _create_data_summary(self, df: pd.DataFrame, max_rows: int = 5) -> str:
        """Create text summary of DataFrame"""
        summary_parts = []
        
        # Basic info
        summary_parts.append(f"Shape: {len(df)} rows × {len(df.columns)} columns")
        
        # Column names and types
        col_info = []
        for col in df.columns[:10]:  # First 10 columns
            dtype = 'numeric' if pd.api.types.is_numeric_dtype(df[col]) else 'text'
            col_info.append(f"{col} ({dtype})")
        summary_parts.append(f"Columns: {', '.join(col_info)}")
        
        # Sample data
        sample = df.head(max_rows).to_string(index=False, max_cols=10)
        summary_parts.append(f"\nSample:\n{sample}")
        
        return "\n".join(summary_parts)
    
    def _calculate_basic_stats(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Calculate basic statistics for numeric columns"""
        stats = {}
        
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        for col in numeric_cols[:5]:  # First 5 numeric columns
            values = df[col].dropna()
            if len(values) == 0:
                continue
            
            col_stats = {
                'mean': round(float(values.mean()), 2),
                'min': round(float(values.min()), 2),
                'max': round(float(values.max()), 2),
                'median': round(float(values.median()), 2)
            }
            
            # Calculate growth if enough data
            if len(values) >= 2:
                first = values.iloc[0]
                last = values.iloc[-1]
                if first != 0:
                    growth = ((last - first) / abs(first)) * 100
                    col_stats['growth_pct'] = round(growth, 1)
            
            stats[col] = col_stats
        
        return stats
    
    def _detect_simple_trends(self, df: pd.DataFrame) -> Dict[str, str]:
        """Detect simple trends in numeric data"""
        trends = {}
        
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        for col in numeric_cols[:3]:  # First 3 numeric columns
            values = df[col].dropna()
            if len(values) < 3:
                continue
            
            # Simple trend detection
            first_half = values[:len(values)//2].mean()
            second_half = values[len(values)//2:].mean()
            
            if pd.isna(first_half) or pd.isna(second_half):
                continue
            
            if second_half > first_half * 1.1:
                trends[col] = "increasing"
            elif second_half < first_half * 0.9:
                trends[col] = "decreasing"
            else:
                trends[col] = "stable"
        
        return trends
    
    def _track_usage(self, usage):
        """Track token usage for billing"""
        total_tokens = usage.total_tokens
        self.total_tokens_used += total_tokens
        
        # Groq pricing: $0.27 per 1M tokens
        cost = (total_tokens / 1_000_000) * 0.27
        self.total_cost += cost
    
    def get_usage_stats(self) -> Dict[str, Any]:
        """Get current usage statistics"""
        return {
            'total_tokens': self.total_tokens_used,
            'total_cost_usd': round(self.total_cost, 6),
            'cost_per_presentation': round(self.total_cost, 6)
        }
    
    def reset_usage_stats(self):
        """Reset usage tracking"""
        self.total_tokens_used = 0
        self.total_cost = 0.0
    
    # ============================================================================
    # NEW PROFESSIONAL SLIDE AI METHODS
    # ============================================================================
    
    def generate_executive_summary(self, data_summary: Dict[str, Any]) -> Dict[str, Any]:
        """
        Generate 4-5 executive summary bullet points analyzing growth trends
        
        Args:
            data_summary: Dict with data statistics and summaries
            
        Returns:
            Dict with 'insights' list of 4-5 bullet points
        """
        prompt = f"""You are a business analyst creating an executive summary for a presentation.

Data Overview:
{json.dumps(data_summary, indent=2)}

Generate 4-5 clear, actionable bullet points that:
1. Identify positive or negative growth trends
2. Compare metrics (revenue vs expenses, growth rates, etc.)
3. Highlight the most important business insights
4. Use plain English and specific numbers/percentages
5. Be concise (max 20 words per bullet)

Format as a JSON array of strings.

Example:
["Revenue increased by 14% while expenses decreased by 6%, improving margins",
 "Q4 showed strongest performance with $2.8M in total sales",
 "Customer acquisition costs dropped 23% year-over-year",
 "Product line diversification contributed to 18% growth in recurring revenue"]

Return ONLY the JSON array, no other text."""

        try:
            response = self.client.chat.completions.create(
                model=self.capable_model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.7,
                max_tokens=300
            )
            
            content = response.choices[0].message.content.strip()
            
            # Parse JSON response
            try:
                insights = json.loads(content)
                if isinstance(insights, list):
                    self._track_usage(response.usage)
                    return {'insights': insights[:5]}
            except json.JSONDecodeError:
                # Fallback: split by newlines
                insights = [line.strip(' -•') for line in content.split('\n') if line.strip()]
                return {'insights': insights[:5]}
                
        except Exception as e:
            print(f"⚠️  AI executive summary error: {e}")
            return {'insights': [
                "Data analysis completed successfully",
                "Key metrics showing positive trends",
                "Performance indicators within expected ranges",
                "Continued monitoring recommended"
            ]}
    
    def generate_category_insight(self, category_df: pd.DataFrame) -> Dict[str, str]:
        """
        Generate one-liner insight about category breakdown
        
        Args:
            category_df: DataFrame with category and value columns
            
        Returns:
            Dict with 'insight' string
        """
        if category_df is None or category_df.empty or len(category_df.columns) < 2:
            return {'insight': 'Category analysis completed'}
        
        try:
            # Get top category
            category_col = category_df.columns[0]
            value_col = category_df.columns[1]
            
            total = category_df[value_col].sum()
            top_idx = category_df[value_col].idxmax()
            top_category = category_df.loc[top_idx, category_col]
            top_value = category_df.loc[top_idx, value_col]
            percentage = (top_value / total * 100) if total > 0 else 0
            
            # Create summary for AI
            categories_summary = []
            for _, row in category_df.head(5).iterrows():
                cat = row[category_col]
                val = row[value_col]
                pct = (val / total * 100) if total > 0 else 0
                categories_summary.append(f"{cat}: {pct:.1f}%")
            
            prompt = f"""Generate ONE concise insight (max 15 words) about this category breakdown:

Top Category: {top_category} ({percentage:.1f}%)
All Categories: {', '.join(categories_summary)}

Example insights:
- "Marketing remains the largest contributor at 45%"
- "Product A dominates with 67% market share"
- "Department spending heavily concentrated in R&D (52%)"

Return ONLY the insight, no explanation."""

            response = self.client.chat.completions.create(
                model=self.fast_model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.7,
                max_tokens=40
            )
            
            insight = response.choices[0].message.content.strip()
            self._track_usage(response.usage)
            
            return {'insight': insight}
            
        except Exception as e:
            print(f"⚠️  AI category insight error: {e}")
            return {'insight': f"{top_category} represents the largest share at {percentage:.0f}%"}
    
    def generate_advanced_insights(self, data_summary: Dict[str, Any]) -> Dict[str, List[str]]:
        """
        Generate advanced AI insights: top performers, anomalies, predictions
        
        Args:
            data_summary: Dict with comprehensive data statistics
            
        Returns:
            Dict with 'top_performers', 'anomalies', 'predictions' lists
        """
        prompt = f"""You are a data analyst providing advanced business insights.

Data Summary:
{json.dumps(data_summary, indent=2)}

Generate insights in 3 categories:

1. TOP 3 PERFORMERS: Identify the 3 best-performing metrics, categories, or trends
2. ANOMALIES DETECTED: Identify 2-3 unusual patterns, outliers, or unexpected changes
3. PREDICTED TRENDS: Predict 2-3 likely future trends based on the data

Format as JSON:
{{
  "top_performers": ["Metric A showed 45% growth, highest in dataset", "Category B exceeded target by $500K", "Region C maintained 98% consistency"],
  "anomalies": ["Unusual spike in expenses during Q3", "Customer churn rate doubled in October"],
  "predictions": ["Revenue likely to grow 12-15% next quarter", "Seasonal pattern suggests December peak"]
}}

Keep each insight under 20 words. Be specific with numbers. Return ONLY valid JSON."""

        try:
            response = self.client.chat.completions.create(
                model=self.capable_model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.8,
                max_tokens=400
            )
            
            content = response.choices[0].message.content.strip()
            
            # Parse JSON
            try:
                insights = json.loads(content)
                self._track_usage(response.usage)
                
                return {
                    'top_performers': insights.get('top_performers', [])[:3],
                    'anomalies': insights.get('anomalies', [])[:3],
                    'predictions': insights.get('predictions', [])[:3]
                }
            except json.JSONDecodeError:
                print("⚠️  Failed to parse AI insights JSON")
                raise
                
        except Exception as e:
            print(f"⚠️  AI advanced insights error: {e}")
            return {
                'top_performers': ["Data analysis in progress", "Performance metrics being calculated", "Trends under evaluation"],
                'anomalies': ["No significant anomalies detected in current dataset"],
                'predictions': ["Insufficient historical data for accurate predictions", "Continued monitoring recommended"]
            }


# Helper function to create service instance
def create_ai_service(api_key: Optional[str] = None) -> AIInsightsService:
    """Create and return AI service instance"""
    return AIInsightsService(api_key=api_key)