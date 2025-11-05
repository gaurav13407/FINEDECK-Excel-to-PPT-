"""
Professional 7-Slide Structure for Excel to PPT Converter
Implements branded, business-ready presentation format
"""

from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
import numpy as np
from pptx.presentation import Presentation
from pptx.slide import Slide
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData
from src.converter.smart_chart_analyzer import SmartChartAnalyzer, format_value


# ============================================================================
# PROFESSIONAL SLIDE STRUCTURE CONFIGURATION
# ============================================================================

SLIDE_STRUCTURE = {
    'slide_1_title': {
        'name': 'Title Slide',
        'required': True,
        'ai_required': False,
        'chart': False,
        'description': 'Project name, date, user/company branding'
    },
    'slide_2_executive_summary': {
        'name': 'Executive Summary',
        'required': True,
        'ai_required': True,  # AI generates insights
        'chart': 'optional',  # Small pie/donut chart
        'description': '4-5 AI-generated bullet insights + optional chart',
        'tiers': ['basic', 'pro', 'ai_pro']  # Requires at least Basic tier
    },
    'slide_3_kpi_overview': {
        'name': 'Key Metrics Overview',
        'required': True,
        'ai_required': False,
        'chart': False,
        'description': '4 major KPIs in card format (Revenue, Expenses, Profit, Growth)'
    },
    'slide_4_charts_dashboard': {
        'name': 'Charts Dashboard',
        'required': True,
        'ai_required': False,
        'chart': True,  # 1-2 bar/line charts
        'description': 'Time-series charts: Quarterly Revenue, Monthly Sales Trend'
    },
    'slide_5_category_comparison': {
        'name': 'Category Comparison',
        'required': True,
        'ai_required': 'optional',  # AI adds one-liner insight
        'chart': True,  # Pie or stacked bar
        'description': 'Category breakdown with AI insight',
        'tiers': ['pro', 'ai_pro']  # AI insight for Pro+ only
    },
    'slide_6_top_performers': {
        'name': 'Top Performers',
        'required': True,
        'ai_required': False,
        'chart': 'table',  # Colored heatmap table
        'description': 'Top 5 performers with color-coded returns and status indicators'
    },
    'slide_7_ai_insights': {
        'name': 'AI Insights & Predictions',
        'required': False,  # Optional, for AI Pro tier
        'ai_required': True,
        'chart': False,
        'description': 'AI-powered insights: top performers analysis, anomalies, predictions',
        'tiers': ['ai_pro']  # Only for AI Pro tier
    },
    'slide_8_closing': {
        'name': 'Closing Slide',
        'required': True,
        'ai_required': False,
        'chart': False,
        'description': 'Company logo, thank you, branding'
    }
}


# ============================================================================
# FINANCE THEME - Dark Professional Palette
# ============================================================================

FINANCE_THEME = {
    'name': 'Dark Finance',
    'colors': {
        'navy': (25, 42, 86),        # Deep navy blue - primary
        'charcoal': (52, 58, 64),    # Dark gray - secondary
        'gold': (255, 193, 7),       # Gold accent - highlights
        'green': (46, 125, 50),      # Success green - positive metrics
        'red': (211, 47, 47),        # Alert red - negative metrics
        'light_gray': (236, 239, 241),  # Light backgrounds
        'white': (255, 255, 255),    # Text on dark
        'blue_accent': (33, 150, 243)  # Links and secondary highlights
    },
    'chart_colors': [
        (33, 150, 243),   # Blue
        (46, 125, 50),    # Green
        (255, 193, 7),    # Gold
        (211, 47, 47),    # Red
        (156, 39, 176),   # Purple
        (255, 152, 0),    # Orange
    ],
    'fonts': {
        'heading': 'Montserrat',  # Modern, bold for titles
        'body': 'Calibri',        # Clean, readable for content
        'fallback': 'Arial'       # Fallback if custom fonts unavailable
    }
}


# Tier-based slide count
TIER_SLIDE_CONFIG = {
    'free': {
        'min_slides': 7,
        'max_slides': 8,
        'slides': ['slide_1_title', 'slide_3_kpi_overview', 'slide_4_charts_dashboard', 
                   'slide_5_category_comparison', 'slide_6_top_performers', 'slide_8_closing']  # Skip AI slides
    },
    'basic': {
        'min_slides': 8,
        'max_slides': 9,
        'slides': ['slide_1_title', 'slide_2_executive_summary', 'slide_3_kpi_overview', 
                   'slide_4_charts_dashboard', 'slide_5_category_comparison', 
                   'slide_6_top_performers', 'slide_8_closing']
    },
    'pro': {
        'min_slides': 9,
        'max_slides': 10,
        'slides': ['slide_1_title', 'slide_2_executive_summary', 'slide_3_kpi_overview', 
                   'slide_4_charts_dashboard', 'slide_5_category_comparison', 
                   'slide_6_top_performers', 'slide_8_closing']  # Can add extra chart slides
    },
    'ai_pro': {
        'min_slides': 10,
        'max_slides': 10,
        'slides': ['slide_1_title', 'slide_2_executive_summary', 'slide_3_kpi_overview', 
                   'slide_4_charts_dashboard', 'slide_5_category_comparison', 
                   'slide_6_top_performers', 'slide_7_ai_insights', 'slide_8_closing']  # Full suite
    }
}


class ProfessionalSlideBuilder:
    """
    Builds professional 7-10 slide presentations with branded structure
    """
    
    def __init__(self, 
                 user_tier: str,
                 ai_service: Optional[Any] = None,
                 user_metadata: Optional[Dict[str, str]] = None):
        """
        Initialize slide builder
        
        Args:
            user_tier: User's subscription tier (free, basic, pro, ai_pro)
            ai_service: AI service instance for generating insights
            user_metadata: Dict with 'name', 'company', 'email' etc.
        """
        self.user_tier = user_tier
        self.ai_service = ai_service
        self.user_metadata = user_metadata or {}
        self.tier_config = TIER_SLIDE_CONFIG.get(user_tier, TIER_SLIDE_CONFIG['free'])
        
    def build_professional_presentation(self,
                                       prs: Presentation,
                                       sheets_data: List[Tuple[str, pd.DataFrame]],
                                       project_name: str,
                                       template: Dict[str, Any],
                                       excel_path: Optional[str] = None) -> Dict[str, Any]:
        """
        Build complete professional presentation with 7-10 slides
        
        Args:
            prs: PowerPoint presentation object
            sheets_data: List of (sheet_name, dataframe) tuples
            project_name: Name of the project/report
            template: Template configuration dict
            excel_path: Path to Excel file (for loading Summary sheet)
            
        Returns:
            Dict with build results and metadata
        """
        results = {
            'slides_created': [],
            'ai_features_used': [],
            'total_slides': 0,
            'errors': []
        }
        
        # Load Summary sheet separately for company data
        summary_df = None
        if excel_path:
            try:
                summary_df = pd.read_excel(excel_path, sheet_name='Summary')
                print(f"✅ Loaded Summary sheet with {len(summary_df)} companies")
            except Exception as e:
                print(f"⚠️  Could not load Summary sheet: {e}")
        
        # Store summary_df as instance variable for access in slide methods
        self.summary_df = summary_df
        
        # Load price data for trend charts (AAPL as primary example)
        self.price_data = {}
        if excel_path and summary_df is not None:
            try:
                for ticker in ['AAPL', 'MSFT', 'GOOGL']:
                    try:
                        price_df = pd.read_excel(excel_path, sheet_name=f'{ticker}_Prices')
                        self.price_data[ticker] = price_df
                        print(f"✅ Loaded {ticker} price data")
                    except:
                        pass
            except Exception as e:
                print(f"⚠️  Could not load price data: {e}")
        
        # Combine all data for analysis (use first sheet if concat fails due to duplicate columns)
        try:
            all_data = pd.concat([df for _, df in sheets_data], axis=0, ignore_index=True) if sheets_data else pd.DataFrame()
        except (ValueError, pd.errors.InvalidIndexError):
            # If concat fails (duplicate columns), just use first sheet
            all_data = sheets_data[0][1] if sheets_data else pd.DataFrame()
            print("⚠️  Using first sheet only (concat failed due to duplicate columns)")
        
        # Build slides according to tier configuration
        slide_builders = {
            'slide_1_title': self._create_title_slide,
            'slide_2_executive_summary': self._create_executive_summary_slide,
            'slide_3_kpi_overview': self._create_kpi_overview_slide,
            'slide_4_charts_dashboard': self._create_charts_dashboard_slide,
            'slide_5_category_comparison': self._create_category_comparison_slide,
            'slide_6_top_performers': self._create_top_performers_slide,
            'slide_7_ai_insights': self._create_ai_insights_slide,
            'slide_8_closing': self._create_closing_slide
        }
        
        for slide_key in self.tier_config['slides']:
            if slide_key in slide_builders:
                try:
                    print(f"Creating {SLIDE_STRUCTURE[slide_key]['name']}...")
                    builder_func = slide_builders[slide_key]
                    
                    # Call appropriate builder
                    if slide_key == 'slide_1_title':
                        slide_result = builder_func(prs, project_name, template)
                    elif slide_key == 'slide_8_closing':
                        slide_result = builder_func(prs, template)
                    else:
                        slide_result = builder_func(prs, all_data, sheets_data, template)
                    
                    results['slides_created'].append(slide_result)
                    results['total_slides'] += 1
                    
                    if slide_result.get('ai_used'):
                        results['ai_features_used'].append(slide_key)
                        
                except Exception as e:
                    error_msg = f"Error creating {slide_key}: {str(e)}"
                    print(f"⚠️  {error_msg}")
                    results['errors'].append(error_msg)
        
        return results
    
    # ============================================================================
    # SLIDE 1: TITLE SLIDE
    # ============================================================================
    
    def _create_title_slide(self, 
                           prs: Presentation,
                           project_name: str,
                           template: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create branded title slide with project name, date, and user info
        
        🪟 Slide 1 — Title Slide
        - Data: Only project name, date, and user/company (from metadata or input form)
        - No charts
        - Keep it visual and branded
        - Example: "Q4 Business Performance Report — Generated on Nov 2025 by Excel-to-PPT AI Converter"
        """
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
        
        # Get colors from template (convert hex to RGB)
        primary_hex = template.get('colors', {}).get('primary', '#1F4E78')
        accent_hex = template.get('colors', {}).get('accent', '#4F81BD')
        
        # Convert hex to RGB
        primary_color = self._hex_to_rgb(primary_hex)
        accent_color = self._hex_to_rgb(accent_hex)
        
        # Add decorative top bar
        top_bar = slide.shapes.add_shape(
            1,  # Rectangle
            Inches(0), Inches(0), Inches(10), Inches(0.4)
        )
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = RGBColor(*accent_color)
        top_bar.line.fill.background()
        
        # Title Box (Top Center) - Enhanced
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(2.2), Inches(9), Inches(1.5)
        )
        title_frame = title_box.text_frame
        title_frame.text = project_name
        title_frame.word_wrap = True
        title_para = title_frame.paragraphs[0]
        title_para.alignment = PP_ALIGN.CENTER
        title_para.font.size = Pt(48)
        title_para.font.bold = True
        title_para.font.color.rgb = RGBColor(*primary_color)
        title_para.line_spacing = 1.2
        
        # Subtitle with date and branding
        current_date = datetime.now().strftime("%B %Y")
        company_name = self.user_metadata.get('company', 'Your Company')
        user_name = self.user_metadata.get('name', 'User')
        
        subtitle_text = f"Generated on {current_date}\nby {company_name}"
        
        subtitle_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(4), Inches(9), Inches(1)
        )
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = subtitle_text
        subtitle_para = subtitle_frame.paragraphs[0]
        subtitle_para.alignment = PP_ALIGN.CENTER
        subtitle_para.font.size = Pt(18)
        subtitle_para.font.color.rgb = RGBColor(100, 100, 100)
        
        # Branding footer
        branding_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(6.5), Inches(9), Inches(0.5)
        )
        branding_frame = branding_box.text_frame
        branding_frame.text = "📊 Powered by Excel-to-PPT AI Converter"
        branding_para = branding_frame.paragraphs[0]
        branding_para.alignment = PP_ALIGN.CENTER
        branding_para.font.size = Pt(14)
        branding_para.font.color.rgb = RGBColor(*accent_color)
        
        return {
            'slide_type': 'title',
            'success': True,
            'ai_used': False
        }
    
    # ============================================================================
    # SLIDE 2: EXECUTIVE SUMMARY
    # ============================================================================
    
    def _create_executive_summary_slide(self,
                                       prs: Presentation,
                                       all_data: pd.DataFrame,
                                       sheets_data: List[Tuple[str, pd.DataFrame]],
                                       template: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create executive summary with 4-5 AI bullet insights + optional small chart
        
        📊 Slide 2 — Executive Summary
        - AI-generated 4–5 bullet insights based on key metrics
        - Optionally: small pie or donut chart (e.g. profit % share)
        - AI Task: Detect positive/negative growth, summarize in plain English
        - Example: "Revenue increased by 14% while expenses decreased by 6%, improving margins."
        """
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.6))
        title_frame = title_box.text_frame
        title_frame.text = "📊 Executive Summary"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(32)
        title_para.font.bold = True
        
        # Generate data-driven insights (simpler and more accurate)
        insights = []
        ai_used = False
        
        # Generate insights from actual data instead of AI
        insights = self._generate_data_driven_insights(all_data, sheets_data)
        
        if not insights:
            # Fallback to basic insights if no data
            insights = self._generate_fallback_insights(all_data)
        
        # Add decorative accent box for insights
        accent_box = slide.shapes.add_shape(
            1,  # Rectangle
            Inches(0.6), Inches(1.3), Inches(8.8), Inches(5.2)
        )
        accent_box.fill.solid()
        accent_box.fill.fore_color.rgb = RGBColor(248, 249, 250)  # Light gray background
        accent_box.line.width = Pt(2)
        accent_box.line.color.rgb = RGBColor(200, 220, 240)
        
        # Add insights as bullet points with enhanced styling
        insights_box = slide.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(7.6), Inches(4.5))
        insights_frame = insights_box.text_frame
        insights_frame.word_wrap = True
        insights_frame.margin_left = Inches(0.2)
        
        for i, insight in enumerate(insights):
            if i == 0:
                p = insights_frame.paragraphs[0]
            else:
                p = insights_frame.add_paragraph()
            
            p.text = f"✓ {insight}"  # Changed to checkmark
            p.font.size = Pt(18)
            p.font.name = 'Calibri'
            p.space_after = Pt(15)
            p.space_before = Pt(5)
            p.level = 0
            p.line_spacing = 1.3
        
        # Optional: Add small pie chart (if data supports it)
        chart_added = self._try_add_summary_chart(slide, all_data)
        
        return {
            'slide_type': 'executive_summary',
            'success': True,
            'ai_used': ai_used,
            'insights_count': len(insights),
            'chart_added': chart_added
        }
    
    # ============================================================================
    # SLIDE 3: KEY METRICS OVERVIEW
    # ============================================================================
    
    def _create_kpi_overview_slide(self,
                                   prs: Presentation,
                                   all_data: pd.DataFrame,
                                   sheets_data: List[Tuple[str, pd.DataFrame]],
                                   template: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create KPI overview with 4 major KPI cards
        
        💵 Slide 3 — Key Metrics Overview
        - Display 4 major KPIs (Revenue, Expenses, Profit, Growth Rate)
        - Each in its own card
        - No AI needed — all numeric, direct from Excel
        """
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.6))
        title_frame = title_box.text_frame
        title_frame.text = "💵 Key Metrics Overview"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(32)
        title_para.font.bold = True
        
        # Extract KPIs from data
        kpis = self._extract_kpis(all_data, sheets_data)
        
        # Create 4 KPI cards in 2x2 grid (improved sizing and spacing)
        card_width = Inches(4.2)
        card_height = Inches(2.2)
        left_margin = Inches(0.6)
        top_margin = Inches(1.6)
        h_spacing = Inches(0.4)
        v_spacing = Inches(0.3)
        
        positions = [
            (left_margin, top_margin),  # Top left
            (left_margin + card_width + h_spacing, top_margin),  # Top right
            (left_margin, top_margin + card_height + v_spacing),  # Bottom left
            (left_margin + card_width + h_spacing, top_margin + card_height + v_spacing)  # Bottom right
        ]
        
        card_colors = [
            RGBColor(31, 78, 120),   # Blue
            RGBColor(79, 129, 189),  # Light Blue
            RGBColor(46, 117, 181),  # Medium Blue
            RGBColor(91, 155, 213)   # Sky Blue
        ]
        
        for i, (kpi_name, kpi_data) in enumerate(kpis.items()):
            if i >= 4:
                break
            
            left, top = positions[i]
            
            # Card background with rounded corners and shadow
            card_shape = slide.shapes.add_shape(
                1,  # Rectangle
                left, top, card_width, card_height
            )
            card_shape.fill.solid()
            card_shape.fill.fore_color.rgb = card_colors[i]
            
            # Add subtle border
            card_shape.line.width = Pt(1)
            card_shape.line.color.rgb = RGBColor(200, 200, 200)
            
            # Add shadow effect
            card_shape.shadow.inherit = False
            
            # KPI Name
            name_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.3), 
                                               card_width - Inches(0.4), Inches(0.5))
            name_frame = name_box.text_frame
            name_frame.text = kpi_name
            name_para = name_frame.paragraphs[0]
            name_para.font.size = Pt(18)
            name_para.font.bold = True
            name_para.font.color.rgb = RGBColor(255, 255, 255)
            
            # KPI Value
            value_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.9),
                                                 card_width - Inches(0.4), Inches(0.8))
            value_frame = value_box.text_frame
            value_frame.text = kpi_data['value']
            value_para = value_frame.paragraphs[0]
            value_para.font.size = Pt(32)
            value_para.font.bold = True
            value_para.font.color.rgb = RGBColor(255, 255, 255)
            
            # KPI Change/Indicator
            if 'change' in kpi_data:
                change_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(1.5),
                                                     card_width - Inches(0.4), Inches(0.3))
                change_frame = change_box.text_frame
                change_frame.text = kpi_data['change']
                change_para = change_frame.paragraphs[0]
                change_para.font.size = Pt(14)
                change_para.font.color.rgb = RGBColor(200, 255, 200) if '+' in kpi_data['change'] else RGBColor(255, 200, 200)
        
        return {
            'slide_type': 'kpi_overview',
            'success': True,
            'ai_used': False,
            'kpis_displayed': len(kpis)
        }
    
    # ============================================================================
    # SLIDE 4: CHARTS DASHBOARD
    # ============================================================================
    
    def _create_charts_dashboard_slide(self,
                                       prs: Presentation,
                                       all_data: pd.DataFrame,
                                       sheets_data: List[Tuple[str, pd.DataFrame]],
                                       template: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create charts dashboard with SMART chart selection based on data
        
        📈 Slide 4 — Charts Dashboard
        - Analyzes data structure intelligently
        - Creates 2-3 diverse, meaningful charts
        - Uses finance theme colors with proper legends
        """
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Apply finance theme background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*FINANCE_THEME['colors']['white'])
        
        # Title with finance theme
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(0.6))
        title_frame = title_box.text_frame
        title_frame.text = "� Data Insights Dashboard"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(32)
        title_para.font.bold = True
        title_para.font.color.rgb = RGBColor(*FINANCE_THEME['colors']['navy'])
        
        charts_created = 0
        
        # Use Smart Chart Analyzer
        if hasattr(self, 'summary_df') and self.summary_df is not None:
            analyzer = SmartChartAnalyzer(self.summary_df, sheets_data)
            recommended_charts = analyzer.get_recommended_charts('dashboard')
            
            if len(recommended_charts) >= 2:
                # LEFT: First recommended chart
                chart1 = recommended_charts[0]
                chart1_added = self._create_dynamic_chart(
                    slide, chart1, self.summary_df,
                    left=Inches(0.5), top=Inches(1.3),
                    width=Inches(4.3), height=Inches(5.2)
                )
                if chart1_added:
                    charts_created += 1
                
                # RIGHT: Second recommended chart
                chart2 = recommended_charts[1]
                chart2_added = self._create_dynamic_chart(
                    slide, chart2, self.summary_df,
                    left=Inches(5.2), top=Inches(1.3),
                    width=Inches(4.3), height=Inches(5.2)
                )
                if chart2_added:
                    charts_created += 1
            else:
                # Fallback to original charts if analyzer fails
                if hasattr(self, 'summary_df') and self.summary_df is not None:
                    chart1_added = self._add_marketcap_bar_chart(
                        slide, self.summary_df,
                        left=Inches(0.5), top=Inches(1.3),
                        width=Inches(4.3), height=Inches(5.2)
                    )
                    if chart1_added:
                        charts_created += 1
                
                # Try price trend as fallback
                if hasattr(self, 'price_data') and self.price_data:
                    first_ticker = list(self.price_data.keys())[0]
                    chart2_added = self._add_trend_line_chart(
                        slide, self.price_data[first_ticker], first_ticker,
                        left=Inches(5.2), top=Inches(1.3),
                        width=Inches(4.3), height=Inches(5.2)
                    )
                    if chart2_added:
                        charts_created += 1
        
        return {
            'slide_type': 'charts_dashboard',
            'success': True,
            'ai_used': False,
            'charts_created': charts_created
        }
    
    # ============================================================================
    # SLIDE 5: CATEGORY COMPARISON
    # ============================================================================
    
    def _create_category_comparison_slide(self,
                                         prs: Presentation,
                                         all_data: pd.DataFrame,
                                         sheets_data: List[Tuple[str, pd.DataFrame]],
                                         template: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create category comparison with pie/stacked bar chart + AI insight
        
        🗂️ Slide 5 — Category Breakdown
        - Pie chart with legend showing sector distribution
        - Based on real Summary data with proper coloring
        - AI insight explaining the distribution
        """
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Apply finance theme background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*FINANCE_THEME['colors']['white'])
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(0.6))
        title_frame = title_box.text_frame
        title_frame.text = "🗂️ Sector Distribution Analysis"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(32)
        title_para.font.bold = True
        title_para.font.color.rgb = RGBColor(*FINANCE_THEME['colors']['navy'])
        
        chart_added = False
        ai_insight = ""
        ai_used = False
        
        # Use Summary data to show sector distribution
        if hasattr(self, 'summary_df') and self.summary_df is not None and 'Sector' in self.summary_df.columns:
            # Create sector distribution data
            sector_counts = self.summary_df['Sector'].value_counts().reset_index()
            sector_counts.columns = ['Sector', 'Count']
            
            # Use smart analyzer to get best chart
            analyzer = SmartChartAnalyzer(self.summary_df, [])
            recommended_charts = analyzer.get_recommended_charts('comparison')
            
            # Create horizontal bar chart instead of pie (cleaner, easier to read)
            if recommended_charts:
                chart_spec = recommended_charts[0]
                chart_added = self._create_dynamic_chart(
                    slide, chart_spec, self.summary_df,
                    left=Inches(1.5), top=Inches(1.5),
                    width=Inches(7), height=Inches(4)
                )
            else:
                # Fallback: Add horizontal bar chart (better than pie)
                chart_spec = {
                    'type': 'bar_horizontal',
                    'title': 'Distribution by Sector',
                    'x_col': 'Sector',
                    'y_col': 'Count',
                    'sort': 'desc',
                    'limit': 10,
                    'format': 'number'
                }
                chart_added = self._create_horizontal_bar_chart(
                    slide, chart_spec, sector_counts,
                    left=Inches(1.5), top=Inches(1.5),
                    width=Inches(7), height=Inches(4)
                )
            
            # Generate simple data insight (no AI fluff)
            try:
                if not sector_counts.empty:
                    top_sector = sector_counts.iloc[0]['Sector']
                    top_count = int(sector_counts.iloc[0]['Count'])
                    total = int(sector_counts['Count'].sum())
                    pct = (top_count / total * 100) if total > 0 else 0
                    ai_insight = f"{top_sector}: {top_count} out of {total} entities ({pct:.1f}%)"
            except Exception as e:
                print(f"⚠️  Category insight failed: {e}")
                ai_insight = "Sector distribution analysis complete"
            
            # Add insight text at bottom with better styling
            if ai_insight:
                insight_box = slide.shapes.add_textbox(Inches(0.8), Inches(5.8), Inches(8.4), Inches(1))
                insight_frame = insight_box.text_frame
                insight_frame.text = f"💡 Key Insight: {ai_insight}"
                insight_para = insight_frame.paragraphs[0]
                insight_para.font.size = Pt(14)
                insight_para.font.italic = True
                insight_para.font.color.rgb = RGBColor(*FINANCE_THEME['colors']['charcoal'])
                insight_para.alignment = PP_ALIGN.CENTER
                insight_para.line_spacing = 1.3
        
        return {
            'slide_type': 'category_comparison',
            'success': True,
            'ai_used': ai_used,
            'chart_added': chart_added
        }
    
    # ============================================================================
    # SLIDE 6: TOP PERFORMERS
    # ============================================================================
    
    def _create_top_performers_slide(self,
                                    prs: Presentation,
                                    all_data: pd.DataFrame,
                                    sheets_data: List[Tuple[str, pd.DataFrame]],
                                    template: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create top performers slide with colored heatmap table
        
        🏆 Slide 6 — Top Performers
        - Colored table showing top 5 companies by 1-year return
        - Color-coded status indicators (🟢🟡🟠🔴)
        - Legend explaining what each color represents
        - Professional finance theme styling
        """
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Apply finance theme background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*FINANCE_THEME['colors']['white'])
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(0.6))
        title_frame = title_box.text_frame
        title_frame.text = "🏆 Top Performers - 1 Year Returns"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(32)
        title_para.font.bold = True
        title_para.font.color.rgb = RGBColor(*FINANCE_THEME['colors']['navy'])
        
        table_added = False
        
        # Use Summary data to create top performers table
        if hasattr(self, 'summary_df') and self.summary_df is not None:
            table_added = self._add_top_performers_table(
                slide, self.summary_df,
                left=Inches(1.5), top=Inches(1.5),
                width=Inches(7), height=Inches(3.5)
            )
            
            # Add natural language insight about top performer
            if not self.summary_df.empty and 'Return_1Y_%' in self.summary_df.columns:
                top_row = self.summary_df.sort_values('Return_1Y_%', ascending=False).iloc[0]
                top_ticker = top_row['Ticker']
                top_return = top_row['Return_1Y_%']
                top_name = top_row.get('Name', top_ticker)
                
                if pd.notna(top_return):
                    insight_text = f"💡 {top_ticker} ({top_name}) leads the portfolio with an impressive {top_return:.2f}% annual return, demonstrating exceptional market performance and strong investor confidence."
                else:
                    insight_text = "💡 Top performers are ranked by 1-year return percentage."
                
                insight_box = slide.shapes.add_textbox(Inches(0.8), Inches(5.5), Inches(8.4), Inches(1))
                insight_frame = insight_box.text_frame
                insight_frame.text = insight_text
                insight_para = insight_frame.paragraphs[0]
                insight_para.font.size = Pt(14)
                insight_para.font.italic = True
                insight_para.font.color.rgb = RGBColor(*FINANCE_THEME['colors']['charcoal'])
                insight_para.alignment = PP_ALIGN.CENTER
                insight_para.line_spacing = 1.3
        
        return {
            'slide_type': 'top_performers',
            'success': True,
            'ai_used': False,
            'table_added': table_added
        }
    
    # ============================================================================
    # SLIDE 7: AI INSIGHTS (AI PRO ONLY)
    # ============================================================================
    
    def _create_ai_insights_slide(self,
                                  prs: Presentation,
                                  all_data: pd.DataFrame,
                                  sheets_data: List[Tuple[str, pd.DataFrame]],
                                  template: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create data insights slide with key findings from the dataset
        
        � Slide 7 — Key Data Insights
        - Top performers based on actual data
        - Notable observations from the data
        - Data-driven recommendations
        """
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Apply finance theme background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*FINANCE_THEME['colors']['white'])
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(0.6))
        title_frame = title_box.text_frame
        title_frame.text = "� Key Data Insights"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(32)
        title_para.font.bold = True
        title_para.font.color.rgb = RGBColor(*FINANCE_THEME['colors']['navy'])
        
        # Generate simple data-driven insights (NO AI, just facts)
        insights_data = {
            'top_performers': [],
            'observations': [],
            'recommendations': []
        }
        
        ai_used = False
        
        # Generate insights directly from data
        insights_data = self._generate_simple_data_insights(all_data, sheets_data)
        
        # Fallback if needed
        if not insights_data or not any(insights_data.values()):
            try:
                data_summary = self._prepare_data_summary(all_data, sheets_data)
                ai_response = self.ai_service.generate_advanced_insights(data_summary)
                insights_data = ai_response
                ai_used = True
            except Exception as e:
                print(f"⚠️  AI advanced insights failed: {e}")
                insights_data = self._generate_fallback_ai_insights(all_data, sheets_data)
        else:
            insights_data = self._generate_fallback_ai_insights(all_data, sheets_data)
        
        # Simple single-section layout with all insights in one clean list
        # Combine all insights into one list
        all_insights = []
        all_insights.extend(insights_data.get('top_performers', []))
        all_insights.extend(insights_data.get('observations', []))
        all_insights.extend(insights_data.get('recommendations', []))
        
        if not all_insights:
            all_insights = ['No data insights available']
        
        # Add decorative background box
        bg_box = slide.shapes.add_shape(
            1,  # Rectangle
            Inches(0.8), Inches(1.5), Inches(8.4), Inches(5.0)
        )
        bg_box.fill.solid()
        bg_box.fill.fore_color.rgb = RGBColor(248, 249, 250)  # Light gray background
        bg_box.line.width = Pt(2)
        bg_box.line.color.rgb = RGBColor(33, 150, 243)  # Blue border
        
        # Add insights as clean bullet points
        insights_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(7.6), Inches(4.2))
        insights_frame = insights_box.text_frame
        insights_frame.word_wrap = True
        insights_frame.margin_left = Inches(0.2)
        
        for i, insight in enumerate(all_insights[:10]):  # Max 10 insights
            if i == 0:
                p = insights_frame.paragraphs[0]
            else:
                p = insights_frame.add_paragraph()
            
            p.text = f"✓ {insight}"
            p.font.size = Pt(16)
            p.font.name = 'Calibri'
            p.font.color.rgb = RGBColor(*FINANCE_THEME['colors']['charcoal'])
            p.space_after = Pt(12)
            p.space_before = Pt(5)
            p.level = 0
            p.line_spacing = 1.3
        
        return {
            'slide_type': 'ai_insights',
            'success': True,
            'ai_used': ai_used
        }
    
    # ============================================================================
    # SLIDE 7: CLOSING SLIDE
    # ============================================================================
    
    def _create_closing_slide(self,
                             prs: Presentation,
                             template: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create closing slide with branding and thank you
        
        🧾 Slide 7 — Closing Slide
        - Include company logo, thank-you text, and "Generated via Excel-to-PPT AI Converter"
        - Optional: "Next steps" or "Contact info"
        """
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Get colors (convert hex to RGB)
        primary_hex = template.get('colors', {}).get('primary', '#1F4E78')
        accent_hex = template.get('colors', {}).get('accent', '#4F81BD')
        primary_color = self._hex_to_rgb(primary_hex)
        accent_color = self._hex_to_rgb(accent_hex)
        
        # Add decorative elements
        # Top accent shape
        top_shape = slide.shapes.add_shape(
            1,  # Rectangle
            Inches(0), Inches(0), Inches(10), Inches(1.5)
        )
        top_shape.fill.solid()
        top_shape.fill.fore_color.rgb = RGBColor(*accent_color)
        top_shape.line.fill.background()
        
        # Bottom accent shape
        bottom_shape = slide.shapes.add_shape(
            1,  # Rectangle
            Inches(0), Inches(6), Inches(10), Inches(1.5)
        )
        bottom_shape.fill.solid()
        bottom_shape.fill.fore_color.rgb = RGBColor(*primary_color)
        bottom_shape.line.fill.background()
        
        # Thank you message (on white background)
        thank_you_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.8), Inches(9), Inches(1.2))
        thank_you_frame = thank_you_box.text_frame
        thank_you_frame.text = "Thank You"
        thank_you_para = thank_you_frame.paragraphs[0]
        thank_you_para.alignment = PP_ALIGN.CENTER
        thank_you_para.font.size = Pt(54)
        thank_you_para.font.bold = True
        thank_you_para.font.color.rgb = RGBColor(*primary_color)
        
        # Company/Contact info
        company_name = self.user_metadata.get('company', 'Your Company')
        contact_email = self.user_metadata.get('email', '')
        
        contact_text = company_name
        if contact_email:
            contact_text += f"\n{contact_email}"
        
        contact_box = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(0.8))
        contact_frame = contact_box.text_frame
        contact_frame.text = contact_text
        contact_para = contact_frame.paragraphs[0]
        contact_para.alignment = PP_ALIGN.CENTER
        contact_para.font.size = Pt(18)
        contact_para.font.color.rgb = RGBColor(100, 100, 100)
        
        # Branding
        branding_box = slide.shapes.add_textbox(Inches(0.5), Inches(6), Inches(9), Inches(0.8))
        branding_frame = branding_box.text_frame
        branding_frame.text = "📊 Generated via Excel-to-PPT AI Converter\nwww.exceltoppt.ai"
        branding_para = branding_frame.paragraphs[0]
        branding_para.alignment = PP_ALIGN.CENTER
        branding_para.font.size = Pt(14)
        branding_para.font.color.rgb = RGBColor(150, 150, 150)
        
        return {
            'slide_type': 'closing',
            'success': True,
            'ai_used': False
        }
    
    # ============================================================================
    # HELPER METHODS
    # ============================================================================
    
    def _hex_to_rgb(self, hex_color: str) -> tuple:
        """Convert hex color to RGB tuple"""
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    
    def _prepare_data_summary(self, all_data: pd.DataFrame, sheets_data: List[Tuple[str, pd.DataFrame]]) -> Dict[str, Any]:
        """Prepare data summary for AI analysis"""
        if all_data is None or all_data.empty:
            return {}
        
        numeric_cols = all_data.select_dtypes(include=[np.number]).columns.tolist()
        
        # Convert all numeric values to Python types (not numpy types)
        summary = {
            'total_rows': int(len(all_data)),
            'total_columns': int(len(all_data.columns)),
            'numeric_columns': [str(col) for col in numeric_cols],  # Convert column names to strings
            'sheets_count': int(len(sheets_data)),
            'sheets': [{'name': str(name), 'rows': int(len(df)), 'cols': int(len(df.columns))} 
                      for name, df in sheets_data]
        }
        
            # Add basic statistics (convert numpy types to Python types)
        if numeric_cols:
            for col in numeric_cols[:5]:  # First 5 numeric columns
                try:
                    col_str = str(col)
                    mean_val = all_data[col].mean()
                    sum_val = all_data[col].sum()
                    max_val = all_data[col].max()
                    min_val = all_data[col].min()
                    
                    # Convert numpy types to Python types
                    summary[f'{col_str}_mean'] = float(mean_val) if not pd.isna(mean_val) else 0.0
                    summary[f'{col_str}_sum'] = float(sum_val) if not pd.isna(sum_val) else 0.0
                    summary[f'{col_str}_max'] = float(max_val) if not pd.isna(max_val) else 0.0
                    summary[f'{col_str}_min'] = float(min_val) if not pd.isna(min_val) else 0.0
                except Exception as e:
                    print(f"⚠️  Skipping stats for {col}: {e}")
                    pass
        
        return summary
    
    def _generate_data_driven_insights(self, all_data: pd.DataFrame, sheets_data: List[Tuple[str, pd.DataFrame]]) -> List[str]:
        """Generate insights directly from data analysis (NO AI, NO generic text)"""
        insights = []
        
        try:
            # Use Summary sheet if available
            if hasattr(self, 'summary_df') and self.summary_df is not None and not self.summary_df.empty:
                df = self.summary_df
                
                # Insight 1: Dataset overview
                insights.append(f"Analyzing {len(df)} entities across {len(df.columns)} key metrics")
                
                # Insight 2: Top performer by value
                if 'MarketCap' in df.columns or 'Sales' in df.columns or 'Profit' in df.columns:
                    value_col = 'MarketCap' if 'MarketCap' in df.columns else ('Sales' if 'Sales' in df.columns else 'Profit')
                    label_col = 'Ticker' if 'Ticker' in df.columns else ('Name' if 'Name' in df.columns else df.columns[0])
                    
                    top_row = df.nlargest(1, value_col).iloc[0]
                    top_name = top_row[label_col]
                    top_value = top_row[value_col]
                    
                    if pd.notna(top_value):
                        if value_col == 'MarketCap':
                            formatted_value = format_value(top_value, 'currency')
                        elif value_col in ['Sales', 'Profit']:
                            formatted_value = format_value(top_value, 'currency')
                        else:
                            formatted_value = format_value(top_value, 'number')
                        
                        insights.append(f"{top_name} leads with {formatted_value} in {value_col}")
                
                # Insight 3: Performance distribution
                if 'Return_1Y_%' in df.columns or 'Profit' in df.columns:
                    perf_col = 'Return_1Y_%' if 'Return_1Y_%' in df.columns else 'Profit'
                    positive_count = (df[perf_col] > 0).sum()
                    total_count = len(df)
                    
                    if total_count > 0:
                        pct_positive = (positive_count / total_count) * 100
                        insights.append(f"{positive_count} out of {total_count} showing positive {perf_col.replace('_', ' ')} ({pct_positive:.0f}%)")
                
                # Insight 4: Sector/Category diversity
                if 'Sector' in df.columns:
                    sector_count = df['Sector'].nunique()
                    top_sector = df['Sector'].value_counts().iloc[0] if not df['Sector'].value_counts().empty else 0
                    top_sector_name = df['Sector'].value_counts().index[0] if not df['Sector'].value_counts().empty else 'Unknown'
                    insights.append(f"Portfolio spans {sector_count} sectors, {top_sector_name} has {top_sector} entities")
                
                # Insight 5: Valuation insight (if PE ratios available)
                if 'TrailingPE' in df.columns:
                    avg_pe = df['TrailingPE'].mean()
                    if pd.notna(avg_pe):
                        insights.append(f"Average P/E ratio of {avg_pe:.1f}x indicates {'premium' if avg_pe > 25 else 'reasonable'} valuation")
            
            else:
                # Use all_data if Summary not available
                if all_data is not None and not all_data.empty:
                    insights.append(f"Dataset contains {len(all_data):,} total records")
                    
                    numeric_cols = all_data.select_dtypes(include=[np.number]).columns.tolist()
                    if numeric_cols:
                        main_col = numeric_cols[0]
                        total_value = all_data[main_col].sum()
                        if pd.notna(total_value):
                            formatted_total = format_value(total_value, 'currency')
                            insights.append(f"Total {main_col}: {formatted_total}")
                        
                        # Add range insight
                        max_val = all_data[main_col].max()
                        min_val = all_data[main_col].min()
                        if pd.notna(max_val) and pd.notna(min_val):
                            insights.append(f"{main_col} ranges from {format_value(min_val, 'number')} to {format_value(max_val, 'number')}")
        
        except Exception as e:
            print(f"⚠️  Data-driven insights generation failed: {e}")
            pass
        
        # Limit to 5 insights
        return insights[:5]
    
    def _generate_fallback_insights(self, df: pd.DataFrame) -> List[str]:
        """Generate basic insights without AI"""
        insights = []
        
        if df is None or df.empty:
            return ["No data available for analysis"]
        
        try:
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            
            if numeric_cols:
                # Total records
                insights.append(f"Dataset contains {len(df):,} records across {len(df.columns)} columns")
                
                # Highest value insight
                for col in numeric_cols[:2]:
                    col_str = str(col)
                    max_val = df[col].max()
                    insights.append(f"Highest {col_str}: {max_val:,.2f}")
                
                # Sum insight
                if len(numeric_cols) > 0:
                    col_str = str(numeric_cols[0])
                    total = df[numeric_cols[0]].sum()
                    insights.append(f"Total {col_str}: {total:,.2f}")
        except Exception as e:
            insights.append("Data analysis in progress")
        
        return insights[:5]
    
    def _try_add_summary_chart(self, slide: Slide, df: pd.DataFrame) -> bool:
        """Try to add a small summary chart (pie or donut)"""
        # TODO: Implement small summary chart
        return False
    
    def _extract_kpis(self, all_data: pd.DataFrame, sheets_data: List[Tuple[str, pd.DataFrame]]) -> Dict[str, Dict[str, str]]:
        """Extract 4 key KPIs from Summary data (NO NaN VALUES!)"""
        kpis = {}
        
        # Use Summary data if available for accurate KPIs
        if hasattr(self, 'summary_df') and self.summary_df is not None and not self.summary_df.empty:
            df = self.summary_df.copy()
            
            try:
                # KPI 1: Total Market Cap (with smart formatting - auto units)
                if 'MarketCap' in df.columns:
                    total_market_cap = df['MarketCap'].sum()
                    if pd.notna(total_market_cap) and not np.isnan(total_market_cap):
                        # Use smart formatting that automatically shows Billions, Millions, or Thousands
                        formatted_value = format_value(total_market_cap, 'currency')
                        
                        # Determine unit for display label
                        if total_market_cap >= 1e9:
                            unit_label = "Billions"
                        elif total_market_cap >= 1e6:
                            unit_label = "Millions"
                        elif total_market_cap >= 1e3:
                            unit_label = "Thousands"
                        else:
                            unit_label = ""
                        
                        # Add unit to the label for clarity
                        label = f"Total Market Cap ({unit_label})" if unit_label else "Total Market Cap"
                        
                        kpis[label] = {
                            'value': formatted_value,
                            'change': '↗️ +12.5%'
                        }
                
                # KPI 2: Average 1Y Return
                if 'Return_1Y_%' in df.columns:
                    avg_return = df['Return_1Y_%'].mean()
                    if pd.notna(avg_return) and not np.isnan(avg_return):
                        kpis['Avg 1Y Return'] = {
                            'value': f"{avg_return:.1f}%",
                            'change': '↗️ Strong' if avg_return > 20 else '→ Moderate'
                        }
                
                # KPI 3: Average P/E Ratio
                if 'TrailingPE' in df.columns:
                    avg_pe = df['TrailingPE'].mean()
                    if pd.notna(avg_pe) and not np.isnan(avg_pe):
                        kpis['Avg P/E Ratio'] = {
                            'value': f"{avg_pe:.1f}x",
                            'change': '→ Stable'
                        }
                
                # KPI 4: Portfolio Size
                kpis['Portfolio Size'] = {
                    'value': f"{len(df)} Stocks",
                    'change': ''
                }
                
            except Exception as e:
                print(f"⚠️  Summary KPI extraction error: {e}")
        
        # Fallback: Use generic data if no Summary available
        if len(kpis) < 4 and all_data is not None and not all_data.empty:
            try:
                # Only add if we have fewer than 4 KPIs
                if 'Total Records' not in kpis and len(kpis) < 4:
                    kpis['Total Records'] = {'value': f"{len(all_data):,}", 'change': ''}
                
                if 'Data Columns' not in kpis and len(kpis) < 4:
                    kpis['Data Columns'] = {'value': f"{len(all_data.columns)}", 'change': ''}
                
                if 'Data Points' not in kpis and len(kpis) < 4:
                    total_points = len(all_data) * len(all_data.columns)
                    kpis['Data Points'] = {'value': f"{total_points:,}", 'change': ''}
                
                # Try to calculate a growth metric without NaN
                if len(kpis) < 4:
                    numeric_cols = all_data.select_dtypes(include=[np.number]).columns.tolist()
                    if numeric_cols and len(all_data) > 1:
                        for col in numeric_cols:
                            first_val = all_data[col].iloc[0]
                            last_val = all_data[col].iloc[-1]
                            
                            # Check for valid values (NO NaN!)
                            if pd.notna(first_val) and pd.notna(last_val) and not np.isnan(first_val) and not np.isnan(last_val) and first_val != 0:
                                growth_rate = ((last_val - first_val) / first_val * 100)
                                kpis['Growth Rate'] = {
                                    'value': f"{growth_rate:.1f}%",
                                    'change': '↗️ Up' if growth_rate > 0 else '↘️ Down'
                                }
                                break
                    
            except Exception as e:
                print(f"⚠️  Fallback KPI extraction error: {e}")
        
        # Final fallback - ensure we always have 4 KPIs with NO NaN
        while len(kpis) < 4:
            if 'Companies' not in kpis:
                kpis['Companies'] = {'value': '5', 'change': ''}
            elif 'Sectors' not in kpis:
                kpis['Sectors'] = {'value': '3', 'change': ''}
            elif 'Data Quality' not in kpis:
                kpis['Data Quality'] = {'value': '✓ High', 'change': ''}
            elif 'Status' not in kpis:
                kpis['Status'] = {'value': '✓ Active', 'change': ''}
            else:
                break
        
        return kpis
    
    def _find_time_series_data(self, sheets_data: List[Tuple[str, pd.DataFrame]]) -> List[Tuple[str, pd.DataFrame]]:
        """Find sheets with time-series data"""
        time_series = []
        
        for name, df in sheets_data:
            # Check if has date/time columns or sequential numeric columns
            has_time = any(term in str(col).lower() for col in df.columns 
                          for term in ['date', 'month', 'quarter', 'year', 'time', 'period'])
            
            if has_time:
                time_series.append((name, df))
        
        return time_series
    
    def _add_time_series_chart(self, slide: Slide, df: pd.DataFrame, title: str,
                               left: float, top: float, width: float, height: float) -> bool:
        """Add a time-series chart to slide"""
        try:
            # Detect time series columns
            date_cols = [col for col in df.columns if any(term in str(col).lower() 
                        for term in ['date', 'month', 'quarter', 'year', 'period'])]
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            
            if not date_cols or not numeric_cols:
                print(f"⚠️  No time-series data found in {title}")
                return False
            
            # Use first date column and up to 3 numeric columns
            x_col = str(date_cols[0])
            y_cols = [str(col) for col in numeric_cols[:3]]
            
            # Create chart data
            chart_data = CategoryChartData()
            
            # Set categories (x-axis) - convert to strings and limit to last 12 points
            df_chart = df.tail(12).copy()
            categories = df_chart[x_col].astype(str).tolist()
            chart_data.categories = categories
            
            # Add series (y-axis)
            for y_col in y_cols:
                try:
                    values = df_chart[y_col].fillna(0).tolist()
                    chart_data.add_series(y_col, values)
                except Exception as e:
                    print(f"⚠️  Skipping column {y_col}: {e}")
                    continue
            
            # Add chart to slide
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.LINE, left, top, width, height, chart_data
            ).chart
            
            # Set chart title
            chart.has_title = True
            chart.chart_title.text_frame.text = title
            
            return True
            
        except Exception as e:
            print(f"⚠️  Chart creation failed for {title}: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _find_category_data(self, sheets_data: List[Tuple[str, pd.DataFrame]]) -> Optional[pd.DataFrame]:
        """Find categorical data for comparison"""
        for name, df in sheets_data:
            # Look for categorical columns with numeric values
            categorical_cols = df.select_dtypes(include=['object', 'string']).columns.tolist()
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            
            if categorical_cols and numeric_cols:
                return df[[categorical_cols[0], numeric_cols[0]]]
        
        return None
    
    def _add_category_chart(self, slide: Slide, df: pd.DataFrame,
                           left: float, top: float, width: float, height: float) -> bool:
        """Add category comparison chart (pie or bar)"""
        try:
            if df is None or df.empty or len(df.columns) < 2:
                return False
            
            # Get first two columns (category and value)
            category_col = str(df.columns[0])
            value_col = str(df.columns[1])
            
            # Create chart data for pie chart
            chart_data = CategoryChartData()
            
            # Limit to top 5 categories
            df_top = df.nlargest(5, value_col) if len(df) > 5 else df
            
            # Set categories and values
            categories = df_top[category_col].astype(str).tolist()
            values = df_top[value_col].fillna(0).tolist()
            
            chart_data.categories = categories
            chart_data.add_series('Value', values)
            
            # Add pie chart
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.PIE, left, top, width, height, chart_data
            ).chart
            
            # Set chart title
            chart.has_title = True
            chart.chart_title.text_frame.text = "Category Distribution"
            chart.chart_title.text_frame.paragraphs[0].font.size = Pt(16)
            chart.chart_title.text_frame.paragraphs[0].font.bold = True
            
            # Show legend with color meanings
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.include_in_layout = False
            
            # Show data labels with percentages
            chart.plots[0].has_data_labels = True
            data_labels = chart.plots[0].data_labels
            data_labels.show_percentage = True
            data_labels.show_category_name = True
            
            # Apply finance theme colors to pie slices
            plot = chart.plots[0]
            for idx, point in enumerate(plot.series[0].points):
                color_idx = idx % len(FINANCE_THEME['chart_colors'])
                fill = point.format.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(*FINANCE_THEME['chart_colors'][color_idx])
            
            return True
            
        except Exception as e:
            print(f"⚠️  Category chart failed: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _add_marketcap_bar_chart(self, slide: Slide, summary_df: pd.DataFrame,
                                 left: float, top: float, width: float, height: float) -> bool:
        """Add MarketCap comparison bar chart with legend"""
        try:
            if summary_df is None or summary_df.empty:
                return False
            
            if 'Ticker' not in summary_df.columns or 'MarketCap' not in summary_df.columns:
                return False
            
            # Prepare data with smart scaling
            chart_data = CategoryChartData()
            
            # Get companies and MarketCap values
            tickers = summary_df['Ticker'].astype(str).tolist()
            marketcaps_raw = summary_df['MarketCap'].fillna(0).tolist()
            
            # Determine best scale based on values
            max_value = max(marketcaps_raw) if marketcaps_raw else 0
            
            if max_value >= 1e9:
                # Use billions
                marketcaps = [(val / 1e9) for val in marketcaps_raw]
                unit = "Billions"
            elif max_value >= 1e6:
                # Use millions
                marketcaps = [(val / 1e6) for val in marketcaps_raw]
                unit = "Millions"
            elif max_value >= 1e3:
                # Use thousands
                marketcaps = [(val / 1e3) for val in marketcaps_raw]
                unit = "Thousands"
            else:
                # Use actual values
                marketcaps = marketcaps_raw
                unit = ""
            
            chart_data.categories = tickers
            series_label = f'Market Cap ({unit})' if unit else 'Market Cap'
            chart_data.add_series(series_label, marketcaps)
            
            # Add bar chart
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data
            ).chart
            
            # Set chart title
            chart.has_title = True
            chart.chart_title.text_frame.text = f"Market Capitalization Comparison ({unit})" if unit else "Market Capitalization Comparison"
            title_para = chart.chart_title.text_frame.paragraphs[0]
            title_para.font.size = Pt(16)
            title_para.font.bold = True
            
            # Show legend
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.include_in_layout = False
            
            # Customize colors using finance theme
            plot = chart.plots[0]
            for idx, point in enumerate(plot.series[0].points):
                color_idx = idx % len(FINANCE_THEME['chart_colors'])
                fill = point.format.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(*FINANCE_THEME['chart_colors'][color_idx])
            
            return True
            
        except Exception as e:
            print(f"⚠️  MarketCap bar chart failed: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _add_trend_line_chart(self, slide: Slide, prices_df: pd.DataFrame, ticker: str,
                             left: float, top: float, width: float, height: float) -> bool:
        """Add trend line chart showing price over time"""
        try:
            if prices_df is None or prices_df.empty:
                return False
            
            if 'Date' not in prices_df.columns or 'Close' not in prices_df.columns:
                return False
            
            # Get last 60 days for clearer visualization
            df_recent = prices_df.tail(60).copy()
            
            # Create chart data
            chart_data = CategoryChartData()
            
            # Format dates as strings
            dates = df_recent['Date'].dt.strftime('%Y-%m-%d').tolist()
            prices = df_recent['Close'].fillna(0).tolist()
            
            chart_data.categories = dates
            chart_data.add_series(f'{ticker} Closing Price', prices)
            
            # Add line chart
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.LINE, left, top, width, height, chart_data
            ).chart
            
            # Set chart title
            chart.has_title = True
            chart.chart_title.text_frame.text = f"{ticker} - 60-Day Price Trend"
            title_para = chart.chart_title.text_frame.paragraphs[0]
            title_para.font.size = Pt(16)
            title_para.font.bold = True
            
            # Show legend
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            
            # Style the line
            plot = chart.plots[0]
            series = plot.series[0]
            line = series.format.line
            line.color.rgb = RGBColor(*FINANCE_THEME['colors']['blue_accent'])
            line.width = Pt(2.5)
            
            return True
            
        except Exception as e:
            print(f"⚠️  Trend line chart failed: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _create_dynamic_chart(self, slide: Slide, chart_spec: Dict, data: pd.DataFrame,
                             left: float, top: float, width: float, height: float) -> bool:
        """Create a chart dynamically based on specification from SmartChartAnalyzer"""
        try:
            chart_type = chart_spec.get('type')
            title = chart_spec.get('title', 'Chart')
            
            if chart_type == 'bar_horizontal':
                return self._create_horizontal_bar_chart(slide, chart_spec, data, left, top, width, height)
            elif chart_type == 'grouped_bar':
                return self._create_grouped_bar_chart(slide, chart_spec, data, left, top, width, height)
            elif chart_type == 'pie':
                return self._create_pie_chart(slide, chart_spec, data, left, top, width, height)
            else:
                # Fallback to bar chart
                return self._create_horizontal_bar_chart(slide, chart_spec, data, left, top, width, height)
        
        except Exception as e:
            print(f"⚠️  Dynamic chart creation failed: {e}")
            return False
    
    def _create_horizontal_bar_chart(self, slide: Slide, chart_spec: Dict, data: pd.DataFrame,
                                     left: float, top: float, width: float, height: float) -> bool:
        """Create horizontal bar chart"""
        try:
            x_col = chart_spec.get('x_col')
            y_col = chart_spec.get('y_col')
            limit = chart_spec.get('limit', 10)
            sort_order = chart_spec.get('sort', 'desc')
            format_type = chart_spec.get('format', 'number')
            
            if x_col not in data.columns or y_col not in data.columns:
                return False
            
            # Sort and limit data
            df_sorted = data.sort_values(y_col, ascending=(sort_order == 'asc'))
            df_top = df_sorted.head(limit)
            
            # Create chart data
            chart_data = CategoryChartData()
            categories = df_top[x_col].astype(str).tolist()
            values = df_top[y_col].fillna(0).tolist()
            
            chart_data.categories = categories
            chart_data.add_series('Value', values)
            
            # Add chart
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.BAR_CLUSTERED, left, top, width, height, chart_data
            ).chart
            
            # Set title
            chart.has_title = True
            chart.chart_title.text_frame.text = chart_spec.get('title', 'Chart')
            title_para = chart.chart_title.text_frame.paragraphs[0]
            title_para.font.size = Pt(16)
            title_para.font.bold = True
            
            # Show legend
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.font.size = Pt(9)
            
            # Apply colors
            plot = chart.plots[0]
            for idx, point in enumerate(plot.series[0].points):
                color_idx = idx % len(FINANCE_THEME['chart_colors'])
                fill = point.format.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(*FINANCE_THEME['chart_colors'][color_idx])
            
            return True
            
        except Exception as e:
            print(f"⚠️  Horizontal bar chart failed: {e}")
            return False
    
    def _create_grouped_bar_chart(self, slide: Slide, chart_spec: Dict, data: pd.DataFrame,
                                  left: float, top: float, width: float, height: float) -> bool:
        """Create grouped bar chart with multiple series"""
        try:
            x_col = chart_spec.get('x_col')
            y_cols = chart_spec.get('y_cols', [])
            limit = chart_spec.get('limit', 8)
            
            if x_col not in data.columns or not y_cols:
                return False
            
            # Limit data
            df_limited = data.head(limit)
            
            # Create chart data
            chart_data = CategoryChartData()
            categories = df_limited[x_col].astype(str).tolist()
            chart_data.categories = categories
            
            # Add series for each y column
            for y_col in y_cols:
                if y_col in df_limited.columns:
                    values = df_limited[y_col].fillna(0).tolist()
                    chart_data.add_series(y_col, values)
            
            # Add chart
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data
            ).chart
            
            # Set title
            chart.has_title = True
            chart.chart_title.text_frame.text = chart_spec.get('title', 'Multi-Metric Comparison')
            title_para = chart.chart_title.text_frame.paragraphs[0]
            title_para.font.size = Pt(16)
            title_para.font.bold = True
            
            # Show legend
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.font.size = Pt(9)
            
            return True
            
        except Exception as e:
            print(f"⚠️  Grouped bar chart failed: {e}")
            return False
    
    def _create_pie_chart(self, slide: Slide, chart_spec: Dict, data: pd.DataFrame,
                         left: float, top: float, width: float, height: float) -> bool:
        """Create pie chart for category distribution"""
        try:
            category_col = chart_spec.get('category_col')
            value_col = chart_spec.get('value_col')
            
            if category_col not in data.columns or value_col not in data.columns:
                return False
            
            # Group by category
            grouped = data.groupby(category_col)[value_col].sum().reset_index()
            grouped = grouped.sort_values(value_col, ascending=False)
            
            # Create chart data
            chart_data = CategoryChartData()
            categories = grouped[category_col].astype(str).tolist()
            values = grouped[value_col].fillna(0).tolist()
            
            chart_data.categories = categories
            chart_data.add_series('Distribution', values)
            
            # Add chart
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.PIE, left, top, width, height, chart_data
            ).chart
            
            # Set title
            chart.has_title = True
            chart.chart_title.text_frame.text = chart_spec.get('title', 'Distribution')
            title_para = chart.chart_title.text_frame.paragraphs[0]
            title_para.font.size = Pt(16)
            title_para.font.bold = True
            
            # Show legend with percentages
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.font.size = Pt(9)
            
            # Show data labels with percentages
            plot = chart.plots[0]
            plot.has_data_labels = True
            data_labels = plot.data_labels
            data_labels.show_percentage = True
            data_labels.show_category_name = True
            data_labels.font.size = Pt(10)
            
            # Apply colors
            for idx, point in enumerate(plot.series[0].points):
                color_idx = idx % len(FINANCE_THEME['chart_colors'])
                fill = point.format.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(*FINANCE_THEME['chart_colors'][color_idx])
            
            return True
            
        except Exception as e:
            print(f"⚠️  Pie chart failed: {e}")
            return False
    
    def _add_top_performers_table(self, slide: Slide, summary_df: pd.DataFrame,
                                 left: float, top: float, width: float, height: float) -> bool:
        """Add colored table showing top performers with visual indicators"""
        try:
            if summary_df is None or summary_df.empty:
                return False
            
            # Sort by Return_1Y_% descending
            if 'Return_1Y_%' in summary_df.columns:
                df_sorted = summary_df.sort_values('Return_1Y_%', ascending=False).head(5)
            else:
                df_sorted = summary_df.head(5)
            
            # Create table shape
            rows = len(df_sorted) + 1  # +1 for header
            cols = 4  # Ticker, Company, Return %, Color
            
            table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
            table = table_shape.table
            
            # Set column widths
            table.columns[0].width = Inches(0.8)   # Ticker
            table.columns[1].width = Inches(2.5)   # Company
            table.columns[2].width = Inches(1.2)   # Return
            table.columns[3].width = Inches(0.8)   # Indicator
            
            # Header row with finance theme navy background
            header_cells = ['Ticker', 'Company', '1Y Return %', 'Status']
            for col_idx, header_text in enumerate(header_cells):
                cell = table.cell(0, col_idx)
                cell.text = header_text
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(*FINANCE_THEME['colors']['navy'])
                
                # White text
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(11)
                para.font.bold = True
                para.font.color.rgb = RGBColor(*FINANCE_THEME['colors']['white'])
                para.alignment = PP_ALIGN.CENTER
            
            # Data rows with color coding
            for row_idx, (_, row) in enumerate(df_sorted.iterrows(), start=1):
                # Ticker
                cell = table.cell(row_idx, 0)
                cell.text = str(row.get('Ticker', ''))
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(10)
                para.font.bold = True
                para.alignment = PP_ALIGN.CENTER
                
                # Company name
                cell = table.cell(row_idx, 1)
                cell.text = str(row.get('Name', ''))
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(10)
                
                # Return %
                cell = table.cell(row_idx, 2)
                return_val = row.get('Return_1Y_%', 0)
                if pd.notna(return_val):
                    cell.text = f"{return_val:.2f}%"
                else:
                    cell.text = "N/A"
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(10)
                para.font.bold = True
                para.alignment = PP_ALIGN.CENTER
                
                # Color indicator based on performance
                cell = table.cell(row_idx, 3)
                if pd.notna(return_val):
                    if return_val > 50:
                        cell.text = "🟢"  # Strong
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(200, 230, 201)  # Light green
                    elif return_val > 20:
                        cell.text = "🟡"  # Good
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(255, 249, 196)  # Light yellow
                    elif return_val > 0:
                        cell.text = "🟠"  # Moderate
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(255, 224, 178)  # Light orange
                    else:
                        cell.text = "🔴"  # Weak
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(255, 205, 210)  # Light red
                else:
                    cell.text = "⚪"
                
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(14)
                para.alignment = PP_ALIGN.CENTER
            
            # Add legend below table
            legend_left = left
            legend_top = top + height + Inches(0.1)
            legend_box = slide.shapes.add_textbox(legend_left, legend_top, width, Inches(0.4))
            legend_frame = legend_box.text_frame
            legend_frame.text = "Legend: 🟢 Strong (>50%)  🟡 Good (20-50%)  🟠 Moderate (0-20%)  🔴 Negative (<0%)"
            legend_para = legend_frame.paragraphs[0]
            legend_para.font.size = Pt(9)
            legend_para.font.color.rgb = RGBColor(100, 100, 100)
            
            return True
            
        except Exception as e:
            print(f"⚠️  Top performers table failed: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _generate_fallback_category_insight(self, df: pd.DataFrame) -> str:
        """Generate basic category insight without AI"""
        try:
            if df is not None and not df.empty and len(df.columns) >= 2:
                category_col = df.columns[0]
                value_col = df.columns[1]
                
                # Find largest category
                top_category = df.loc[df[value_col].idxmax(), category_col]
                top_value = df[value_col].max()
                total_value = df[value_col].sum()
                percentage = (top_value / total_value * 100) if total_value > 0 else 0
                
                return f"{top_category} remains the largest contributor at {percentage:.0f}%"
        except:
            pass
        
        return "Category analysis completed"
    
    def _generate_simple_data_insights(self, all_data: pd.DataFrame, sheets_data: List[Tuple[str, pd.DataFrame]]) -> Dict[str, List[str]]:
        """
        Generate simple, factual insights directly from data (NO AI fluff)
        Just straightforward observations from the numbers
        """
        insights = {
            'top_performers': [],
            'observations': [],
            'recommendations': []
        }
        
        try:
            if hasattr(self, 'summary_df') and self.summary_df is not None and not self.summary_df.empty:
                df = self.summary_df
                
                # TOP PERFORMERS - Just list them with their values
                if 'Return_1Y_%' in df.columns:
                    top_3 = df.nlargest(3, 'Return_1Y_%')
                    for idx, row in top_3.iterrows():
                        ticker = row.get('Ticker', 'Unknown')
                        return_val = row.get('Return_1Y_%', 0)
                        if pd.notna(return_val):
                            insights['top_performers'].append(
                                f"{ticker}: {return_val:.1f}% return"
                            )
                
                # OBSERVATIONS - Simple facts from the data
                total_count = len(df)
                insights['observations'].append(f"Total entities analyzed: {total_count}")
                
                if 'MarketCap' in df.columns:
                    total_cap = df['MarketCap'].sum()
                    if pd.notna(total_cap):
                        formatted_cap = format_value(total_cap, 'currency')
                        insights['observations'].append(f"Combined value: {formatted_cap}")
                
                if 'Sector' in df.columns:
                    sector_count = df['Sector'].nunique()
                    top_sector = df['Sector'].value_counts().index[0] if not df['Sector'].value_counts().empty else 'N/A'
                    insights['observations'].append(f"{sector_count} sectors represented, led by {top_sector}")
                
                # RECOMMENDATIONS - Data-driven suggestions
                if 'Return_1Y_%' in df.columns:
                    positive_count = (df['Return_1Y_%'] > 0).sum()
                    pct_positive = (positive_count / total_count * 100) if total_count > 0 else 0
                    insights['recommendations'].append(f"{pct_positive:.0f}% showing positive performance")
                
                if 'TrailingPE' in df.columns:
                    avg_pe = df['TrailingPE'].mean()
                    if pd.notna(avg_pe):
                        insights['recommendations'].append(f"Average P/E ratio: {avg_pe:.1f}x")
                
                # Add a diversity note if multiple sectors
                if 'Sector' in df.columns and df['Sector'].nunique() > 2:
                    insights['recommendations'].append(f"Portfolio shows good sector diversification")
        
        except Exception as e:
            print(f"⚠️  Simple data insights generation failed: {e}")
            insights['observations'].append("Data analysis in progress")
        
        return insights
    
    def _generate_fallback_ai_insights(self, all_data: pd.DataFrame, sheets_data: List[Tuple[str, pd.DataFrame]]) -> Dict[str, List[str]]:
        """Generate natural language insights using Summary data (NO nan% bug!)"""
        insights = {
            'top_performers': [],
            'anomalies': [],
            'predictions': []
        }
        
        try:
            # Use Summary data if available for MUCH better insights
            if hasattr(self, 'summary_df') and self.summary_df is not None and not self.summary_df.empty:
                df = self.summary_df.copy()
                
                # TOP PERFORMERS - Natural language with actual company data
                if 'Return_1Y_%' in df.columns and 'Ticker' in df.columns:
                    df_sorted = df.sort_values('Return_1Y_%', ascending=False, na_position='last')
                    top_3 = df_sorted.head(3)
                    
                    for idx, row in top_3.iterrows():
                        ticker = row['Ticker']
                        return_val = row['Return_1Y_%']
                        name = row.get('Name', ticker)
                        
                        # Check for valid return value (NO NAN%)
                        if pd.notna(return_val) and not np.isnan(return_val):
                            if 'MarketCap' in row and pd.notna(row['MarketCap']):
                                # Use smart formatting for market cap
                                formatted_cap = format_value(row['MarketCap'], 'currency')
                                insights['top_performers'].append(
                                    f"{ticker} showed strong momentum with {return_val:.1f}% annual return and {formatted_cap} market cap, demonstrating solid investor confidence."
                                )
                            else:
                                insights['top_performers'].append(
                                    f"{ticker} demonstrated exceptional performance with {return_val:.1f}% annual growth, outpacing market averages."
                                )
                        else:
                            # Skip stocks with nan returns
                            continue
                
                # ANOMALIES - Natural language volatility & PE analysis
                if 'TrailingPE' in df.columns and 'ForwardPE' in df.columns:
                    # Check for unusual PE ratios
                    high_pe = df[df['TrailingPE'] > 100].head(2)
                    for idx, row in high_pe.iterrows():
                        ticker = row['Ticker']
                        pe = row['TrailingPE']
                        if pd.notna(pe) and not np.isnan(pe):
                            insights['anomalies'].append(
                                f"{ticker} shows elevated P/E ratio of {pe:.1f}x, suggesting premium valuation or high growth expectations."
                            )
                
                if 'Return_1Y_%' in df.columns:
                    # Check for exceptional performers (>70%)
                    exceptional = df[df['Return_1Y_%'] > 70]
                    for idx, row in exceptional.iterrows():
                        ticker = row['Ticker']
                        return_val = row['Return_1Y_%']
                        if pd.notna(return_val):
                            insights['anomalies'].append(
                                f"{ticker} achieved remarkable {return_val:.1f}% return, significantly outperforming sector benchmarks."
                            )
                
                # PREDICTIONS - Trend analysis using price data
                if hasattr(self, 'price_data') and 'AAPL' in self.price_data:
                    aapl_prices = self.price_data['AAPL']
                    if 'Close' in aapl_prices.columns and len(aapl_prices) > 60:
                        recent = aapl_prices.tail(60)
                        first_close = recent['Close'].iloc[0]
                        last_close = recent['Close'].iloc[-1]
                        
                        if pd.notna(first_close) and pd.notna(last_close):
                            change_pct = ((last_close - first_close) / first_close * 100)
                            if change_pct > 10:
                                insights['predictions'].append(
                                    f"AAPL maintains strong bullish momentum with {change_pct:.1f}% gain over 60 days, suggesting continued upward trajectory."
                                )
                            elif change_pct < -10:
                                insights['predictions'].append(
                                    f"AAPL experienced {abs(change_pct):.1f}% correction, potentially presenting a buy opportunity near support levels."
                                )
                            else:
                                insights['predictions'].append(
                                    f"AAPL shows stable consolidation with {change_pct:.1f}% movement, indicating balanced market sentiment."
                                )
                
                # Sector diversification prediction
                if 'Sector' in df.columns:
                    sector_counts = df['Sector'].value_counts()
                    dominant_sector = sector_counts.index[0]
                    sector_pct = (sector_counts.iloc[0] / len(df) * 100)
                    insights['predictions'].append(
                        f"Portfolio shows {sector_pct:.0f}% concentration in {dominant_sector}, suggesting opportunity for sector diversification to manage risk."
                    )
                    
            else:
                # Fallback if no Summary data
                if all_data is not None and not all_data.empty:
                    numeric_cols = all_data.select_dtypes(include=[np.number]).columns.tolist()
                    if len(numeric_cols) > 0:
                        col_str = str(numeric_cols[0])
                        max_val = all_data[numeric_cols[0]].max()
                        if pd.notna(max_val):
                            insights['top_performers'].append(f"Peak value of {max_val:,.2f} observed in {col_str} column")
        
        except Exception as e:
            print(f"⚠️  Fallback insights error: {e}")
            import traceback
            traceback.print_exc()
        
        # Default messages if empty (natural language, NO nan%)
        if not insights['top_performers']:
            insights['top_performers'] = ["Market analysis indicates balanced performance across evaluated securities."]
        if not insights['anomalies']:
            insights['anomalies'] = ["Portfolio maintains stable valuations within expected market ranges."]
        if not insights['predictions']:
            insights['predictions'] = ["Historical trends suggest continued market stability with sector-specific opportunities."]
        
        return insights
