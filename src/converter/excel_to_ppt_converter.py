"""
Excel to PowerPoint Converter with Tiered Plans
Supports Free, Basic ($25), Pro ($49), and AI Pro ($99) plans
"""

import os
import sys
from typing import Optional, List, Dict, Any
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData

# Add paths for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from src.converter.excel_reader import excel_reader_all_sheets
from src.converter.chart_detector import detect_chart_type, should_create_chart
from src.templates.template_manager import TemplateManager
from src.converter.professional_slide_builder import ProfessionalSlideBuilder

# Import AI service (optional, only for paid tiers)
try:
    from src.backend.app.services.ai_service import create_ai_service
    AI_AVAILABLE = True
except:
    AI_AVAILABLE = False


# ============================================================================
# TIER CONFIGURATION
# ============================================================================

TIER_CONFIG = {
    'free': {
        'name': 'Free',
        'ppt_limit': 1,
        'templates': ['minimal_white'],  # Basic template only
        'ai_features': [],  # No AI
        'multi_sheet': False,
        'max_sheets': 1
    },
    'basic': {
        'name': 'Basic ($25/month)',
        'ppt_limit': 7,
        'templates': ['minimal_white'],  # Basic template only
        'ai_features': ['title'],  # AI titles only
        'multi_sheet': True,
        'max_sheets': 5
    },
    'pro': {
        'name': 'Pro ($49/month)',
        'ppt_limit': 15,
        'templates': 'all',  # All 10 professional templates
        'ai_features': ['title', 'template_selection'],  # AI titles + template selection
        'multi_sheet': True,
        'max_sheets': 20
    },
    'ai_pro': {
        'name': 'AI Pro ($99/month)',
        'ppt_limit': -1,  # Unlimited
        'templates': 'all',  # All 10 professional templates
        'ai_features': ['title', 'summary', 'insights', 'template_selection', 'layout', 'chart_type'],  # All AI features
        'multi_sheet': True,
        'max_sheets': -1  # Unlimited
    }
}


class ExcelToPPTConverter:
    """
    Excel to PowerPoint converter with tiered subscription support
    """
    
    def __init__(self, user_tier: str = 'free', user_id: Optional[str] = None, user_metadata: Optional[Dict[str, str]] = None):
        """
        Initialize converter with user tier
        
        Args:
            user_tier: One of 'free', 'basic', 'pro', 'ai_pro'
            user_id: Optional user ID for tracking usage
            user_metadata: Optional dict with 'name', 'company', 'email' for branding
        """
        if user_tier not in TIER_CONFIG:
            raise ValueError(f"Invalid tier: {user_tier}. Must be one of: {list(TIER_CONFIG.keys())}")
        
        self.user_tier = user_tier
        self.user_id = user_id
        self.user_metadata = user_metadata or {}
        self.config = TIER_CONFIG[user_tier]
        self.template_manager = TemplateManager()
        
        # Initialize AI service if available and tier supports it
        self.ai_service = None
        if AI_AVAILABLE and self.config['ai_features']:
            try:
                self.ai_service = create_ai_service()
            except Exception as e:
                print(f"Warning: AI service initialization failed: {e}")
    
    def check_limits(self, user_ppt_count: int) -> bool:
        """
        Check if user has reached their PPT limit
        
        Args:
            user_ppt_count: Number of PPTs user has created this month
            
        Returns:
            True if user can create more PPTs, False otherwise
        """
        limit = self.config['ppt_limit']
        if limit == -1:  # Unlimited
            return True
        return user_ppt_count < limit
    
    def get_allowed_templates(self) -> List[str]:
        """Get list of templates allowed for this tier"""
        if self.config['templates'] == 'all':
            # Extract template IDs from template manager
            template_dicts = self.template_manager.list_templates()
            return [t['id'] for t in template_dicts]
        return self.config['templates']
    
    def can_use_ai_feature(self, feature: str) -> bool:
        """Check if tier supports specific AI feature"""
        return feature in self.config['ai_features']
    
    def convert(self, 
                excel_path: str, 
                output_path: str,
                template_name: Optional[str] = None,
                presentation_title: Optional[str] = None,
                user_ppt_count: int = 0) -> Dict[str, Any]:
        """
        Convert Excel file to PowerPoint presentation
        
        Args:
            excel_path: Path to Excel file
            output_path: Path to save PowerPoint file
            template_name: Template to use (optional, will use AI for pro+ tiers)
            presentation_title: Custom title (optional)
            user_ppt_count: Number of PPTs user has created this month
            
        Returns:
            Dict with conversion results and metadata
        """
        
        # Check limits
        if not self.check_limits(user_ppt_count):
            return {
                'success': False,
                'error': f'PPT limit reached. {self.config["name"]} allows {self.config["ppt_limit"]} PPTs/month.',
                'upgrade_required': True
            }
        
        try:
            # Read Excel file
            print(f"Reading Excel file: {excel_path}")
            sheets_dict = excel_reader_all_sheets(excel_path)
            
            if not sheets_dict:
                return {
                    'success': False,
                    'error': 'No data found in Excel file'
                }
            
            # Convert dict to list of tuples for processing
            sheets_data = [(name, df) for name, df in sheets_dict.items() if df is not None]
            
            # Limit sheets based on tier
            max_sheets = self.config['max_sheets']
            if max_sheets > 0 and len(sheets_data) > max_sheets:
                sheets_data = sheets_data[:max_sheets]
                print(f"Limited to {max_sheets} sheets for {self.config['name']} tier")
            
            # AI Template Selection (for Pro and AI Pro tiers)
            if template_name is None and self.can_use_ai_feature('template_selection') and self.ai_service:
                print("Using AI to select best template...")
                sheets_info = [
                    {
                        'name': sheet_name,
                        'data_type': 'time_series' if any('date' in str(col).lower() or 'month' in str(col).lower() 
                                                          or 'quarter' in str(col).lower() 
                                                          for col in df.columns) else 'comparison',
                        'rows': len(df),
                        'cols': len(df.columns)
                    }
                    for sheet_name, df in sheets_data
                ]
                
                file_name = os.path.basename(excel_path)
                template_rec = self.ai_service.recommend_template(
                    file_name, 
                    sheets_info, 
                    self.get_allowed_templates()
                )
                template_name = template_rec.get('auto_selected', 'corporate_blue')
                print(f"AI selected template: {template_name}")
            
            # Fallback to default template
            if template_name is None:
                allowed_templates = self.get_allowed_templates()
                template_name = allowed_templates[0] if allowed_templates else 'corporate_blue'
            
            # Validate template is allowed for this tier
            if template_name not in self.get_allowed_templates():
                print(f"Warning: Template {template_name} not allowed for {self.config['name']} tier. Using default.")
                allowed_templates = self.get_allowed_templates()
                template_name = allowed_templates[0] if allowed_templates else 'corporate_blue'
            
            # Load template
            template = self.template_manager.load_template(template_name)
            
            # Create presentation
            prs = Presentation()
            
            # Create title slide
            if presentation_title is None:
                presentation_title = os.path.splitext(os.path.basename(excel_path))[0]
            
            self._create_title_slide(prs, presentation_title, template)
            
            # Process each sheet
            slides_created = 0
            ai_usage = {'tokens': 0, 'cost': 0.0}
            
            for sheet_name, df in sheets_data:
                print(f"\nProcessing sheet: {sheet_name}")
                
                # Create slide for this sheet
                slide_result = self._create_sheet_slide(
                    prs, sheet_name, df, template
                )
                
                if slide_result['success']:
                    slides_created += 1
                    if 'ai_usage' in slide_result:
                        ai_usage['tokens'] += slide_result['ai_usage'].get('tokens', 0)
                        ai_usage['cost'] += slide_result['ai_usage'].get('cost', 0.0)
            
            # Save presentation
            prs.save(output_path)
            print(f"\n✅ Presentation saved: {output_path}")
            
            return {
                'success': True,
                'output_path': output_path,
                'slides_created': slides_created + 1,  # +1 for title slide
                'template_used': template_name,
                'user_tier': self.user_tier,
                'ai_features_used': self.config['ai_features'],
                'ai_usage': ai_usage
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }
    
    def convert_professional(self, 
                            excel_path: str, 
                            output_path: str,
                            template_name: Optional[str] = None,
                            presentation_title: Optional[str] = None,
                            user_ppt_count: int = 0,
                            use_professional_structure: bool = True) -> Dict[str, Any]:
        """
        Convert Excel file to professional 7-10 slide presentation
        
        Args:
            excel_path: Path to Excel file
            output_path: Path to save PowerPoint file
            template_name: Template to use (optional)
            presentation_title: Custom title (optional)
            user_ppt_count: Number of PPTs user has created this month
            use_professional_structure: Use new 7-slide structure (default True)
            
        Returns:
            Dict with conversion results and metadata
        """
        
        # Check limits
        if not self.check_limits(user_ppt_count):
            return {
                'success': False,
                'error': f'PPT limit reached. {self.config["name"]} allows {self.config["ppt_limit"]} PPTs/month.',
                'upgrade_required': True
            }
        
        try:
            # Read Excel file
            print(f"\n📊 Reading Excel file: {excel_path}")
            sheets_dict = excel_reader_all_sheets(excel_path)
            
            if not sheets_dict:
                return {
                    'success': False,
                    'error': 'No data found in Excel file'
                }
            
            # Convert dict to list of tuples for processing
            sheets_data = [(name, df) for name, df in sheets_dict.items() if df is not None]
            
            # Limit sheets based on tier
            max_sheets = self.config['max_sheets']
            if max_sheets > 0 and len(sheets_data) > max_sheets:
                sheets_data = sheets_data[:max_sheets]
                print(f"⚠️  Limited to {max_sheets} sheets for {self.config['name']} tier")
            
            # AI Template Selection
            if template_name is None:
                allowed_templates = self.get_allowed_templates()
                template_name = allowed_templates[0] if allowed_templates else 'corporate_blue'
            
            # Validate template
            if template_name not in self.get_allowed_templates():
                print(f"⚠️  Template {template_name} not allowed for tier. Using default.")
                allowed_templates = self.get_allowed_templates()
                template_name = allowed_templates[0] if allowed_templates else 'corporate_blue'
            
            # Load template or use FINANCE_THEME as fallback
            template = self.template_manager.load_template(template_name)
            if template is None:
                # Use FINANCE_THEME as fallback template
                from src.converter.professional_slide_builder import FINANCE_THEME
                template = {
                    'name': 'Dark Finance',
                    'colors': {
                        'primary': '#192A56',  # Navy
                        'secondary': '#343A40',  # Charcoal
                        'accent': '#FFC107',  # Gold
                        'success': '#2E7D32',  # Green
                        'danger': '#D32F2F',  # Red
                        'chart_colors': ['#2196F3', '#2E7D32', '#FFC107', '#D32F2F', '#9C27B0', '#FF9800']
                    },
                    'fonts': {
                        'title': {'size': 44, 'bold': True, 'color': '#192A56'},
                        'subtitle': {'size': 18, 'bold': False, 'color': '#646464'},
                        'heading': {'size': 28, 'bold': True, 'color': '#192A56'},
                        'body': {'size': 14, 'bold': False, 'color': '#343A40'}
                    }
                }
                print(f"✅ Using FINANCE_THEME (template file not found)")
            else:
                print(f"✅ Using template: {template_name}")
            
            # Create presentation
            prs = Presentation()
            
            # Set project name
            if presentation_title is None:
                presentation_title = os.path.splitext(os.path.basename(excel_path))[0]
                presentation_title = presentation_title.replace('_', ' ').title()
            
            # Use professional slide builder
            print(f"\n🎨 Building professional {self.user_tier.upper()} presentation...")
            slide_builder = ProfessionalSlideBuilder(
                user_tier=self.user_tier,
                ai_service=self.ai_service,
                user_metadata=self.user_metadata
            )
            
            # Build all slides
            build_results = slide_builder.build_professional_presentation(
                prs=prs,
                sheets_data=sheets_data,
                project_name=presentation_title,
                template=template,
                excel_path=excel_path  # Pass excel path for Summary and price data
            )
            
            # Save presentation
            prs.save(output_path)
            print(f"\n✅ Professional presentation saved: {output_path}")
            print(f"📊 Total slides: {build_results['total_slides']}")
            print(f"🤖 AI features used: {len(build_results['ai_features_used'])}")
            
            if build_results['errors']:
                print(f"⚠️  Errors encountered: {len(build_results['errors'])}")
                for error in build_results['errors']:
                    print(f"   - {error}")
            
            return {
                'success': True,
                'output_path': output_path,
                'slides_created': build_results['total_slides'],
                'template_used': template_name,
                'user_tier': self.user_tier,
                'ai_features_used': build_results['ai_features_used'],
                'presentation_type': 'professional_7_slide',
                'errors': build_results['errors']
            }
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {
                'success': False,
                'error': str(e)
            }
    
    def _create_title_slide(self, prs: Presentation, title: str, template: Dict[str, Any]):
        """Create title slide with template styling"""
        title_slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(title_slide_layout)
        
        # Set title
        title_shape = slide.shapes.title
        title_shape.text = title
        
        # Apply template colors
        if template and 'colors' in template:
            colors = template['colors']
            # Apply primary color to title
            if 'primary' in colors:
                try:
                    r, g, b = int(colors['primary'][1:3], 16), int(colors['primary'][3:5], 16), int(colors['primary'][5:7], 16)
                    for paragraph in title_shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(r, g, b)
                except:
                    pass
        
        # Set subtitle
        if len(slide.placeholders) > 1:
            subtitle = slide.placeholders[1]
            subtitle.text = f"{self.config['name']} Plan"
    
    def _create_sheet_slide(self, prs: Presentation, sheet_name: str, 
                           df: pd.DataFrame, template: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create a slide for a single sheet
        
        Returns:
            Dict with success status and AI usage info
        """
        
        ai_usage = {'tokens': 0, 'cost': 0.0}
        
        try:
            # Check if we should create a chart
            if not should_create_chart(df):
                print(f"Skipping {sheet_name}: No suitable chart data")
                return {'success': False}
            
            # Detect chart type
            chart_type, chart_config = detect_chart_type(df)
            
            if chart_type == 'unknown':
                print(f"Skipping {sheet_name}: Unknown chart type")
                return {'success': False}
            
            # AI Chart Type Recommendation (for AI Pro only)
            if self.can_use_ai_feature('chart_type') and self.ai_service:
                try:
                    chart_rec = self.ai_service.recommend_chart_type(
                        df, list(df.columns), f"data from {sheet_name}"
                    )
                    recommended_type = chart_rec['recommended']['type']
                    print(f"AI recommends: {recommended_type} (confidence: {chart_rec['recommended']['confidence']})")
                    
                    # Use AI recommendation if confidence is high
                    if chart_rec['recommended']['confidence'] > 0.8:
                        chart_type = recommended_type
                    
                    # Track AI usage
                    if hasattr(self.ai_service, 'get_usage_stats'):
                        stats = self.ai_service.get_usage_stats()
                        ai_usage['tokens'] = stats.get('total_tokens', 0)
                        ai_usage['cost'] = stats.get('total_cost_usd', 0.0)
                except Exception as e:
                    print(f"AI chart recommendation failed: {e}")
            
            # AI-generated title (for Basic, Pro, AI Pro)
            if self.can_use_ai_feature('title') and self.ai_service:
                try:
                    slide_title = self.ai_service.generate_slide_title(df, sheet_name, chart_type)
                    print(f"AI generated title: '{slide_title}'")
                except Exception as e:
                    print(f"AI title generation failed: {e}")
                    slide_title = sheet_name
            else:
                slide_title = sheet_name
            
            # Create blank slide
            blank_slide_layout = prs.slide_layouts[6]  # Blank layout
            slide = prs.slides.add_slide(blank_slide_layout)
            
            # Add title
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.6))
            title_frame = title_box.text_frame
            title_frame.text = slide_title
            title_para = title_frame.paragraphs[0]
            title_para.font.size = Pt(28)
            title_para.font.bold = True
            
            # Apply template color to title
            if template and 'colors' in template and 'primary' in template['colors']:
                try:
                    r, g, b = int(template['colors']['primary'][1:3], 16), int(template['colors']['primary'][3:5], 16), int(template['colors']['primary'][5:7], 16)
                    title_para.font.color.rgb = RGBColor(r, g, b)
                except:
                    pass
            
            # AI Layout Optimization (for AI Pro only)
            layout_config = None
            if self.can_use_ai_feature('layout') and self.ai_service:
                try:
                    has_insights = self.can_use_ai_feature('insights')
                    layout_config = self.ai_service.optimize_slide_layout(df, chart_type, has_insights)
                    print(f"AI layout: {layout_config.get('layout', 'chart_insights')}")
                except Exception as e:
                    print(f"AI layout optimization failed: {e}")
            
            # Determine positions
            if layout_config and 'positioning' in layout_config:
                pos = layout_config['positioning']
                chart_pos = pos.get('chart', {})
                chart_left = chart_pos.get('left', 0.5)
                chart_top = chart_pos.get('top', 1.5)
                chart_width = chart_pos.get('width', 5.5)
                chart_height = chart_pos.get('height', 5.0)
            else:
                # Default positions
                chart_left = 0.5
                chart_top = 1.5
                chart_width = 5.5
                chart_height = 5.0
            
            # Create chart (with fallback to base chart system if it fails)
            chart_created = False
            try:
                self._add_chart_to_slide(
                    slide, df, chart_type, chart_config,
                    chart_left, chart_top, chart_width, chart_height,
                    template
                )
                chart_created = True
                print(f"✅ Chart created successfully using AI-recommended config")
            except Exception as e:
                print(f"⚠️  AI chart creation failed: {e}")
                print(f"   Falling back to base chart system...")
                
                # Fallback: Use base chart detection and creation
                try:
                    # Re-detect chart using base system
                    fallback_chart_type, fallback_config = detect_chart_type(df)
                    
                    if fallback_chart_type != 'unknown' and fallback_config.get('x_col') and fallback_config.get('y_cols'):
                        print(f"   Base system detected: {fallback_chart_type}")
                        self._add_chart_to_slide(
                            slide, df, fallback_chart_type, fallback_config,
                            chart_left, chart_top, chart_width, chart_height,
                            template
                        )
                        chart_created = True
                        print(f"✅ Chart created successfully using base system")
                    else:
                        print(f"   Base system also couldn't create chart. Continuing without chart...")
                except Exception as fallback_error:
                    print(f"⚠️  Base chart system also failed: {fallback_error}")
                    print(f"   Continuing with slide creation without chart...")
            
            # AI-generated insights (for AI Pro only)
            if self.can_use_ai_feature('insights') and self.ai_service:
                try:
                    insights = self.ai_service.generate_data_insights(df, sheet_name, num_insights=5)
                    
                    # Add insights to right side
                    insights_box = slide.shapes.add_textbox(
                        Inches(6.5), Inches(1.5), Inches(3.5), Inches(5.0)
                    )
                    text_frame = insights_box.text_frame
                    text_frame.word_wrap = True
                    
                    # Add "Key Insights" header
                    p = text_frame.paragraphs[0]
                    p.text = "Key Insights:"
                    p.font.size = Pt(16)
                    p.font.bold = True
                    
                    # Add each insight as bullet point
                    for insight in insights:
                        p = text_frame.add_paragraph()
                        p.text = f"• {insight}"
                        p.level = 0
                        p.font.size = Pt(12)
                    
                    print(f"Added {len(insights)} AI insights")
                except Exception as e:
                    print(f"AI insights generation failed: {e}")
            
            # AI Summary (for AI Pro only)
            if self.can_use_ai_feature('summary') and self.ai_service:
                try:
                    summary = self.ai_service.generate_slide_summary(df, sheet_name, chart_type)
                    
                    # Add summary below chart
                    summary_box = slide.shapes.add_textbox(
                        Inches(0.5), Inches(6.8), Inches(9.0), Inches(0.7)
                    )
                    text_frame = summary_box.text_frame
                    text_frame.word_wrap = True
                    p = text_frame.paragraphs[0]
                    p.text = summary
                    p.font.size = Pt(11)
                    p.font.italic = True
                    
                    print(f"Added AI summary")
                except Exception as e:
                    print(f"AI summary generation failed: {e}")
            
            # Get final AI usage stats
            if self.ai_service and hasattr(self.ai_service, 'get_usage_stats'):
                stats = self.ai_service.get_usage_stats()
                ai_usage['tokens'] = stats.get('total_tokens', 0)
                ai_usage['cost'] = stats.get('total_cost_usd', 0.0)
            
            return {
                'success': True,
                'chart_type': chart_type,
                'ai_usage': ai_usage
            }
            
        except Exception as e:
            print(f"Error creating slide for {sheet_name}: {e}")
            return {'success': False, 'error': str(e)}
    
    def _add_chart_to_slide(self, slide, df: pd.DataFrame, chart_type: str, 
                           chart_config: Dict, left: float, top: float, 
                           width: float, height: float, template: Dict[str, Any]):
        """Add chart to slide based on type"""
        
        x_col = chart_config.get('x_col')
        y_cols = chart_config.get('y_cols', [])
        
        if not x_col or not y_cols:
            raise ValueError(
                f"Invalid chart config: x_col={x_col}, y_cols={y_cols}. "
                f"AI chart detection may have failed - will fallback to base chart system."
            )
        
        # Get template colors for chart
        chart_colors = []
        if template and 'colors' in template and 'chart_colors' in template['colors']:
            for color_hex in template['colors']['chart_colors']:
                try:
                    r = int(color_hex[1:3], 16)
                    g = int(color_hex[3:5], 16)
                    b = int(color_hex[5:7], 16)
                    chart_colors.append(RGBColor(r, g, b))
                except:
                    pass
        
        # Default colors if template doesn't provide them
        if not chart_colors:
            chart_colors = [
                RGBColor(68, 114, 196),   # Blue
                RGBColor(237, 125, 49),   # Orange
                RGBColor(165, 165, 165),  # Gray
                RGBColor(255, 192, 0),    # Yellow
                RGBColor(91, 155, 213),   # Light Blue
            ]
        
        try:
            if chart_type == 'pie':
                self._create_pie_chart(slide, df, x_col, y_cols[0], left, top, width, height, chart_colors)
            elif chart_type == 'bar':
                self._create_bar_chart(slide, df, x_col, y_cols, left, top, width, height, chart_colors)
            elif chart_type == 'line':
                self._create_line_chart(slide, df, x_col, y_cols, left, top, width, height, chart_colors)
            elif chart_type == 'column':
                self._create_column_chart(slide, df, x_col, y_cols, left, top, width, height, chart_colors)
            elif chart_type == 'scatter':
                if len(y_cols) >= 2:
                    self._create_scatter_chart(slide, df, y_cols[0], y_cols[1], left, top, width, height, chart_colors)
        except Exception as e:
            import traceback
            print(f"Error creating {chart_type} chart: {e}")
            print(f"Traceback: {traceback.format_exc()}")
    
    def _create_pie_chart(self, slide, df, category_col, value_col, left, top, width, height, colors):
        """Create pie chart"""
        chart_data = CategoryChartData()
        category_col_str = str(category_col)
        value_col_str = str(value_col)
        chart_data.categories = df[category_col_str].astype(str).tolist()
        chart_data.add_series('Values', df[value_col_str].fillna(0).tolist())
        
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.PIE, Inches(left), Inches(top), Inches(width), Inches(height), chart_data
        ).chart
        
        # Apply colors
        for i, point in enumerate(chart.series[0].points):
            if i < len(colors):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = colors[i]
    
    def _create_bar_chart(self, slide, df, x_col, y_cols, left, top, width, height, colors):
        """Create bar chart"""
        chart_data = CategoryChartData()
        x_col_str = str(x_col)
        chart_data.categories = df[x_col_str].astype(str).tolist()
        
        for i, col in enumerate(y_cols[:3]):  # Max 3 series
            col_str = str(col)
            chart_data.add_series(col_str, df[col_str].fillna(0).tolist())
        
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_CLUSTERED, Inches(left), Inches(top), Inches(width), Inches(height), chart_data
        ).chart
        
        # Apply colors
        for i, series in enumerate(chart.series):
            if i < len(colors):
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = colors[i]
    
    def _create_line_chart(self, slide, df, x_col, y_cols, left, top, width, height, colors):
        """Create line chart"""
        chart_data = CategoryChartData()
        # Convert column name to string in case it's numpy type
        x_col_str = str(x_col)
        chart_data.categories = df[x_col_str].astype(str).tolist()
        
        for i, col in enumerate(y_cols[:3]):  # Max 3 series
            col_str = str(col)  # Convert to string in case it's numpy type
            chart_data.add_series(col_str, df[col_str].dropna().tolist())
        
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE, Inches(left), Inches(top), Inches(width), Inches(height), chart_data
        ).chart
        
        # Apply colors
        for i, series in enumerate(chart.series):
            if i < len(colors):
                series.format.line.color.rgb = colors[i]
                series.format.line.width = Pt(2)
    
    def _create_column_chart(self, slide, df, x_col, y_cols, left, top, width, height, colors):
        """Create column chart"""
        chart_data = CategoryChartData()
        x_col_str = str(x_col)
        chart_data.categories = df[x_col_str].astype(str).tolist()
        
        for i, col in enumerate(y_cols[:3]):  # Max 3 series
            col_str = str(col)
            chart_data.add_series(col_str, df[col_str].fillna(0).tolist())
        
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(left), Inches(top), Inches(width), Inches(height), chart_data
        ).chart
        
        # Apply colors
        for i, series in enumerate(chart.series):
            if i < len(colors):
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = colors[i]
    
    def _create_scatter_chart(self, slide, df, x_col, y_col, left, top, width, height, colors):
        """Create scatter chart"""
        from pptx.chart.data import XyChartData
        
        x_col_str = str(x_col)
        y_col_str = str(y_col)
        
        chart_data = XyChartData()
        series = chart_data.add_series('Data')
        
        for _, row in df[[x_col_str, y_col_str]].dropna().iterrows():
            series.add_data_point(float(row[x_col_str]), float(row[y_col_str]))
        
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.XY_SCATTER, Inches(left), Inches(top), Inches(width), Inches(height), chart_data
        ).chart
        
        # Apply color
        if colors:
            chart.series[0].format.fill.solid()
            chart.series[0].format.fill.fore_color.rgb = colors[0]


# ============================================================================
# CONVENIENCE FUNCTIONS
# ============================================================================

def convert_excel_to_ppt(excel_path: str, 
                        output_path: str,
                        user_tier: str = 'free',
                        template_name: Optional[str] = None,
                        presentation_title: Optional[str] = None,
                        user_ppt_count: int = 0) -> Dict[str, Any]:
    """
    Convenience function to convert Excel to PPT
    
    Args:
        excel_path: Path to Excel file
        output_path: Path to save PowerPoint
        user_tier: User subscription tier ('free', 'basic', 'pro', 'ai_pro')
        template_name: Optional template name
        presentation_title: Optional custom title for the presentation
        user_ppt_count: Number of PPTs user created this month
        
    Returns:
        Dict with conversion results
    """
    converter = ExcelToPPTConverter(user_tier=user_tier)
    return converter.convert(excel_path, output_path, template_name, presentation_title=presentation_title, user_ppt_count=user_ppt_count)
