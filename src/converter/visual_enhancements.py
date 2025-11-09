"""
Professional PPT Visual Enhancements
=====================================
This module adds:
1. Enhanced cover page with branding
2. Keyword highlighting with colors
3. Dynamic icons and visual aids
4. AI-powered insights section
"""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from datetime import datetime
import re


class VisualEnhancer:
    """Add visual polish to PowerPoint presentations"""
    
    # Keyword dictionaries for smart highlighting
    POSITIVE_KEYWORDS = {
        'excellent', 'strong', 'growth', 'increase', 'improved', 'positive',
        'gain', 'profit', 'success', 'outstanding', 'exceptional', 'high',
        'rising', 'surge', 'boost', 'advance', 'expand', 'achieve'
    }
    
    NEGATIVE_KEYWORDS = {
        'decline', 'loss', 'risk', 'decrease', 'negative', 'fall', 'drop',
        'weak', 'poor', 'concern', 'warning', 'alert', 'danger', 'threat',
        'deficit', 'shortfall', 'underperform'
    }
    
    NEUTRAL_KEYWORDS = {
        'stable', 'steady', 'maintain', 'consistent', 'unchanged', 'flat',
        'moderate', 'average', 'normal', 'standard'
    }
    
    # Icon mappings (using shapes as icons)
    METRIC_ICONS = {
        'growth': '📈',
        'business': '💼',
        'analytics': '📊',
        'money': '💰',
        'chart': '📉',
        'target': '🎯',
        'trend': '📈',
        'performance': '⚡',
        'success': '✅',
        'warning': '⚠️',
        'info': 'ℹ️'
    }
    
    def __init__(self, template_colors):
        """Initialize with template colors"""
        self.template_colors = template_colors
        
    def create_enhanced_cover_slide(self, prs, project_name, subtitle=None, data_period=None):
        """
        Create professional cover page with:
        - Company branding
        - Project name
        - Data period
        - "Powered by FinDeck AI" tagline
        """
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
        
        # ===== BACKGROUND =====
        background = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0, 0,
            prs.slide_width, prs.slide_height
        )
        background.fill.solid()
        background.fill.fore_color.rgb = RGBColor(*self.template_colors['navy'])
        background.line.color.rgb = RGBColor(*self.template_colors['navy'])
        
        # ===== ACCENT STRIPE (Top) =====
        accent_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0, 0,
            prs.slide_width, Inches(0.3)
        )
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = RGBColor(*self.template_colors['light_blue'])
        accent_bar.line.fill.background()
        
        # ===== LOGO PLACEHOLDER =====
        # Add a circular shape for logo placement
        logo_circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(0.5), Inches(0.5),
            Inches(1.2), Inches(1.2)
        )
        logo_circle.fill.solid()
        logo_circle.fill.fore_color.rgb = RGBColor(*self.template_colors['white'])
        logo_circle.line.color.rgb = RGBColor(*self.template_colors['light_blue'])
        logo_circle.line.width = Pt(3)
        
        # Add "FD" text in logo circle (FinDeck logo)
        logo_text = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.5),
            Inches(1.2), Inches(1.2)
        )
        logo_frame = logo_text.text_frame
        logo_frame.text = "FD"
        logo_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = logo_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*self.template_colors['navy'])
        
        # ===== MAIN TITLE =====
        title_box = slide.shapes.add_textbox(
            Inches(1), Inches(2.5),
            Inches(8), Inches(1.5)
        )
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        title_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        p = title_frame.paragraphs[0]
        p.text = project_name
        p.font.size = Pt(54)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*self.template_colors['white'])
        p.alignment = PP_ALIGN.CENTER
        
        # ===== SUBTITLE (if provided) =====
        if subtitle:
            subtitle_box = slide.shapes.add_textbox(
                Inches(1.5), Inches(4.2),
                Inches(7), Inches(0.6)
            )
            subtitle_frame = subtitle_box.text_frame
            p = subtitle_frame.paragraphs[0]
            p.text = subtitle
            p.font.size = Pt(24)
            p.font.color.rgb = RGBColor(*self.template_colors['light_blue'])
            p.alignment = PP_ALIGN.CENTER
        
        # ===== DATA PERIOD =====
        period_text = data_period if data_period else f"Generated {datetime.now().strftime('%B %Y')}"
        period_box = slide.shapes.add_textbox(
            Inches(1.5), Inches(5.0),
            Inches(7), Inches(0.5)
        )
        period_frame = period_box.text_frame
        p = period_frame.paragraphs[0]
        p.text = period_text
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(*self.template_colors['gray'])
        p.alignment = PP_ALIGN.CENTER
        
        # ===== BRANDING FOOTER =====
        branding_box = slide.shapes.add_textbox(
            Inches(2), Inches(6.5),
            Inches(6), Inches(0.5)
        )
        branding_frame = branding_box.text_frame
        p = branding_frame.paragraphs[0]
        p.text = "Powered by FinDeck AI"
        p.font.size = Pt(14)
        p.font.italic = True
        p.font.color.rgb = RGBColor(*self.template_colors['light_blue'])
        p.alignment = PP_ALIGN.CENTER
        
        # ===== DECORATIVE ELEMENTS =====
        # Bottom right accent
        corner_accent = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(7.5), Inches(6),
            Inches(2), Inches(1)
        )
        corner_accent.fill.solid()
        corner_accent.fill.fore_color.rgb = RGBColor(*self.template_colors['light_blue'])
        corner_accent.fill.fore_color.brightness = 0.5  # Make it semi-transparent effect
        corner_accent.line.fill.background()
        
        return slide
    
    def highlight_keywords_in_text(self, text_frame, template_colors):
        """
        Apply color highlighting to keywords in text
        
        Args:
            text_frame: PowerPoint text frame object
            template_colors: Dictionary of template colors
        """
        for paragraph in text_frame.paragraphs:
            text = paragraph.text
            words = text.split()
            
            # Clear existing text
            paragraph.clear()
            
            # Rebuild with colored words
            for i, word in enumerate(words):
                word_lower = word.lower().strip('.,!?;:')
                
                # Add word with appropriate color
                run = paragraph.add_run()
                run.text = word + (' ' if i < len(words) - 1 else '')
                
                if word_lower in self.POSITIVE_KEYWORDS:
                    # Green for positive
                    run.font.color.rgb = RGBColor(46, 125, 50)  # Success green
                    run.font.bold = True
                elif word_lower in self.NEGATIVE_KEYWORDS:
                    # Red for negative
                    run.font.color.rgb = RGBColor(211, 47, 47)  # Warning red
                    run.font.bold = True
                elif word_lower in self.NEUTRAL_KEYWORDS:
                    # Blue for neutral
                    run.font.color.rgb = RGBColor(*template_colors.get('light_blue', (79, 129, 189)))
                    run.font.italic = True
    
    def add_metric_icon(self, slide, left, top, metric_type, size=0.3):
        """
        Add visual icon next to metrics
        
        Args:
            slide: PowerPoint slide object
            left: Left position in inches
            top: Top position in inches
            metric_type: Type of metric ('growth', 'money', etc.)
            size: Icon size in inches
        """
        # Use colored shapes as icons
        icon_colors = {
            'growth': (46, 125, 50),  # Green
            'decline': (211, 47, 47),  # Red
            'money': (255, 193, 7),  # Gold
            'business': (33, 150, 243),  # Blue
            'analytics': (156, 39, 176),  # Purple
            'target': (255, 87, 34),  # Orange
            'success': (76, 175, 80),  # Light green
            'warning': (255, 152, 0),  # Amber
        }
        
        # Get icon shape based on type
        icon_shapes = {
            'growth': MSO_SHAPE.UP_ARROW,
            'decline': MSO_SHAPE.DOWN_ARROW,
            'money': MSO_SHAPE.DIAMOND,
            'business': MSO_SHAPE.ROUNDED_RECTANGLE,
            'analytics': MSO_SHAPE.HEXAGON,
            'target': MSO_SHAPE.OVAL,
            'success': MSO_SHAPE.FLOWCHART_DECISION,
            'warning': MSO_SHAPE.ISOSCELES_TRIANGLE,
        }
        
        shape_type = icon_shapes.get(metric_type, MSO_SHAPE.OVAL)
        color = icon_colors.get(metric_type, self.template_colors['light_blue'])
        
        icon = slide.shapes.add_shape(
            shape_type,
            Inches(left), Inches(top),
            Inches(size), Inches(size)
        )
        icon.fill.solid()
        icon.fill.fore_color.rgb = RGBColor(*color)
        icon.line.color.rgb = RGBColor(*color)
        
        return icon
    
    def create_status_badge(self, slide, left, top, status_text, status_type='success'):
        """
        Create a colored status badge
        
        Args:
            slide: PowerPoint slide object
            left, top: Position in inches
            status_text: Text to display (e.g., "Excellent", "At Risk")
            status_type: 'success', 'warning', or 'danger'
        """
        badge_colors = {
            'success': (46, 125, 50),  # Green
            'warning': (255, 152, 0),  # Orange
            'danger': (211, 47, 47),  # Red
            'info': (33, 150, 243),  # Blue
        }
        
        color = badge_colors.get(status_type, (100, 100, 100))
        
        # Badge background
        badge = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left), Inches(top),
            Inches(1.5), Inches(0.4)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = RGBColor(*color)
        badge.line.fill.background()
        
        # Badge text
        text_frame = badge.text_frame
        text_frame.clear()
        p = text_frame.paragraphs[0]
        p.text = status_text
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        return badge
    
    def create_ai_insights_slide(self, prs, insights_list, analyst_notes=None):
        """
        Create dedicated AI Insights slide
        
        Args:
            prs: Presentation object
            insights_list: List of key insights (strings)
            analyst_notes: Optional analyst commentary
        """
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
        
        # ===== TITLE =====
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.3),
            Inches(9), Inches(0.6)
        )
        title_frame = title_box.text_frame
        p = title_frame.paragraphs[0]
        p.text = "🤖 AI-Powered Insights"
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*self.template_colors['navy'])
        
        # ===== INSIGHTS SECTION =====
        insights_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(1.2),
            Inches(9), Inches(4.5)
        )
        insights_frame = insights_box.text_frame
        insights_frame.word_wrap = True
        
        for i, insight in enumerate(insights_list[:6], 1):  # Max 6 insights
            p = insights_frame.add_paragraph() if i > 1 else insights_frame.paragraphs[0]
            
            # Add checkmark icon
            run = p.add_run()
            run.text = "✓ "
            run.font.size = Pt(18)
            run.font.color.rgb = RGBColor(46, 125, 50)  # Green
            run.font.bold = True
            
            # Add insight text
            run = p.add_run()
            run.text = insight
            run.font.size = Pt(16)
            run.font.color.rgb = RGBColor(*self.template_colors['dark_gray'])
            
            p.space_before = Pt(12)
        
        # ===== ANALYST NOTES SECTION =====
        if analyst_notes:
            notes_box = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(0.5), Inches(6),
                Inches(9), Inches(1)
            )
            notes_box.fill.solid()
            notes_box.fill.fore_color.rgb = RGBColor(*self.template_colors['light_gray'])
            notes_box.line.color.rgb = RGBColor(*self.template_colors['gray'])
            notes_box.line.width = Pt(1)
            
            notes_frame = notes_box.text_frame
            notes_frame.word_wrap = True
            notes_frame.margin_left = Inches(0.2)
            notes_frame.margin_right = Inches(0.2)
            notes_frame.margin_top = Inches(0.1)
            
            # Notes title
            p = notes_frame.paragraphs[0]
            run = p.add_run()
            run.text = "📝 Analyst Notes: "
            run.font.size = Pt(14)
            run.font.bold = True
            run.font.color.rgb = RGBColor(*self.template_colors['navy'])
            
            # Notes content
            run = p.add_run()
            run.text = analyst_notes
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(*self.template_colors['dark_gray'])
        
        return slide
    
    def add_performance_meter(self, slide, left, top, value, max_value=100, label="Performance"):
        """
        Add a visual performance meter (progress bar style)
        
        Args:
            slide: PowerPoint slide object
            left, top: Position in inches
            value: Current value
            max_value: Maximum value (for percentage calculation)
            label: Label text
        """
        width = 3.0
        height = 0.4
        
        # Background bar
        bg_bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left), Inches(top + 0.3),
            Inches(width), Inches(height)
        )
        bg_bar.fill.solid()
        bg_bar.fill.fore_color.rgb = RGBColor(*self.template_colors['light_gray'])
        bg_bar.line.fill.background()
        
        # Progress bar
        percentage = min(value / max_value, 1.0)
        progress_width = width * percentage
        
        # Color based on performance
        if percentage >= 0.8:
            color = (46, 125, 50)  # Green
        elif percentage >= 0.5:
            color = (255, 193, 7)  # Yellow
        else:
            color = (211, 47, 47)  # Red
        
        if progress_width > 0.1:  # Only show if > 10%
            progress_bar = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(left), Inches(top + 0.3),
                Inches(progress_width), Inches(height)
            )
            progress_bar.fill.solid()
            progress_bar.fill.fore_color.rgb = RGBColor(*color)
            progress_bar.line.fill.background()
        
        # Label
        label_box = slide.shapes.add_textbox(
            Inches(left), Inches(top),
            Inches(width), Inches(0.3)
        )
        label_frame = label_box.text_frame
        p = label_frame.paragraphs[0]
        p.text = f"{label}: {value:.1f}%"
        p.font.size = Pt(12)
        p.font.color.rgb = RGBColor(*self.template_colors['dark_gray'])
        
        return bg_bar, progress_bar if progress_width > 0.1 else None
