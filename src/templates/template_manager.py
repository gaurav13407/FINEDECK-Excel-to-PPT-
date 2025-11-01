"""
Template Manager for Professional PowerPoint Presentations
Handles loading, validation, and management of presentation templates
"""

import json
import os
from pathlib import Path
from typing import Dict, List, Optional, Any
import shutil
from datetime import datetime


class TemplateManager:
    """Manages presentation templates (built-in and custom)"""
    
    def __init__(self, templates_dir: str = "templates"):
        self.templates_dir = Path(templates_dir)
        self.built_in_dir = self.templates_dir / "built_in"
        self.custom_dir = self.templates_dir / "custom"
        
        # Ensure directories exist
        self.built_in_dir.mkdir(parents=True, exist_ok=True)
        self.custom_dir.mkdir(parents=True, exist_ok=True)
    
    def load_template(self, template_name: str) -> Optional[Dict[str, Any]]:
        """
        Load a template by name (checks built-in first, then custom)
        
        Args:
            template_name: Name of template (with or without .json)
        
        Returns:
            Template configuration dictionary or None if not found
        """
        if not template_name.endswith('.json'):
            template_name = f"{template_name}.json"
        
        # Check built-in templates first
        built_in_path = self.built_in_dir / template_name
        if built_in_path.exists():
            try:
                with open(built_in_path, 'r', encoding='utf-8') as f:
                    template = json.load(f)
                    template['source'] = 'built-in'
                    template['path'] = str(built_in_path)
                    return template
            except Exception as e:
                print(f"Error loading built-in template {template_name}: {e}")
                return None
        
        # Check custom templates
        custom_path = self.custom_dir / template_name
        if custom_path.exists():
            try:
                with open(custom_path, 'r', encoding='utf-8') as f:
                    template = json.load(f)
                    template['source'] = 'custom'
                    template['path'] = str(custom_path)
                    return template
            except Exception as e:
                print(f"Error loading custom template {template_name}: {e}")
                return None
        
        print(f"Template '{template_name}' not found")
        return None
    
    def list_templates(self, include_custom: bool = True) -> List[Dict[str, str]]:
        """
        List all available templates
        
        Args:
            include_custom: Whether to include custom templates
        
        Returns:
            List of template info dictionaries
        """
        templates = []
        
        # Built-in templates
        for template_file in self.built_in_dir.glob("*.json"):
            try:
                with open(template_file, 'r', encoding='utf-8') as f:
                    template = json.load(f)
                    templates.append({
                        'id': template_file.stem,
                        'name': template.get('name', template_file.stem),
                        'description': template.get('description', ''),
                        'category': template.get('category', 'general'),
                        'source': 'built-in',
                        'path': str(template_file)
                    })
            except Exception as e:
                print(f"Error reading template {template_file}: {e}")
        
        # Custom templates
        if include_custom:
            for template_file in self.custom_dir.glob("*.json"):
                try:
                    with open(template_file, 'r', encoding='utf-8') as f:
                        template = json.load(f)
                        templates.append({
                            'id': template_file.stem,
                            'name': template.get('name', template_file.stem),
                            'description': template.get('description', 'Custom template'),
                            'category': template.get('category', 'custom'),
                            'source': 'custom',
                            'path': str(template_file)
                        })
                except Exception as e:
                    print(f"Error reading custom template {template_file}: {e}")
        
        return sorted(templates, key=lambda x: (x['source'], x['category'], x['name']))
    
    def save_custom_template(self, template_data: Dict[str, Any], name: str, overwrite: bool = False) -> bool:
        """
        Save a custom template
        
        Args:
            template_data: Template configuration dictionary
            name: Template name (without .json)
            overwrite: Whether to overwrite existing template
        
        Returns:
            True if successful, False otherwise
        """
        if not name.endswith('.json'):
            name = f"{name}.json"
        
        template_path = self.custom_dir / name
        
        # Check if exists and overwrite not allowed
        if template_path.exists() and not overwrite:
            print(f"Template '{name}' already exists. Use overwrite=True to replace.")
            return False
        
        # Validate template
        if not self.validate_template(template_data):
            print("Template validation failed")
            return False
        
        # Add metadata
        template_data['created_at'] = datetime.now().isoformat()
        template_data['source'] = 'custom'
        
        try:
            with open(template_path, 'w', encoding='utf-8') as f:
                json.dump(template_data, f, indent=2, ensure_ascii=False)
            print(f"✅ Custom template '{name}' saved successfully")
            return True
        except Exception as e:
            print(f"❌ Error saving template: {e}")
            return False
    
    def delete_custom_template(self, name: str) -> bool:
        """
        Delete a custom template (built-in templates cannot be deleted)
        
        Args:
            name: Template name
        
        Returns:
            True if successful, False otherwise
        """
        if not name.endswith('.json'):
            name = f"{name}.json"
        
        template_path = self.custom_dir / name
        
        if not template_path.exists():
            print(f"Template '{name}' not found")
            return False
        
        try:
            template_path.unlink()
            print(f"✅ Template '{name}' deleted")
            return True
        except Exception as e:
            print(f"❌ Error deleting template: {e}")
            return False
    
    def validate_template(self, template_data: Dict[str, Any]) -> bool:
        """
        Validate template structure and required fields
        
        Args:
            template_data: Template configuration to validate
        
        Returns:
            True if valid, False otherwise
        """
        required_sections = ['name', 'colors', 'fonts', 'layout', 'styling']
        
        # Check required sections exist
        for section in required_sections:
            if section not in template_data:
                print(f"❌ Missing required section: {section}")
                return False
        
        # Validate colors section
        required_colors = ['primary', 'secondary', 'accent', 'background', 'text']
        for color in required_colors:
            if color not in template_data['colors']:
                print(f"❌ Missing required color: {color}")
                return False
            if not self._is_valid_hex_color(template_data['colors'][color]):
                print(f"❌ Invalid hex color for {color}: {template_data['colors'][color]}")
                return False
        
        # Validate fonts section
        required_fonts = ['title', 'subtitle', 'heading', 'body']
        for font_type in required_fonts:
            if font_type not in template_data['fonts']:
                print(f"❌ Missing required font: {font_type}")
                return False
            font_config = template_data['fonts'][font_type]
            if 'name' not in font_config or 'size' not in font_config:
                print(f"❌ Font '{font_type}' missing name or size")
                return False
        
        # Validate layout section
        if 'title_slide' not in template_data['layout'] or 'content_slide' not in template_data['layout']:
            print("❌ Layout must include 'title_slide' and 'content_slide'")
            return False
        
        # Validate styling section
        required_styling = ['table_header_color', 'table_alt_row_color']
        for style in required_styling:
            if style not in template_data['styling']:
                print(f"❌ Missing required styling: {style}")
                return False
        
        return True
    
    def _is_valid_hex_color(self, color: str) -> bool:
        """Check if string is valid hex color"""
        if not isinstance(color, str):
            return False
        if not color.startswith('#'):
            return False
        if len(color) not in [4, 7]:  # #RGB or #RRGGBB
            return False
        try:
            int(color[1:], 16)
            return True
        except ValueError:
            return False
    
    def duplicate_template(self, source_name: str, new_name: str) -> bool:
        """
        Duplicate a template to create a custom version
        
        Args:
            source_name: Name of template to duplicate
            new_name: Name for the new template
        
        Returns:
            True if successful
        """
        source_template = self.load_template(source_name)
        if not source_template:
            return False
        
        # Remove metadata from source
        source_template.pop('source', None)
        source_template.pop('path', None)
        source_template.pop('created_at', None)
        
        # Update name
        source_template['name'] = new_name
        source_template['description'] = f"Custom template based on {source_name}"
        
        return self.save_custom_template(source_template, new_name)
    
    def get_template_colors(self, template_name: str) -> Optional[List[str]]:
        """
        Get chart colors from a template
        
        Args:
            template_name: Template name
        
        Returns:
            List of hex color codes for charts
        """
        template = self.load_template(template_name)
        if template and 'colors' in template and 'chart_colors' in template['colors']:
            return template['colors']['chart_colors']
        return None
    
    def export_template(self, template_name: str, output_path: str) -> bool:
        """
        Export a template to a file
        
        Args:
            template_name: Template to export
            output_path: Destination file path
        
        Returns:
            True if successful
        """
        template = self.load_template(template_name)
        if not template:
            return False
        
        # Remove internal metadata
        template.pop('source', None)
        template.pop('path', None)
        
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(template, f, indent=2, ensure_ascii=False)
            print(f"✅ Template exported to {output_path}")
            return True
        except Exception as e:
            print(f"❌ Error exporting template: {e}")
            return False
    
    def import_template(self, template_path: str, name: Optional[str] = None) -> bool:
        """
        Import a template from a file
        
        Args:
            template_path: Path to template JSON file
            name: Optional new name for the template
        
        Returns:
            True if successful
        """
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template_data = json.load(f)
            
            if not self.validate_template(template_data):
                return False
            
            # Use provided name or template's name or filename
            if name:
                template_name = name
            elif 'name' in template_data:
                template_name = template_data['name'].lower().replace(' ', '_')
            else:
                template_name = Path(template_path).stem
            
            return self.save_custom_template(template_data, template_name)
        
        except Exception as e:
            print(f"❌ Error importing template: {e}")
            return False


# Convenience functions
def get_default_template() -> str:
    """Get the default template name"""
    return "corporate_blue"


def hex_to_rgb(hex_color: str) -> tuple:
    """
    Convert hex color to RGB tuple
    
    Args:
        hex_color: Hex color string (e.g., "#1F4788")
    
    Returns:
        RGB tuple (r, g, b)
    """
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 3:
        hex_color = ''.join([c*2 for c in hex_color])
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


if __name__ == "__main__":
    # Demo usage
    manager = TemplateManager()
    
    print("="*70)
    print("📋 TEMPLATE MANAGER DEMO")
    print("="*70 + "\n")
    
    # List all templates
    print("Available Templates:")
    print("-" * 70)
    templates = manager.list_templates()
    for i, template in enumerate(templates, 1):
        print(f"{i:2}. [{template['source']:8}] {template['name']:25} - {template['category']}")
        print(f"    {template['description'][:60]}")
    
    print(f"\n✅ Total: {len(templates)} templates loaded")
    print("="*70)
