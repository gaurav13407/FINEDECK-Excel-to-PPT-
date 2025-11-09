"""
Quick test to see what templates are being detected
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.templates.template_manager import TemplateManager

print("\n🔍 Testing Template Manager...")
print("="*70)

manager = TemplateManager()

print(f"\n📂 Templates directory: {manager.templates_dir}")
print(f"📂 Built-in directory: {manager.built_in_dir}")
print(f"📂 Custom directory: {manager.custom_dir}")

print(f"\n📋 Listing all templates...")
templates = manager.list_templates()

print(f"\n✅ Found {len(templates)} templates:")
for t in templates:
    print(f"   - {t['id']}: {t['name']} ({t['source']})")

print(f"\n🔍 Testing royal_purple specifically...")
royal_purple = manager.load_template('royal_purple')
if royal_purple:
    print(f"✅ royal_purple loaded successfully!")
    print(f"   Name: {royal_purple.get('name')}")
    print(f"   Source: {royal_purple.get('source')}")
else:
    print(f"❌ royal_purple NOT FOUND!")
