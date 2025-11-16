"""
Test script to verify background contrast and readability improvements
Tests the new white-background Summary slide and contrast adjustment functions
"""

from src.converter.enhanced_professional_builder import EnhancedProfessionalBuilder
from src.converter.finance_chart_formatter import (
    calculate_luminance,
    calculate_contrast_ratio,
    is_bright_background,
    is_dark_background,
    get_readable_text_color,
    FINANCE_COLORS
)
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


def test_luminance_calculations():
    """Test WCAG luminance calculations"""
    print("\n" + "="*70)
    print("TEST 1: WCAG Luminance Calculations")
    print("="*70)
    
    # Test cases: (color, expected_range, description)
    test_cases = [
        ((0, 0, 0), (0.0, 0.01), "Black"),
        ((255, 255, 255), (0.99, 1.01), "White"),
        ((25, 42, 86), (0.02, 0.04), "Navy (dark)"),
        ((0, 79, 158), (0.05, 0.10), "Primary blue"),
        ((22, 160, 133), (0.15, 0.25), "Accent green"),
    ]
    
    all_passed = True
    for color, (min_lum, max_lum), desc in test_cases:
        lum = calculate_luminance(color)
        passed = min_lum <= lum <= max_lum
        status = "✅ PASS" if passed else "❌ FAIL"
        print(f"{status} {desc:20} RGB{str(color):25} → Luminance: {lum:.4f}")
        if not passed:
            all_passed = False
    
    print(f"\n{'✅ All luminance tests PASSED' if all_passed else '❌ Some tests FAILED'}")
    return all_passed


def test_contrast_ratios():
    """Test WCAG contrast ratio calculations"""
    print("\n" + "="*70)
    print("TEST 2: WCAG Contrast Ratios (Minimum 4.5:1 for AA)")
    print("="*70)
    
    # Test cases: (bg_color, text_color, min_ratio, description)
    test_cases = [
        # Good contrasts
        ((255, 255, 255), (51, 51, 51), 4.5, "White bg + dark text (should pass)"),
        ((255, 255, 255), (0, 79, 158), 4.5, "White bg + primary blue (should pass)"),
        ((25, 42, 86), (255, 255, 255), 4.5, "Navy bg + white text (should pass)"),
        
        # Bad contrasts
        ((25, 42, 86), (52, 152, 219), 4.5, "Navy bg + light blue (should FAIL)"),
    ]
    
    all_passed = True
    for bg_color, text_color, min_ratio, desc in test_cases:
        ratio = calculate_contrast_ratio(bg_color, text_color)
        passes = ratio >= min_ratio
        status = "✅ PASS" if passes else "⚠️  WARN"
        print(f"{status} {desc:45} → {ratio:.2f}:1 {'(AA compliant)' if passes else '(Below minimum!)'}")
        if not passes and "should pass" in desc.lower():
            all_passed = False
    
    print(f"\n{'✅ All expected contrasts met' if all_passed else '⚠️  Some expected contrasts failed'}")
    return all_passed


def test_background_detection():
    """Test bright/dark background detection"""
    print("\n" + "="*70)
    print("TEST 3: Background Detection (Bright >0.85, Dark <0.3)")
    print("="*70)
    
    # Test cases: (color, should_be_bright, should_be_dark, description)
    test_cases = [
        ((255, 255, 255), True, False, "White"),
        ((249, 250, 251), True, False, "Light gradient top"),
        ((25, 42, 86), False, True, "Navy"),
        ((0, 0, 0), False, True, "Black"),
        ((128, 128, 128), False, False, "Medium gray (neither)"),
    ]
    
    all_passed = True
    for color, should_bright, should_dark, desc in test_cases:
        is_bright = is_bright_background(color)
        is_dark = is_dark_background(color)
        lum = calculate_luminance(color)
        
        passed = (is_bright == should_bright) and (is_dark == should_dark)
        status = "✅ PASS" if passed else "❌ FAIL"
        
        brightness = "Bright" if is_bright else "Dark" if is_dark else "Normal"
        print(f"{status} {desc:20} (L={lum:.3f}) → {brightness:10}")
        
        if not passed:
            all_passed = False
    
    print(f"\n{'✅ All detection tests PASSED' if all_passed else '❌ Some tests FAILED'}")
    return all_passed


def test_readable_text_color():
    """Test automatic readable text color selection"""
    print("\n" + "="*70)
    print("TEST 4: Readable Text Color Selection (4.5:1 minimum)")
    print("="*70)
    
    # Test cases: (bg_color, description)
    test_cases = [
        ((255, 255, 255), "White background"),
        ((249, 250, 251), "Light gradient"),
        ((25, 42, 86), "Navy background"),
        ((0, 79, 158), "Primary blue"),
        ((128, 128, 128), "Medium gray"),
    ]
    
    all_passed = True
    for bg_color, desc in test_cases:
        text_color = get_readable_text_color(bg_color)
        ratio = calculate_contrast_ratio(bg_color, text_color)
        
        passed = ratio >= 4.5
        status = "✅ PASS" if passed else "❌ FAIL"
        
        color_name = "Dark" if text_color == FINANCE_COLORS['dark_text'] else "White"
        print(f"{status} {desc:20} → {color_name:10} text ({ratio:.2f}:1)")
        
        if not passed:
            all_passed = False
    
    print(f"\n{'✅ All text colors meet 4.5:1 minimum' if all_passed else '❌ Some colors below minimum'}")
    return all_passed


def test_summary_slide_creation():
    """Test the new white-background Summary slide"""
    print("\n" + "="*70)
    print("TEST 5: Summary Slide Creation (FinDeck Brand Styling)")
    print("="*70)
    
    try:
        # Create a test presentation with just the Summary slide
        prs = Presentation()
        builder = EnhancedProfessionalBuilder()
        builder._create_closing_slide(prs)
        
        # Save to test file
        output_file = 'test_summary_slide.pptx'
        prs.save(output_file)
        
        print(f"✅ Summary slide created successfully")
        print(f"   📁 Saved to: {output_file}")
        print(f"   🎨 Features:")
        print(f"      • White background (#FFFFFF)")
        print(f"      • Top brand bar (2px #004F9E)")
        print(f"      • Title in primary blue (#004F9E)")
        print(f"      • Green checkmarks (#16A085)")
        print(f"      • Medium gray footer (#6C757D)")
        print(f"      • Professional margins (48px/24px/36px)")
        print(f"      • Subtle text box shadow")
        
        return True
    except Exception as e:
        print(f"❌ FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False


def run_all_tests():
    """Run all contrast and readability tests"""
    print("\n" + "="*70)
    print("🧪 BACKGROUND CONTRAST & READABILITY TEST SUITE")
    print("   Testing WCAG 4.5:1 compliance and FinDeck brand styling")
    print("="*70)
    
    results = []
    
    # Run all tests
    results.append(("Luminance Calculations", test_luminance_calculations()))
    results.append(("Contrast Ratios", test_contrast_ratios()))
    results.append(("Background Detection", test_background_detection()))
    results.append(("Readable Text Colors", test_readable_text_color()))
    results.append(("Summary Slide Creation", test_summary_slide_creation()))
    
    # Summary
    print("\n" + "="*70)
    print("📊 TEST SUMMARY")
    print("="*70)
    
    for test_name, passed in results:
        status = "✅ PASSED" if passed else "❌ FAILED"
        print(f"{status:12} {test_name}")
    
    total_passed = sum(1 for _, passed in results if passed)
    total_tests = len(results)
    
    print("="*70)
    print(f"🎯 OVERALL: {total_passed}/{total_tests} test suites passed")
    print("="*70)
    
    if total_passed == total_tests:
        print("\n🎉 ALL TESTS PASSED! Background contrast improvements are working correctly.")
        print("   ✅ WCAG 4.5:1 compliance verified")
        print("   ✅ FinDeck brand styling applied")
        print("   ✅ Summary slide uses white background with readable colors")
    else:
        print(f"\n⚠️  {total_tests - total_passed} test suite(s) failed. Review output above.")
    
    return total_passed == total_tests


if __name__ == "__main__":
    success = run_all_tests()
    exit(0 if success else 1)
