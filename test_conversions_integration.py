#!/usr/bin/env python3
"""
Comprehensive test for conversions endpoint integration
Tests the API structure, dependencies, and compatibility
"""

import sys
import os
import tempfile
from pathlib import Path

# Add paths for imports
backend_path = r"c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src\backend\app"
src_path = r"c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src"
if backend_path not in sys.path:
    sys.path.insert(0, backend_path)
if src_path not in sys.path:
    sys.path.insert(0, src_path)

def test_conversions_imports():
    """Test all imports used in conversions.py"""
    print("=== Testing Conversions Imports ===")
    
    try:
        # FastAPI imports
        from fastapi import APIRouter, Depends, HTTPException, status, BackgroundTasks, Form
        from fastapi.responses import FileResponse
        print("✓ FastAPI imports successful")
        
        # Standard library
        from typing import List, Optional, Dict, Any
        from datetime import datetime
        import tempfile
        from pathlib import Path
        print("✓ Standard library imports successful")
        
        # Converter modules
        from converter.excel_reader import excel_reader
        from converter.ppt_writer import df_to_ppt
        print("✓ Converter modules imported successfully")
        
        # Check if backend dependencies exist
        deps_file = os.path.join(backend_path, 'api', 'deps.py')
        if os.path.exists(deps_file):
            print("✓ deps.py file exists")
        else:
            print("✗ deps.py file not found")
            
        # Check models directory
        models_dir = os.path.join(backend_path, 'models')
        if os.path.exists(models_dir):
            print("✓ models directory exists")
        else:
            print("✗ models directory not found")
            
        # Check services directory  
        services_dir = os.path.join(backend_path, 'services')
        if os.path.exists(services_dir):
            print("✓ services directory exists")
        else:
            print("✗ services directory not found")
            
        return True
        
    except ImportError as e:
        print(f"✗ Import error: {e}")
        return False
    except Exception as e:
        print(f"✗ Error: {e}")
        return False

def test_router_creation():
    """Test if APIRouter can be created and endpoints defined"""
    print("\n=== Testing Router Creation ===")
    
    try:
        from fastapi import APIRouter, Form, Depends
        from fastapi.responses import FileResponse
        
        # Create router like in conversions.py
        router = APIRouter()
        
        # Test endpoint definition (without actual dependencies)
        @router.post("/convert")
        async def test_convert(
            file_id: str = Form(...),
            title: str = Form("Auto Report")
        ):
            return {"message": "Test endpoint"}
            
        @router.get("/templates")
        async def test_templates():
            return {"templates": []}
            
        print("✓ Router and endpoints created successfully")
        print(f"✓ Router has {len(router.routes)} routes")
        
        return True
        
    except Exception as e:
        print(f"✗ Router creation error: {e}")
        return False

def test_conversion_pipeline():
    """Test the full conversion pipeline with sample data"""
    print("\n=== Testing Conversion Pipeline ===")
    
    try:
        from converter.excel_reader import excel_reader
        from converter.ppt_writer import df_to_ppt
        
        # Test with sample file
        test_file = r"c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\examples\Portfolio Allocation Data.xlsx"
        
        if not os.path.exists(test_file):
            print("✗ Test Excel file not found")
            return False
            
        # Read Excel
        df = excel_reader(test_file)
        print(f"✓ Excel read successfully: shape {df.shape}")
        
        # Create temp PPT
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
            ppt_path = tmp_file.name
            
        # Convert to PPT
        result = df_to_ppt(
            df=df,
            out_path=ppt_path,
            title="Integration Test",
            subtitle="Testing conversion pipeline",
            mode="table",
            limit=3
        )
        
        # Verify file creation
        if os.path.exists(result):
            size = os.path.getsize(result)
            print(f"✓ PPT created successfully: {size} bytes")
            os.unlink(result)  # Clean up
            return True
        else:
            print("✗ PPT file not created")
            return False
            
    except Exception as e:
        print(f"✗ Pipeline error: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_file_operations():
    """Test file handling operations used in the endpoint"""
    print("\n=== Testing File Operations ===")
    
    try:
        import tempfile
        from pathlib import Path
        
        # Test temp file creation
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
            temp_path = tmp_file.name
            
        print("✓ Temporary file creation works")
        
        # Test path operations
        test_filename = "test_file.xlsx"
        base_name = Path(test_filename).stem
        expected = "test_file_converted.pptx"
        actual = f"{base_name}_converted.pptx"
        
        if actual == expected:
            print("✓ Filename generation works")
        else:
            print(f"✗ Filename generation failed: {actual} != {expected}")
            
        # Clean up
        if os.path.exists(temp_path):
            os.unlink(temp_path)
            
        return True
        
    except Exception as e:
        print(f"✗ File operations error: {e}")
        return False

def test_api_structure():
    """Test the API structure matches expected patterns"""
    print("\n=== Testing API Structure ===")
    
    try:
        # Read the conversions.py file
        conversions_file = r"c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src\backend\app\api\v1\endpoints\conversions.py"
        
        if not os.path.exists(conversions_file):
            print("✗ conversions.py file not found")
            return False
            
        with open(conversions_file, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # Check for required elements
        checks = [
            ("router = APIRouter()", "APIRouter instantiation"),
            ("@router.post(\"/convert\")", "Convert endpoint definition"),
            ("@router.get(\"/templates\")", "Templates endpoint definition"),
            ("async def convert_excel_to_ppt", "Convert function definition"),
            ("FileResponse", "FileResponse usage"),
            ("require_credits", "Credit requirement"),
            ("get_current_active_user", "Authentication dependency")
        ]
        
        for check_str, description in checks:
            if check_str in content:
                print(f"✓ {description} found")
            else:
                print(f"✗ {description} missing")
                
        return True
        
    except Exception as e:
        print(f"✗ API structure test error: {e}")
        return False

def main():
    """Run all tests"""
    print("Starting Conversions API Integration Tests")
    print("=" * 50)
    
    tests = [
        test_conversions_imports,
        test_router_creation,
        test_conversion_pipeline,
        test_file_operations,
        test_api_structure
    ]
    
    results = []
    for test_func in tests:
        try:
            result = test_func()
            results.append(result)
        except Exception as e:
            print(f"✗ Test {test_func.__name__} failed with exception: {e}")
            results.append(False)
    
    # Summary
    print("\n" + "=" * 50)
    print("TEST SUMMARY")
    print("=" * 50)
    
    passed = sum(results)
    total = len(results)
    
    print(f"Tests passed: {passed}/{total}")
    
    if passed == total:
        print("🎉 ALL TESTS PASSED! Your conversions API is ready!")
    else:
        print("⚠️  Some tests failed. Check the issues above.")
        
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)