#!/usr/bin/env python3
"""
Verify Content Updated
Check if the content controls are actually getting the correct values.
"""

import zipfile
import xml.etree.ElementTree as ET

def check_content_controls(filename, expected_values):
    """Check if content controls contain expected values."""
    
    print(f"🔍 Checking content controls in: {filename}")
    
    try:
        with zipfile.ZipFile(filename, 'r') as docx_zip:
            document_xml = docx_zip.read('word/document.xml').decode('utf-8')
            
            # Parse XML
            root = ET.fromstring(document_xml)
            w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            
            controls_found = {}
            
            # Find all content controls
            for sdt in root.iter(f'{w_ns}sdt'):
                try:
                    # Get control name
                    sdt_pr = sdt.find(f'{w_ns}sdtPr')
                    if sdt_pr is not None:
                        control_name = None
                        
                        # Try alias first, then tag
                        alias_elem = sdt_pr.find(f'{w_ns}alias')
                        if alias_elem is not None:
                            control_name = alias_elem.get(f'{w_ns}val')
                        
                        if not control_name:
                            tag_elem = sdt_pr.find(f'{w_ns}tag')
                            if tag_elem is not None:
                                control_name = tag_elem.get(f'{w_ns}val')
                        
                        if control_name:
                            # Get the content
                            sdt_content = sdt.find(f'{w_ns}sdtContent')
                            if sdt_content is not None:
                                content_text = ''
                                for t in sdt_content.iter(f'{w_ns}t'):
                                    if t.text:
                                        content_text += t.text
                                
                                controls_found[control_name] = content_text
                                
                except Exception as e:
                    continue
            
            print(f"📊 Found {len(controls_found)} content controls:")
            
            # Check against expected values
            matches = 0
            for control_name, expected_value in expected_values.items():
                actual_value = controls_found.get(control_name, '[NOT FOUND]')
                if actual_value == expected_value:
                    print(f"   ✅ {control_name}: '{actual_value}' (matches)")
                    matches += 1
                else:
                    print(f"   ❌ {control_name}: Expected '{expected_value}', got '{actual_value}'")
            
            print(f"\\n📈 Summary: {matches}/{len(expected_values)} controls match expected values")
            
            return matches == len(expected_values)
            
    except Exception as e:
        print(f"❌ Error checking controls: {e}")
        return False

def main():
    """Test content control values in generated files."""
    
    # Expected values for the test we just ran
    expected_values = {
        'praktijk': 'Test Company 2025',
        'naam': 'John Doe',
        'adres': 'Test Street 456',
        'postcode': '5678CD',
        'stad': 'Test City',
        'btw': 'NL123456789B01',
        'beschrijving': 'Test quotation for debugging',
        'totaaleenmalig': '500.00',
        'totaaljaarlijks': '600.00',
        'total': '1100.00',
        'vat': '231.00',
        'grandtotal': '1331.00'
    }
    
    test_files = [
        'Offerte_Test_Company_2025_20250905_215707.docx',
        'content_control_output.docx'
    ]
    
    for filename in test_files:
        try:
            print("\\n" + "="*60)
            print(f"TESTING: {filename}")
            print("="*60)
            
            success = check_content_controls(filename, expected_values)
            print(f"\\n{'✅ PASSED' if success else '❌ FAILED'}")
            
        except Exception as e:
            print(f"❌ Error testing {filename}: {e}")

if __name__ == "__main__":
    main()