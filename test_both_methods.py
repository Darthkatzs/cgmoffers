#!/usr/bin/env python3
"""
Test Both Methods
Compare direct processor vs server results.
"""

from content_control_processor import ContentControlProcessor
import json
import requests

def test_direct_method():
    """Test direct content processor."""
    
    processor = ContentControlProcessor()
    
    test_data = {
        "companyName": "COMPARISON TEST",
        "contactName": "Test User",
        "address": "Test Street 123",
        "postalCode": "1234AB",
        "city": "Test City",
        "companyId": "TEST123",
        "description": "Comparison test",
        "oneTimeCosts": [
            {
                "material": "Direct Test Module",
                "quantity": 2,
                "unitPrice": 500.00,
                "total": 1000.00
            }
        ],
        "recurringCosts": [
            {
                "material": "Direct Test Support",
                "quantity": 12,
                "unitPrice": 75.00,
                "total": 900.00
            }
        ]
    }
    
    print("🧪 Testing direct method...")
    success = processor.process_word_template(
        "standaardofferte Compufit NL.docx",
        test_data,
        "direct_method_test.docx"
    )
    
    return success, test_data

def test_server_method(test_data):
    """Test server method with same data."""
    
    print("🌐 Testing server method...")
    
    try:
        response = requests.post(
            'http://localhost:8001/generate-quotation',
            headers={'Content-Type': 'application/json'},
            json=test_data,
            timeout=10
        )
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                return True, result.get('filename')
            else:
                print(f"❌ Server error: {result.get('error')}")
                return False, None
        else:
            print(f"❌ HTTP error: {response.status_code}")
            return False, None
            
    except Exception as e:
        print(f"❌ Request error: {e}")
        return False, None

def check_table_fields(filename, method_name):
    """Check table field values in a document."""
    
    print(f"\\n📋 Checking table fields in {method_name} result:")
    
    import zipfile
    import xml.etree.ElementTree as ET
    
    try:
        with zipfile.ZipFile(filename, 'r') as docx_zip:
            document_xml = docx_zip.read('word/document.xml').decode('utf-8')
            root = ET.fromstring(document_xml)
            w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            
            table_fields = ['Module', 'Aantal', 'éénmalige setupkost', 'calctotaalsetup', 'Jaarlijks', 'calctotaaljaarlijks']
            
            results = {}
            for sdt in root.iter(f'{w_ns}sdt'):
                try:
                    sdt_pr = sdt.find(f'{w_ns}sdtPr')
                    if sdt_pr is not None:
                        control_name = None
                        alias_elem = sdt_pr.find(f'{w_ns}alias')
                        if alias_elem is not None:
                            control_name = alias_elem.get(f'{w_ns}val')
                        if not control_name:
                            tag_elem = sdt_pr.find(f'{w_ns}tag')
                            if tag_elem is not None:
                                control_name = tag_elem.get(f'{w_ns}val')
                        
                        if control_name in table_fields:
                            sdt_content = sdt.find(f'{w_ns}sdtContent')
                            content_text = ''
                            if sdt_content is not None:
                                for t in sdt_content.iter(f'{w_ns}t'):
                                    if t.text:
                                        content_text += t.text
                            
                            if control_name not in results:
                                results[control_name] = []
                            results[control_name].append(content_text)
                except:
                    continue
            
            for field in table_fields:
                values = results.get(field, ['[NOT FOUND]'])
                if len(values) == 1:
                    print(f"   {field}: '{values[0]}'")
                else:
                    print(f"   {field} (x{len(values)}): {values}")
                    
            return results
            
    except Exception as e:
        print(f"❌ Error checking {filename}: {e}")
        return {}

def main():
    """Compare both methods."""
    
    # Test direct method
    direct_success, test_data = test_direct_method()
    
    if direct_success:
        direct_results = check_table_fields("direct_method_test.docx", "DIRECT")
    else:
        print("❌ Direct method failed")
        return
    
    # Test server method
    server_success, server_filename = test_server_method(test_data)
    
    if server_success and server_filename:
        server_results = check_table_fields(server_filename, "SERVER")
    else:
        print("❌ Server method failed")
        return
    
    # Compare results
    print(f"\\n🔍 COMPARISON:")
    table_fields = ['Module', 'Aantal', 'éénmalige setupkost', 'calctotaalsetup', 'Jaarlijks', 'calctotaaljaarlijks']
    
    for field in table_fields:
        direct_val = direct_results.get(field, [''])[0]
        server_val = server_results.get(field, [''])[0]
        
        if direct_val == server_val:
            print(f"   ✅ {field}: Both have '{direct_val}'")
        else:
            print(f"   ❌ {field}: Direct='{direct_val}', Server='{server_val}'")

if __name__ == "__main__":
    main()