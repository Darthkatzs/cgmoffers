#!/usr/bin/env python3
"""
Test Table Processing
Test just the table field processing in isolation.
"""

from content_control_processor import ContentControlProcessor

def main():
    """Test table processing."""
    
    processor = ContentControlProcessor()
    
    test_data = {
        "companyName": "DEBUG TEST",
        "contactName": "Debug User", 
        "address": "Debug Address",
        "postalCode": "1111AB",
        "city": "Debug City",
        "companyId": "DEBUG123",
        "description": "Debug test",
        "oneTimeCosts": [
            {
                "material": "Test Module",
                "quantity": 3,
                "unitPrice": 200.00,
                "total": 600.00
            }
        ],
        "recurringCosts": [
            {
                "material": "Test Support", 
                "quantity": 6,
                "unitPrice": 50.00,
                "total": 300.00
            }
        ]
    }
    
    print("🧪 Testing table field processing...")
    
    success = processor.process_word_template(
        "standaardofferte Compufit NL.docx",
        test_data,
        "debug_table_test.docx"
    )
    
    if success:
        print("\\n✅ Document generated successfully")
        
        # Check the specific table fields
        import zipfile
        import xml.etree.ElementTree as ET
        
        with zipfile.ZipFile("debug_table_test.docx", 'r') as docx_zip:
            document_xml = docx_zip.read('word/document.xml').decode('utf-8')
            root = ET.fromstring(document_xml)
            w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            
            table_controls = ['Module', 'Aantal', 'éénmalige setupkost', 'calctotaalsetup', 'Jaarlijks', 'calctotaaljaarlijks']
            
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
                        
                        if control_name in table_controls:
                            sdt_content = sdt.find(f'{w_ns}sdtContent')
                            content_text = ''
                            if sdt_content is not None:
                                for t in sdt_content.iter(f'{w_ns}t'):
                                    if t.text:
                                        content_text += t.text
                            
                            expected = {
                                'Module': 'Test Module',
                                'Aantal': '3', 
                                'éénmalige setupkost': '€200.00',
                                'calctotaalsetup': '€600.00',
                                'Jaarlijks': '€50.00',
                                'calctotaaljaarlijks': '€300.00'
                            }
                            
                            exp_val = expected.get(control_name, 'N/A')
                            if content_text == exp_val:
                                print(f"   ✅ {control_name}: '{content_text}' (correct)")
                            else:
                                print(f"   ❌ {control_name}: Expected '{exp_val}', got '{content_text}'")
                except:
                    continue
    else:
        print("❌ Document generation failed")

if __name__ == "__main__":
    main()