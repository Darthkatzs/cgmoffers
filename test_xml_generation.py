#!/usr/bin/env python3
"""
Test XML Generation Issues
Check for potential XML namespace or structure issues that could cause Word compatibility problems.
"""

import zipfile
import xml.etree.ElementTree as ET
import tempfile
import shutil
import os

def compare_xml_namespaces(original_file, generated_file):
    """Compare XML namespace declarations between original and generated files."""
    
    print(f"🔍 Comparing XML namespaces...")
    
    def extract_document_xml(filename):
        with zipfile.ZipFile(filename, 'r') as z:
            return z.read('word/document.xml').decode('utf-8')
    
    try:
        orig_xml = extract_document_xml(original_file)
        gen_xml = extract_document_xml(generated_file)
        
        print(f"📄 Original XML starts with: {orig_xml[:200]}...")
        print(f"📄 Generated XML starts with: {gen_xml[:200]}...")
        
        # Check if namespaces match
        orig_root_start = orig_xml.split('>')[0] + '>'
        gen_root_start = gen_xml.split('>')[0] + '>'
        
        print(f"\n🔍 Original root element: {orig_root_start}")
        print(f"🔍 Generated root element: {gen_root_start}")
        
        if orig_root_start == gen_root_start:
            print("✅ Root elements match")
        else:
            print("❌ Root elements differ - this could cause compatibility issues")
            
        return orig_xml == gen_xml
            
    except Exception as e:
        print(f"❌ Error comparing XML: {e}")
        return False

def check_xml_formatting(filename):
    """Check for XML formatting issues."""
    
    print(f"🔍 Checking XML formatting in: {filename}")
    
    try:
        with zipfile.ZipFile(filename, 'r') as docx_zip:
            document_xml = docx_zip.read('word/document.xml').decode('utf-8')
            
            # Check for common XML issues
            issues = []
            
            # Check for proper XML declaration
            if not document_xml.startswith('<?xml version="1.0"'):
                issues.append("Missing or incorrect XML declaration")
            
            # Check for encoding declaration
            if 'encoding=' not in document_xml[:100]:
                issues.append("Missing encoding declaration")
            
            # Check for unclosed tags (basic check)
            open_tags = document_xml.count('<w:')
            close_tags = document_xml.count('</w:')
            self_closing = document_xml.count('/>')
            
            print(f"📊 XML statistics:")
            print(f"   Open tags: {open_tags}")
            print(f"   Close tags: {close_tags}")  
            print(f"   Self-closing: {self_closing}")
            
            # Simple balance check (not perfect but catches major issues)
            if abs(open_tags - close_tags - self_closing) > 10:  # Allow small variance
                issues.append(f"Potential unbalanced tags: {open_tags} open, {close_tags} close, {self_closing} self-closing")
            
            # Check for invalid characters
            if '\x00' in document_xml:
                issues.append("Contains null bytes")
                
            if issues:
                print("❌ XML issues found:")
                for issue in issues:
                    print(f"   • {issue}")
                return False
            else:
                print("✅ XML formatting appears correct")
                return True
                
    except Exception as e:
        print(f"❌ Error checking XML: {e}")
        return False

def test_minimal_generation():
    """Test generating a minimal document to isolate the issue."""
    
    print(f"\n🧪 Testing minimal document generation...")
    
    try:
        from content_control_processor import ContentControlProcessor
        
        # Very minimal test data
        minimal_data = {
            "companyName": "TEST",
            "contactName": "TEST",
            "address": "TEST",
            "postalCode": "1234AB",
            "city": "TEST",
            "companyId": "TEST",
            "description": "TEST",
            "oneTimeCosts": [],
            "recurringCosts": []
        }
        
        processor = ContentControlProcessor()
        success = processor.process_word_template(
            "standaardofferte Compufit NL.docx",
            minimal_data,
            "minimal_test.docx"
        )
        
        if success and os.path.exists("minimal_test.docx"):
            print("✅ Minimal generation successful")
            
            # Test opening with python-docx
            from docx import Document
            doc = Document("minimal_test.docx")
            print(f"✅ Minimal document can be opened ({len(doc.paragraphs)} paragraphs)")
            
            return True
        else:
            print("❌ Minimal generation failed")
            return False
            
    except Exception as e:
        print(f"❌ Error in minimal generation: {e}")
        return False

def main():
    """Run all XML generation tests."""
    
    original_file = "standaardofferte Compufit NL.docx"
    test_file = "Offerte_Test_Company_BV_20250905_215004.docx"
    
    if not os.path.exists(original_file):
        print(f"❌ Original template not found: {original_file}")
        return
        
    if not os.path.exists(test_file):
        print(f"❌ Test file not found: {test_file}")
        return
    
    print("="*60)
    print("XML GENERATION ANALYSIS")
    print("="*60)
    
    # Check XML formatting
    print("\n1. Checking original template XML...")
    check_xml_formatting(original_file)
    
    print("\n2. Checking generated file XML...")
    check_xml_formatting(test_file)
    
    print("\n3. Comparing namespaces...")
    compare_xml_namespaces(original_file, test_file)
    
    print("\n4. Testing minimal generation...")
    test_minimal_generation()

if __name__ == "__main__":
    main()