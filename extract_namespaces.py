#!/usr/bin/env python3
"""
Extract Namespaces from Word Document
Find all namespace declarations in the original Word document.
"""

import zipfile
import re

def extract_namespaces(filename):
    """Extract all namespace declarations from a Word document."""
    
    print(f"🔍 Extracting namespaces from: {filename}")
    
    try:
        with zipfile.ZipFile(filename, 'r') as docx_zip:
            document_xml = docx_zip.read('word/document.xml').decode('utf-8')
            
            # Find the root element with all namespace declarations
            root_match = re.search(r'<w:document[^>]+>', document_xml)
            if root_match:
                root_element = root_match.group(0)
                print(f"📄 Root element: {root_element}")
                
                # Extract all namespace declarations
                ns_pattern = r'xmlns:([^=]+)="([^"]+)"'
                namespaces = re.findall(ns_pattern, root_element)
                
                print(f"\n📊 Found {len(namespaces)} namespace declarations:")
                for prefix, uri in namespaces:
                    print(f"   {prefix} -> {uri}")
                
                # Generate registration code
                print(f"\n💻 ElementTree registration code:")
                print("# Register all original namespaces")
                for prefix, uri in namespaces:
                    print(f"ET.register_namespace('{prefix}', '{uri}')")
                
                return namespaces
            else:
                print("❌ Could not find root element")
                return []
                
    except Exception as e:
        print(f"❌ Error extracting namespaces: {e}")
        return []

if __name__ == "__main__":
    extract_namespaces("standaardofferte Compufit NL.docx")