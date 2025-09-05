#!/usr/bin/env python3
"""
List All Controls
List all content control names found in the Word document.
"""

import zipfile
import xml.etree.ElementTree as ET

def list_controls(filename):
    """List all content controls in a Word document."""
    
    print(f"🔍 Listing all content controls in: {filename}")
    
    with zipfile.ZipFile(filename, 'r') as docx_zip:
        document_xml = docx_zip.read('word/document.xml').decode('utf-8')
        root = ET.fromstring(document_xml)
        w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        
        controls_found = []
        
        for sdt in root.iter(f'{w_ns}sdt'):
            try:
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
                        # Get current content
                        sdt_content = sdt.find(f'{w_ns}sdtContent')
                        content_text = ''
                        if sdt_content is not None:
                            for t in sdt_content.iter(f'{w_ns}t'):
                                if t.text:
                                    content_text += t.text
                        
                        controls_found.append((control_name, content_text))
            except:
                continue
        
        print(f"📊 Found {len(controls_found)} content controls:")
        
        # Group by control name to show duplicates
        control_counts = {}
        for control_name, content in controls_found:
            if control_name not in control_counts:
                control_counts[control_name] = []
            control_counts[control_name].append(content)
        
        for control_name in sorted(control_counts.keys()):
            contents = control_counts[control_name]
            if len(contents) == 1:
                print(f"   {control_name}: '{contents[0]}'")
            else:
                print(f"   {control_name} (x{len(contents)}):")
                for i, content in enumerate(contents):
                    print(f"      [{i+1}]: '{content}'")

if __name__ == "__main__":
    list_controls("Offerte_Table_Fields_Test_20250905_221155.docx")