#!/usr/bin/env python3
"""
Analyze Word Controls
Examines the Word document to find all controls, content controls, and form fields.
"""

import zipfile
import xml.etree.ElementTree as ET

def analyze_word_controls(docx_path):
    """Analyze all controls in the Word document."""
    
    print(f"🔍 Analyzing Word controls in: {docx_path}")
    print("=" * 60)
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx:
            # Files to analyze
            xml_files = [
                'word/document.xml',
                'word/header1.xml',
                'word/header2.xml',
                'word/header3.xml',
                'word/footer1.xml',
                'word/footer2.xml',
                'word/footer3.xml'
            ]
            
            all_controls = set()
            all_text_content = []
            
            for xml_file in xml_files:
                try:
                    xml_content = docx.read(xml_file).decode('utf-8')
                    
                    print(f"\n📄 Analyzing: {xml_file}")
                    
                    # Parse XML
                    root = ET.fromstring(xml_content)
                    
                    # Method 1: Look for structured document tags (content controls)
                    print("   🎯 Content Controls (SDT):")
                    sdt_count = 0
                    
                    # Register namespace
                    ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                    
                    # Find all SDT elements
                    for sdt in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'):
                        sdt_count += 1
                        
                        # Try to find the control name/alias
                        for sdt_pr in sdt.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtPr'):
                            # Look for alias
                            for alias in sdt_pr.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}alias'):
                                val = alias.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                                if val:
                                    all_controls.add(val)
                                    print(f"      - Alias: {val}")
                            
                            # Look for tag
                            for tag in sdt_pr.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag'):
                                val = tag.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                                if val:
                                    all_controls.add(val)
                                    print(f"      - Tag: {val}")
                    
                    if sdt_count == 0:
                        print("      None found")
                    
                    # Method 2: Look for form fields
                    print("   📝 Form Fields:")
                    form_field_count = 0
                    
                    for fld_simple in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldSimple'):
                        instr = fld_simple.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instr')
                        if instr:
                            all_controls.add(instr)
                            form_field_count += 1
                            print(f"      - Field: {instr}")
                    
                    if form_field_count == 0:
                        print("      None found")
                    
                    # Method 3: Extract all text content to look for patterns
                    print("   📋 All Text Content (looking for patterns):")
                    text_content = ""
                    for elem in root.iter():
                        if elem.tag.endswith('}t') and elem.text:  # Text runs
                            text_content += elem.text + " "
                    
                    all_text_content.append(text_content)
                    
                    # Look for potential control names in text
                    import re
                    
                    # Look for words that might be control names
                    words = re.findall(r'\b[a-zA-Z][a-zA-Z0-9_]*\b', text_content)
                    potential_controls = []
                    
                    # Filter for likely control names
                    for word in set(words):
                        if (len(word) > 3 and 
                            not word.lower() in ['the', 'and', 'for', 'are', 'but', 'not', 'you', 'all', 'can', 'had', 'her', 'was', 'one', 'our', 'out', 'day', 'get', 'has', 'him', 'his', 'how', 'man', 'new', 'now', 'old', 'see', 'two', 'way', 'who', 'boy', 'did', 'its', 'let', 'put', 'say', 'she', 'too', 'use'] and
                            not word.lower() in ['this', 'that', 'with', 'have', 'will', 'your', 'from', 'they', 'know', 'want', 'been', 'good', 'much', 'some', 'time', 'very', 'when', 'come', 'here', 'just', 'like', 'long', 'make', 'many', 'over', 'such', 'take', 'than', 'them', 'well', 'work']):
                            
                            # Look for specific patterns that suggest controls
                            if (word.lower() in ['praktijknaam', 'naam', 'straat', 'nummer', 'postcode', 'stad', 'btw', 'items1', 'items2'] or
                                'sig' in word.lower() or 
                                'block' in word.lower() or
                                len(word) > 8):
                                potential_controls.append(word)
                    
                    if potential_controls:
                        print(f"      Potential controls found: {', '.join(potential_controls[:10])}")
                        all_controls.update(potential_controls)
                    
                except FileNotFoundError:
                    continue
                except Exception as e:
                    print(f"   ❌ Error analyzing {xml_file}: {e}")
    
    except Exception as e:
        print(f"❌ Error opening document: {e}")
        return []
    
    # Summary
    print(f"\n" + "=" * 60)
    print("📊 SUMMARY")
    print("=" * 60)
    
    if all_controls:
        print(f"✅ Found {len(all_controls)} potential controls:")
        for control in sorted(all_controls):
            print(f"   - {control}")
    else:
        print("❌ No controls found")
    
    print(f"\n💡 FULL TEXT SAMPLE (first 500 chars):")
    full_text = " ".join(all_text_content)
    print(repr(full_text[:500]) + "...")
    
    return sorted(all_controls)

if __name__ == "__main__":
    controls = analyze_word_controls("standaardofferte Compufit NL.docx")
    
    if controls:
        print(f"\n🔧 SUGGESTED CONTROL MAPPINGS:")
        print("Add these to your word_controls_processor.py:")
        print()
        for control in controls:
            if control.lower() in ['praktijknaam']:
                print(f"    '{control}': data.get('companyName', ''),")
            elif control.lower() in ['naam']:
                print(f"    '{control}': data.get('contactName', ''),")
            elif control.lower() in ['straat']:
                print(f"    '{control}': data.get('address', ''),")
            elif control.lower() in ['postcode']:
                print(f"    '{control}': data.get('postalCode', ''),")
            elif control.lower() in ['stad']:
                print(f"    '{control}': data.get('city', ''),")
            elif control.lower() in ['btw']:
                print(f"    '{control}': data.get('companyId', ''),")
            elif control.lower() in ['items1']:
                print(f"    '{control}': self.format_cost_list(data.get('oneTimeCosts', [])),")
            elif control.lower() in ['items2']:
                print(f"    '{control}': self.format_cost_list(data.get('recurringCosts', [])),")
            else:
                print(f"    '{control}': '',  # TODO: Map this control")