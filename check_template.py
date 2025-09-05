#!/usr/bin/env python3
"""
Word Template Tag Checker
Analyzes .docx files to find template variables and detect potential issues.
"""

import zipfile
import xml.etree.ElementTree as ET
import re
import sys
from collections import defaultdict

def extract_text_from_xml(xml_content):
    """Extract all text content from Word XML, preserving tag fragments."""
    try:
        root = ET.fromstring(xml_content)
        
        # Find all text elements
        text_elements = []
        for elem in root.iter():
            if elem.text:
                text_elements.append(elem.text)
            if elem.tail:
                text_elements.append(elem.tail)
        
        return ''.join(text_elements)
    except ET.ParseError:
        return ""

def find_template_tags(text):
    """Find all potential template tags in text."""
    # Find complete tags like {variable}
    complete_tags = re.findall(r'\{[^{}]*\}', text)
    
    # Find potential broken tags
    open_braces = re.findall(r'\{+[^{}]*(?=\{|\}|$)', text)
    close_braces = re.findall(r'[^{}]*\}+', text)
    
    return complete_tags, open_braces, close_braces

def analyze_docx_template(filename):
    """Analyze a .docx file for template tags and issues."""
    print(f"🔍 Analyzing template: {filename}")
    print("=" * 60)
    
    try:
        with zipfile.ZipFile(filename, 'r') as docx:
            # Files to check for template tags
            files_to_check = [
                'word/document.xml',
                'word/header1.xml',
                'word/header2.xml', 
                'word/header3.xml',
                'word/footer1.xml',
                'word/footer2.xml',
                'word/footer3.xml'
            ]
            
            all_complete_tags = set()
            all_issues = []
            
            for xml_file in files_to_check:
                try:
                    xml_content = docx.read(xml_file).decode('utf-8')
                    text = extract_text_from_xml(xml_content)
                    
                    complete_tags, open_braces, close_braces = find_template_tags(text)
                    
                    if complete_tags or open_braces or close_braces:
                        print(f"\n📄 File: {xml_file}")
                        
                        if complete_tags:
                            print(f"✅ Complete tags found: {len(complete_tags)}")
                            for tag in complete_tags:
                                print(f"   {tag}")
                                all_complete_tags.add(tag)
                        
                        if open_braces:
                            print(f"⚠️  Potential broken opening tags: {len(open_braces)}")
                            for tag in open_braces[:10]:  # Limit output
                                print(f"   {repr(tag)}")
                                all_issues.append(f"{xml_file}: {repr(tag)}")
                        
                        if close_braces:
                            print(f"⚠️  Potential broken closing tags: {len(close_braces)}")
                            for tag in close_braces[:10]:  # Limit output
                                print(f"   {repr(tag)}")
                                all_issues.append(f"{xml_file}: {repr(tag)}")
                
                except KeyError:
                    # File doesn't exist, skip it
                    continue
                except Exception as e:
                    print(f"❌ Error reading {xml_file}: {e}")
    
    except Exception as e:
        print(f"❌ Error opening {filename}: {e}")
        return
    
    # Summary
    print("\n" + "=" * 60)
    print("📊 SUMMARY")
    print("=" * 60)
    
    if all_complete_tags:
        print(f"✅ Valid template tags found: {len(all_complete_tags)}")
        for tag in sorted(all_complete_tags):
            print(f"   {tag}")
    else:
        print("❌ No valid template tags found")
    
    if all_issues:
        print(f"\n⚠️  Issues found: {len(all_issues)}")
        print("These may cause 'duplicate tag' errors:")
        for issue in all_issues[:20]:  # Limit output
            print(f"   {issue}")
        
        if len(all_issues) > 20:
            print(f"   ... and {len(all_issues) - 20} more issues")
    else:
        print("\n✅ No tag issues detected")
    
    # Recommendations
    print("\n💡 RECOMMENDATIONS")
    print("=" * 60)
    
    if all_issues:
        print("❌ Template has issues that need fixing:")
        print("1. Use Find & Replace in Word to clean up broken tags")
        print("2. Follow the instructions in TEMPLATE_FIX.md")
        print("3. Re-run this script to verify fixes")
    else:
        print("✅ Template looks clean!")
        print("🚀 Ready for use with the quotation generator")
    
    return len(all_issues) == 0

if __name__ == "__main__":
    template_file = "standaardofferte Compufit NL.docx"
    
    if len(sys.argv) > 1:
        template_file = sys.argv[1]
    
    analyze_docx_template(template_file)