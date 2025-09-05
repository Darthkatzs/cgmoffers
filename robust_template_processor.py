#!/usr/bin/env python3
"""
Robust Word Template Processor
Handles broken XML tags in Word documents by reconstructing them from fragments.
This allows using existing Word templates without recreation.
"""

import zipfile
import xml.etree.ElementTree as ET
import re
import sys
import os
import tempfile
import shutil
from collections import defaultdict

class RobustTemplateProcessor:
    def __init__(self):
        self.broken_tag_patterns = [
            # Common broken patterns found in Word documents
            r'(\{\{?)([^{}]*?)(\}?\})',  # Match any brace combinations
            r'\{([^{}]*)\{',             # Opening brace with content
            r'\}([^{}]*)\}',             # Closing brace with content
        ]
        
        # Expected template variables
        self.expected_variables = [
            'companyName', 'contactName', 'address', 'postalCode', 
            'city', 'companyId', 'date', 'oneTimeCosts', 'recurringCosts',
            'oneTimeTotal', 'recurringTotal', 'hasOneTimeCosts', 'hasRecurringCosts'
        ]
    
    def extract_all_text_content(self, xml_content):
        """Extract all text content from XML, including broken tags."""
        try:
            # Parse XML
            root = ET.fromstring(xml_content)
            
            # Extract all text runs
            text_runs = []
            
            # Find all text elements
            for elem in root.iter():
                if elem.tag.endswith('}t'):  # Word text elements
                    if elem.text:
                        text_runs.append(elem.text)
                elif elem.text:
                    text_runs.append(elem.text)
                if elem.tail:
                    text_runs.append(elem.tail)
            
            return ''.join(text_runs)
            
        except ET.ParseError as e:
            print(f"XML Parse Error: {e}")
            return ""
    
    def find_broken_tags(self, text):
        """Find and reconstruct broken template tags."""
        print(f"🔍 Analyzing text content...")
        
        # Find all potential tag fragments
        fragments = []
        
        # Look for various brace patterns
        brace_patterns = [
            r'\{+[^{}]*',     # Opening patterns
            r'[^{}]*\}+',     # Closing patterns
            r'\{[^{}]*\}',    # Complete tags
        ]
        
        for pattern in brace_patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                fragments.append({
                    'text': match.group(),
                    'start': match.start(),
                    'end': match.end()
                })
        
        # Sort fragments by position
        fragments.sort(key=lambda x: x['start'])
        
        print(f"📝 Found {len(fragments)} tag fragments")
        
        return fragments
    
    def reconstruct_tags(self, fragments, text):
        """Attempt to reconstruct proper template tags from fragments."""
        reconstructed = []
        
        for fragment in fragments:
            frag_text = fragment['text']
            print(f"   Fragment: {repr(frag_text)}")
            
            # Try to match known variable names
            for var_name in self.expected_variables:
                # Check if this fragment contains part of a variable name
                if var_name.lower() in frag_text.lower():
                    reconstructed_tag = f"{{{var_name}}}"
                    reconstructed.append({
                        'original': frag_text,
                        'reconstructed': reconstructed_tag,
                        'variable': var_name,
                        'position': fragment['start']
                    })
                    print(f"      → Reconstructed as: {reconstructed_tag}")
                    break
                
                # Check partial matches
                for i in range(3, len(var_name)):
                    if var_name[:i].lower() in frag_text.lower() or var_name[-i:].lower() in frag_text.lower():
                        reconstructed_tag = f"{{{var_name}}}"
                        reconstructed.append({
                            'original': frag_text,
                            'reconstructed': reconstructed_tag,
                            'variable': var_name,
                            'position': fragment['start']
                        })
                        print(f"      → Partial match reconstructed as: {reconstructed_tag}")
                        break
        
        return reconstructed
    
    def fix_xml_content(self, xml_content):
        """Fix broken template tags in XML content."""
        
        # Extract text content
        text_content = self.extract_all_text_content(xml_content)
        
        if not text_content.strip():
            return xml_content
        
        # Find broken tags
        fragments = self.find_broken_tags(text_content)
        
        if not fragments:
            print("   No tag fragments found")
            return xml_content
        
        # Reconstruct proper tags
        reconstructed = self.reconstruct_tags(fragments, text_content)
        
        if not reconstructed:
            print("   No tags could be reconstructed")
            return xml_content
        
        # Apply fixes to XML content
        fixed_xml = xml_content
        
        # Sort by position (reverse order to avoid position shifts)
        reconstructed.sort(key=lambda x: x['position'], reverse=True)
        
        for fix in reconstructed:
            # Replace the broken fragment with the reconstructed tag
            fixed_xml = fixed_xml.replace(fix['original'], fix['reconstructed'])
        
        # Additional cleanup
        fixed_xml = self.cleanup_remaining_issues(fixed_xml)
        
        return fixed_xml
    
    def cleanup_remaining_issues(self, xml_content):
        """Clean up any remaining tag issues."""
        
        # Remove multiple consecutive braces
        xml_content = re.sub(r'\{\{+', '{{', xml_content)
        xml_content = re.sub(r'\}+\}', '}}', xml_content)
        
        # Fix common malformed patterns
        replacements = [
            (r'\{\s*\{', '{{'),
            (r'\}\s*\}', '}}'),
            (r'\{\{([^}]*)\}\{([^}]*)\}\}', r'{{\1\2}}'),  # Merge split variables
        ]
        
        for pattern, replacement in replacements:
            xml_content = re.sub(pattern, replacement, xml_content)
        
        return xml_content
    
    def process_docx_template(self, input_file, output_file=None):
        """Process the Word template to fix broken tags."""
        
        if output_file is None:
            output_file = input_file
        
        print(f"🔧 Processing template: {input_file}")
        print("=" * 60)
        
        # Create backup
        backup_file = input_file.replace('.docx', '_backup.docx')
        if not os.path.exists(backup_file):
            shutil.copy2(input_file, backup_file)
            print(f"📋 Backup created: {backup_file}")
        
        files_processed = 0
        
        # Process the docx file
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_docx = os.path.join(temp_dir, "processing.docx")
            shutil.copy2(input_file, temp_docx)
            
            # Files to process
            target_files = [
                'word/document.xml',
                'word/header1.xml',
                'word/header2.xml',
                'word/header3.xml',
                'word/footer1.xml',
                'word/footer2.xml',
                'word/footer3.xml'
            ]
            
            try:
                with zipfile.ZipFile(temp_docx, 'r') as input_zip:
                    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as output_zip:
                        
                        for item in input_zip.infolist():
                            data = input_zip.read(item.filename)
                            
                            if item.filename in target_files:
                                try:
                                    print(f"\n📄 Processing: {item.filename}")
                                    
                                    # Decode and fix XML content
                                    xml_content = data.decode('utf-8')
                                    fixed_content = self.fix_xml_content(xml_content)
                                    
                                    # Check if changes were made
                                    if fixed_content != xml_content:
                                        print(f"   ✅ Fixed template tags")
                                        files_processed += 1
                                    else:
                                        print(f"   ℹ️  No changes needed")
                                    
                                    # Write fixed content
                                    output_zip.writestr(item, fixed_content.encode('utf-8'))
                                    
                                except Exception as e:
                                    print(f"   ⚠️  Error processing {item.filename}: {e}")
                                    # Copy original if processing fails
                                    output_zip.writestr(item, data)
                            else:
                                # Copy other files unchanged
                                output_zip.writestr(item, data)
                
                print(f"\n✅ Template processing complete!")
                print(f"📊 Files processed: {files_processed}")
                print(f"💾 Output saved to: {output_file}")
                
                return True
                
            except Exception as e:
                print(f"❌ Error processing template: {e}")
                return False

def main():
    """Main function."""
    
    template_file = "standaardofferte Compufit NL.docx"
    
    if len(sys.argv) > 1:
        template_file = sys.argv[1]
    
    if not os.path.exists(template_file):
        print(f"❌ Error: Template file '{template_file}' not found!")
        return 1
    
    # Initialize processor
    processor = RobustTemplateProcessor()
    
    # Process the template
    success = processor.process_docx_template(template_file)
    
    if success:
        print(f"\n🎉 Template processing completed!")
        print(f"💡 Test your template with: python3 check_template.py")
        print(f"🚀 Ready to use with the quotation generator")
    else:
        print(f"\n❌ Template processing failed")
    
    return 0 if success else 1

if __name__ == "__main__":
    exit(main())