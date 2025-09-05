#!/usr/bin/env python3
"""
Precise Template Fixer
Uses direct XML manipulation to fix broken template tags in Word documents.
This version handles the complex XML structure more precisely.
"""

import zipfile
import xml.etree.ElementTree as ET
import re
import sys
import os
import tempfile
import shutil

class PreciseTemplateFixer:
    def __init__(self):
        # Known broken patterns and their correct replacements
        self.replacements = {
            # Broken patterns from your template
            '{{praktijknaam': '{companyName}',
            'praktijknaam}}': '',
            '{{naam': '{contactName}', 
            'naam}}': '',
            '{{straat': '{address}',
            'straat}}': '',
            '{{postcode': '{postalCode}',
            'postcode}}': '',
            '{{stad': '{city}',
            'stad}}': '',
            '{{btw': '{companyId}',
            'btw}}': '',
            '{{SigB_es_:signer1:signatureblock': '{date}',
            'SigB_es_:signer1:signatureblock}}': '',
            
            # Handle any remaining fragments
            '{{prak': '{companyName}',
            'tijk': '',
            'naam}': '}',
            '{stra': '{address}',
            'raat}': '}',
            '{post': '{postalCode}',
            'code}': '}',
            '{stad': '{city}',
            'stad}': '}',
            '{btw}': '{companyId}',
            
            # Clean up multiple braces
            '{{{': '{{',
            '}}}': '}}',
            '{{{{': '{{',
            '}}}}': '}}',
        }

    def fix_xml_content_direct(self, xml_content):
        """Directly fix XML content with string replacements."""
        
        original_content = xml_content
        
        # Apply all replacements
        for broken, fixed in self.replacements.items():
            xml_content = xml_content.replace(broken, fixed)
        
        # Additional cleanup patterns
        cleanup_patterns = [
            # Remove orphaned parts
            (r'\{[^}]*praktijk[^}]*\}', ''),
            (r'\{[^}]*naam[^}]*\}(?!})', ''),
            (r'(?<!\{)\{straat[^}]*\}', ''),
            (r'\{[^}]*raat[^}]*\}', ''),
            (r'\{[^}]*post[^}]*\}(?!alCode)', ''),
            (r'\{[^}]*code[^}]*\}(?!\})', ''),
            (r'(?<!\{)\{stad[^}]*\}(?!\})', ''),
            
            # Fix double braces issues
            (r'\{\{([^}]+)\}\{([^}]+)\}\}', r'{{\1\2}}'),
            (r'\{\{([^}]+)\}([^}]+)\}\}', r'{{\1\2}}'),
            (r'\{\{([^}]+)([^}]+)\}\}', r'{{\1\2}}'),
        ]
        
        for pattern, replacement in cleanup_patterns:
            xml_content = re.sub(pattern, replacement, xml_content)
        
        # Final cleanup - ensure no broken tags remain
        xml_content = re.sub(r'\{[^}]*praktijk[^}]*\}', '', xml_content)
        xml_content = re.sub(r'\{[^}]*naam[^}]*\}(?!e\})', '', xml_content)
        xml_content = re.sub(r'\{[^}]*straat[^}]*\}', '', xml_content)
        xml_content = re.sub(r'\{[^}]*postcode[^}]*\}', '', xml_content)
        xml_content = re.sub(r'\{[^}]*stad[^}]*\}', '', xml_content)
        
        changes_made = xml_content != original_content
        
        return xml_content, changes_made

    def process_docx_file(self, input_file, output_file=None):
        """Process the DOCX file to fix broken template tags."""
        
        if output_file is None:
            output_file = input_file
        
        print(f"🔧 Fixing template: {input_file}")
        print("=" * 60)
        
        # Create backup
        backup_file = input_file.replace('.docx', '_backup.docx')
        if not os.path.exists(backup_file):
            shutil.copy2(input_file, backup_file)
            print(f"📋 Backup created: {backup_file}")
        
        files_processed = 0
        
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_file = os.path.join(temp_dir, 'processing.docx')
                shutil.copy2(input_file, temp_file)
                
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
                
                with zipfile.ZipFile(temp_file, 'r') as input_zip:
                    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as output_zip:
                        
                        for item in input_zip.infolist():
                            data = input_zip.read(item.filename)
                            
                            if item.filename in target_files:
                                try:
                                    print(f"📄 Processing: {item.filename}")
                                    
                                    # Decode XML content
                                    xml_content = data.decode('utf-8')
                                    
                                    # Fix the content
                                    fixed_content, changes_made = self.fix_xml_content_direct(xml_content)
                                    
                                    if changes_made:
                                        print(f"   ✅ Fixed broken template tags")
                                        files_processed += 1
                                    else:
                                        print(f"   ℹ️  No changes needed")
                                    
                                    # Write the fixed content
                                    output_zip.writestr(item, fixed_content.encode('utf-8'))
                                    
                                except Exception as e:
                                    print(f"   ⚠️  Error processing {item.filename}: {e}")
                                    # Copy original if processing fails
                                    output_zip.writestr(item, data)
                            else:
                                # Copy other files unchanged
                                output_zip.writestr(item, data)
                
                print(f"\n✅ Template fixing complete!")
                print(f"📊 Files processed: {files_processed}")
                print(f"💾 Output saved to: {output_file}")
                
                return True
                
        except Exception as e:
            print(f"❌ Error fixing template: {e}")
            return False

def main():
    """Main function."""
    
    template_file = "standaardofferte Compufit NL.docx"
    
    if len(sys.argv) > 1:
        template_file = sys.argv[1]
    
    if not os.path.exists(template_file):
        print(f"❌ Error: Template file '{template_file}' not found!")
        return 1
    
    # Initialize fixer
    fixer = PreciseTemplateFixer()
    
    # Fix the template
    success = fixer.process_docx_file(template_file)
    
    if success:
        print(f"\n🎉 Template fixed successfully!")
        print(f"💡 Test your template with: python3 check_template.py")
        print(f"🚀 Ready to use with the quotation generator")
    else:
        print(f"\n❌ Template fixing failed")
    
    return 0 if success else 1

if __name__ == "__main__":
    exit(main())