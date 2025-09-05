#!/usr/bin/env python3
"""
Word Template Auto-Fixer
Automatically fixes broken template tags in .docx files by directly editing the XML content.
"""

import zipfile
import xml.etree.ElementTree as ET
import re
import sys
import tempfile
import os
import shutil

def fix_xml_content(xml_content):
    """Fix broken template tags in XML content."""
    
    # Define the replacement mappings for broken tags
    replacements = [
        # Fix broken opening tags
        ('{{praktijknaam', '{companyName'),
        ('{{naam', '{contactName'),
        ('{{straat', '{address'),
        ('{{nummer', '{houseNumber'),
        ('{{postcode', '{postalCode'),
        ('{{stad', '{city'),
        ('{{btw}', '{companyId'),
        ('{{SigB_es_:signer1:signatureblock', '{date'),
        
        # Fix broken closing tags
        ('praktijknaam}}', '}'),
        ('naam}}', '}'),
        ('straat}}', '}'),
        ('nummer}}', '}'),
        ('postcode}}', '}'),
        ('stad}}', '}'),
        ('{btw}}', '}'),
        ('SigB_es_:signer1:signatureblock}}', '}'),
        
        # Fix complete tags to proper variable names
        ('{praktijknaam}', '{companyName}'),
        ('{naam}', '{contactName}'),
        ('{straat}', '{address}'),
        ('{nummer}', '{houseNumber}'),
        ('{postcode}', '{postalCode}'),
        ('{stad}', '{city}'),
        ('{btw}', '{companyId}'),
        ('{SigB_es_:signer1:signatureblock}', '{date}'),
        
        # Clean up any remaining triple braces or malformed tags
        ('{{{', '{{'),
        ('}}}', '}}'),
    ]
    
    # Apply all replacements
    for old, new in replacements:
        xml_content = xml_content.replace(old, new)
    
    # Additional cleanup for any remaining malformed patterns
    # Remove any standalone opening braces that might be left
    xml_content = re.sub(r'\{\{(?![a-zA-Z])', '{', xml_content)
    xml_content = re.sub(r'(?<![a-zA-Z])\}\}', '}', xml_content)
    
    return xml_content

def fix_docx_template(input_filename, output_filename=None):
    """Fix broken template tags in a .docx file."""
    
    if output_filename is None:
        output_filename = input_filename
    
    print(f"🔧 Fixing template: {input_filename}")
    print("=" * 60)
    
    # Create a temporary directory for processing
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_docx = os.path.join(temp_dir, "temp.docx")
        
        # Copy input file to temp location
        shutil.copy2(input_filename, temp_docx)
        
        # Files to process
        files_to_fix = [
            'word/document.xml',
            'word/header1.xml',
            'word/header2.xml', 
            'word/header3.xml',
            'word/footer1.xml',
            'word/footer2.xml',
            'word/footer3.xml'
        ]
        
        files_processed = 0
        
        # Process the docx file
        with zipfile.ZipFile(temp_docx, 'r') as input_zip:
            with zipfile.ZipFile(output_filename, 'w', zipfile.ZIP_DEFLATED) as output_zip:
                
                # Process each file in the zip
                for item in input_zip.infolist():
                    data = input_zip.read(item.filename)
                    
                    # Check if this is a file we need to fix
                    if item.filename in files_to_fix:
                        try:
                            # Decode XML content
                            xml_content = data.decode('utf-8')
                            
                            # Count issues before fixing
                            issues_before = len(re.findall(r'\{\{[^}]+|\}[^}]*\}\}', xml_content))
                            
                            # Fix the content
                            fixed_content = fix_xml_content(xml_content)
                            
                            # Count issues after fixing
                            issues_after = len(re.findall(r'\{\{[^}]+|\}[^}]*\}\}', fixed_content))
                            
                            if issues_before > issues_after:
                                print(f"📄 Fixed {item.filename}: {issues_before} → {issues_after} issues")
                                files_processed += 1
                            
                            # Write the fixed content
                            output_zip.writestr(item, fixed_content.encode('utf-8'))
                            
                        except Exception as e:
                            print(f"⚠️  Error processing {item.filename}: {e}")
                            # If there's an error, copy the original
                            output_zip.writestr(item, data)
                    else:
                        # Copy other files unchanged
                        output_zip.writestr(item, data)
    
    print(f"\n✅ Processing complete!")
    print(f"📊 Files processed: {files_processed}")
    print(f"💾 Output saved to: {output_filename}")
    
    return files_processed > 0

def main():
    input_file = "standaardofferte Compufit NL.docx"
    
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    
    if not os.path.exists(input_file):
        print(f"❌ Error: File '{input_file}' not found!")
        return 1
    
    # Create backup
    backup_file = input_file.replace('.docx', '_backup.docx')
    if not os.path.exists(backup_file):
        shutil.copy2(input_file, backup_file)
        print(f"📋 Backup created: {backup_file}")
    
    # Fix the template
    success = fix_docx_template(input_file)
    
    if success:
        print(f"\n🎉 Template fixed successfully!")
        print(f"💡 Run 'python3 check_template.py' to verify the fixes")
    else:
        print(f"\n⚠️  No issues found or no changes made")
    
    return 0

if __name__ == "__main__":
    exit(main())