#!/usr/bin/env python3
"""
Unified Quotation Server
Combines web interface serving and quotation processing on a single port (8000).
"""

import http.server
import socketserver
import json
import os
import sys
import tempfile
import shutil
from urllib.parse import parse_qs, urlparse
from datetime import datetime
import io

class UnifiedQuotationHandler(http.server.SimpleHTTPRequestHandler):
    
    def end_headers(self):
        """Add CORS headers for all responses."""
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()
    
    def do_OPTIONS(self):
        """Handle preflight requests."""
        self.send_response(200)
        self.end_headers()
    
    def do_POST(self):
        """Handle POST requests for quotation generation."""
        
        if self.path == '/generate-quotation':
            try:
                # Read the request data
                content_length = int(self.headers['Content-Length'])
                post_data = self.rfile.read(content_length)
                quotation_data = json.loads(post_data.decode('utf-8'))
                
                print(f"🎯 Generating quotation for: {quotation_data.get('companyName', 'Unknown')}")
                
                # Process the quotation
                result = self.create_quotation_document(quotation_data)
                
                if result['success']:
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    
                    response = {
                        'success': True,
                        'message': 'Quotation generated successfully',
                        'filename': result['filename'],
                        'download_url': f'/download/{result["filename"]}'
                    }
                    self.wfile.write(json.dumps(response).encode())
                else:
                    self.send_error(500, result['error'])
            
            except Exception as e:
                print(f"❌ Error generating quotation: {e}")
                self.send_error(500, f"Error generating quotation: {str(e)}")
        else:
            self.send_error(404, "Endpoint not found")
    
    def do_GET(self):
        """Handle GET requests, including file downloads and static files."""
        
        if self.path.startswith('/download/'):
            filename = self.path[10:]  # Remove '/download/' prefix
            if os.path.exists(filename) and filename.endswith('.docx'):
                try:
                    # Send DOCX file directly
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                    self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                    self.end_headers()
                    
                    with open(filename, 'rb') as f:
                        self.wfile.write(f.read())
                    
                    print(f"📥 Downloaded: {filename}")
                    
                    # Clean up file after download
                    try:
                        os.remove(filename)
                        print(f"🧹 Cleaned up temporary file")
                    except Exception as cleanup_error:
                        print(f"⚠️ Cleanup warning: {cleanup_error}")
                    
                except Exception as e:
                    print(f"❌ Download error: {e}")
                    self.send_error(500, f"Download error: {str(e)}")
            else:
                self.send_error(404, "File not found")
        else:
            # Handle static file serving (HTML, CSS, JS, etc.)
            super().do_GET()
    
    def create_quotation_document(self, data):
        """Create a quotation document using XML processing for broken tags."""
        
        template_file = "standaardofferte Compufit NL.docx"
        
        if not os.path.exists(template_file):
            return {'success': False, 'error': f'Template file {template_file} not found'}
        
        try:
            from content_control_processor import ContentControlProcessor
            
            print("📄 Creating quotation document with content control processing...")
            
            # Generate filename
            company_safe = data.get('companyName', 'unknown').replace(' ', '_').replace('/', '_')
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Offerte_{company_safe}_{timestamp}.docx"
            
            # Process the template using content control processor
            processor = ContentControlProcessor()
            success = processor.process_word_template(template_file, data, filename)
            
            if success:
                print(f"✅ Quotation created: {filename}")
                return {'success': True, 'filename': filename}
            else:
                return {'success': False, 'error': 'Failed to process template'}
            
        except ImportError:
            return {'success': False, 'error': 'Required libraries not available'}
        except Exception as e:
            print(f"❌ Error creating quotation: {e}")
            return {'success': False, 'error': str(e)}

def main():
    """Start the unified quotation server."""
    
    # Use Railway's PORT environment variable if available, otherwise default to 80
    port = int(os.environ.get('PORT', 80))
    if len(sys.argv) > 1:
        port = int(sys.argv[1])
    
    print(f"🚀 Starting Unified Quotation Server")
    print(f"📍 Port: {port}")
    print(f"🌐 URL: http://localhost:{port}")
    print(f"📋 Template: standaardofferte Compufit NL.docx")
    print("=" * 60)
    
    # Check if template exists
    if not os.path.exists("standaardofferte Compufit NL.docx"):
        print("❌ Warning: Template file 'standaardofferte Compufit NL.docx' not found!")
        print("   Make sure the template is in the same directory as this script.")
    
    # Check required files
    required_files = [
        "index.html",
        "script.js", 
        "style.css"
    ]
    
    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print("❌ Missing required web files:")
        for file in missing_files:
            print(f"   - {file}")
    
    # Check required dependencies
    required_packages = [
        ('docx', 'python-docx'),
        ('lxml', 'lxml')
    ]
    
    missing_packages = []
    for package_name, pip_name in required_packages:
        try:
            __import__(package_name)
            print(f"✅ {pip_name} library found")
        except ImportError:
            missing_packages.append(pip_name)
            print(f"❌ {pip_name} library not found!")
    
    if missing_packages:
        print(f"\n📦 Install missing packages with:")
        print(f"   pip install {' '.join(missing_packages)}")
        print(f"   or: pip install -r requirements.txt")
        return 1
    
    try:
        with socketserver.TCPServer(("", port), UnifiedQuotationHandler) as httpd:
            print(f"✅ Unified server running! Access your quotation system at:")
            print(f"   http://localhost:{port}")
            print(f"\n💡 This server handles both:")
            print(f"   📱 Web interface (HTML, CSS, JS)")
            print(f"   ⚙️  Quotation API (/generate-quotation)")
            print(f"   📥 File downloads (/download/) - DOCX format")
            print(f"\n🛑 Press Ctrl+C to stop the server")
            
            httpd.serve_forever()
            
    except KeyboardInterrupt:
        print(f"\n👋 Server stopped")
    except Exception as e:
        print(f"❌ Server error: {e}")

if __name__ == "__main__":
    main()
