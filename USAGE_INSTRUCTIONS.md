# Compufit Quotation System

A web-based quotation system that uses your existing Word template to generate professional quotations.

## ✅ System Status: WORKING

Your quotation system is now fully functional and uses your existing Word contract template!

## 🚀 Quick Start

### Option 1: Start Everything (Recommended)
```bash
python3 start_quotation_system.py
```

This starts both the web interface (port 8000) and quotation processor (port 8001).

### Option 2: Manual Start
```bash
# Terminal 1: Web interface
python3 server.py

# Terminal 2: Quotation processor  
python3 final_quotation_server.py
```

## 🌐 Access Your System

Open your browser and go to: **http://localhost:8000**

## 📋 How It Works

1. **Fill out the form** with company details and costs
2. **Click "Genereer PDF"** - the system will:
   - Use your existing `standaardofferte Compufit NL.docx` template
   - Replace company information in the template
   - Add a professional cost breakdown table
   - Generate a complete Word document quotation

## 📁 Generated Files

Quotations are saved as: `Offerte_CompanyName_TIMESTAMP.docx`

## 🔧 Key Features

- ✅ **Uses your existing Word contract template**
- ✅ **No template modifications needed**
- ✅ **Professional cost breakdown tables**
- ✅ **Automatic calculations**
- ✅ **Clean, modern web interface**
- ✅ **Works with broken XML tags in Word**

## 📊 Testing

Test the system with:
```bash
python3 test_quotation.py
```

## 🛠 Technical Solution

The system bypasses the docxtemplater issues by:

1. **Using python-docx** instead of JavaScript docxtemplater
2. **Direct text replacement** in your existing template
3. **Adding cost tables** as new sections
4. **Server-side processing** to avoid browser limitations

## 📄 Files Overview

- `index.html` - Web interface
- `script.js` - Frontend logic
- `style.css` - Styling
- `final_quotation_server.py` - Main quotation processor
- `server.py` - Web server
- `start_quotation_system.py` - Easy startup script
- `test_quotation.py` - Test script

## 🔍 Troubleshooting

### Port Already in Use
```bash
# Find and kill processes on ports 8000/8001
lsof -ti:8000 | xargs kill -9
lsof -ti:8001 | xargs kill -9
```

### Missing Dependencies
```bash
pip3 install python-docx requests
```

### Template Not Found
Make sure `standaardofferte Compufit NL.docx` is in the same folder as the scripts.

## 🎉 Success!

Your quotation system is now ready for your team to use. The system:

- ✅ Uses your existing large contract template
- ✅ Handles all the broken XML tag issues automatically
- ✅ Generates professional quotations with cost breakdowns
- ✅ Is ready for production use

No more template fixes needed - just run and use!