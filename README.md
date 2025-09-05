# CGM Offers - Quotation Generator

A web-based application for generating professional quotations using Word templates.

## Features

- 📄 **Template-based PDF Generation** - Uses existing Word templates
- 💰 **Dynamic Cost Calculation** - Real-time totals for one-time and recurring costs
- 📱 **Responsive Design** - Works on desktop, tablet, and mobile
- 🇳🇱 **Dutch Interface** - Fully localized for Dutch business use
- ⚡ **Easy to Use** - Simple form-based interface

## Quick Start

1. **Prepare your Word template** following the instructions in `TEMPLATE_SETUP.md`
2. **Start the local server** (required for template loading):
   - **Python**: `python3 server.py` or `python server.py`
   - **Node.js**: `node serve.js`
   - **Alternative**: Any HTTP server in the project directory
3. **Open** your browser to `http://localhost:8000`
4. **Fill in** the quotation details
5. **Generate** your professional quotation

### Why use a server?

The application needs to load the Word template file, which requires a local server due to browser security restrictions. You can't open `index.html` directly in your browser - it must be served through HTTP.

## File Structure

```
cgmoffers/
├── index.html              # Main application page
├── style.css               # Application styles
├── script.js               # Application functionality
├── server.py               # Python HTTP server
├── serve.js                # Node.js HTTP server
├── TEMPLATE_SETUP.md       # Word template preparation guide
├── standaardofferte Compufit NL.docx  # Word template
└── README.md               # This file
```

## How It Works

1. **Data Input** - Enter customer details and cost items through the web form
2. **Template Processing** - JavaScript processes your Word template with docxtemplater
3. **Document Generation** - Creates a filled Word document with all your data
4. **Download** - Automatically downloads the completed quotation

## Template Variables

The application uses these template variables in your Word document:

### Customer Information
- `{companyName}` - Customer company name
- `{contactName}` - Contact person
- `{fullAddress}` - Complete address
- `{companyId}` - Company registration number
- `{date}` - Current date

### Cost Items
- `{#oneTimeCosts}...{/oneTimeCosts}` - One-time cost items
- `{#recurringCosts}...{/recurringCosts}` - Recurring cost items
- `{oneTimeTotal}` - Total one-time costs
- `{recurringTotal}` - Total recurring costs

## Browser Support

- ✅ Chrome 60+
- ✅ Firefox 55+
- ✅ Safari 12+
- ✅ Edge 79+

## Dependencies

The application uses these external libraries via CDN:
- [PizZip](https://github.com/Stuk/jszip) - ZIP file handling
- [docxtemplater](https://docxtemplater.com/) - Word document templating
- [FileSaver.js](https://github.com/eligrey/FileSaver.js) - File download functionality

## Setup Instructions

1. Clone this repository
2. Follow the template setup instructions in `TEMPLATE_SETUP.md`
3. Start a local server:
   ```bash
   # Using Python (most systems have this)
   python3 server.py
   
   # Or using Node.js (if you have it installed)
   node serve.js
   
   # Or any other HTTP server in the project directory
   ```
4. Open `http://localhost:8000` in your browser
5. Start creating quotations!

## License

This project is licensed under the MIT License.

## Support

For questions or issues, please create an issue in this repository.