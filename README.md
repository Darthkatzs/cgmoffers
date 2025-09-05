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
2. **Open** `index.html` in a web browser
3. **Fill in** the quotation details
4. **Generate** your professional quotation

## File Structure

```
cgmoffers/
├── index.html              # Main application page
├── style.css               # Application styles
├── script.js               # Application functionality
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
3. Open `index.html` in a web browser
4. Start creating quotations!

## License

This project is licensed under the MIT License.

## Support

For questions or issues, please create an issue in this repository.