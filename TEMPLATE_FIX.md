# Template Fix Instructions

## The Problem
Your Word template has "duplicate tag" errors because Word splits template variables like `{companyName}` across multiple formatting elements. This happens when you edit the template or apply formatting.

## How to Fix Your Template

### Method 1: Quick Fix (Recommended)
1. **Open your Word template** (`standaardofferte Compufit NL.docx`)
2. **Find and Replace ALL instances** of template variables with clean ones:
   - Find: `{prak` → Replace: `{companyName}`
   - Find: `naam}` → Delete (part of the broken tag)
   - Find: `{stra` → Replace: `{address}`  
   - Find: `raat}` → Delete (part of the broken tag)
   - Find: `{numm` → Replace: `{postalCode}`
   - Find: `mmer}` → Delete (part of the broken tag)
   - Find: `{post` → Replace: `{city}`
   - Find: `code}` → Delete (part of the broken tag) 
   - Find: `{stad` → Replace: `{contactName}`
   - Find: `stad}` → Delete (part of the broken tag)
   - Find: `{btw}` → Replace: `{companyId}`
   - Find: `{SigB` → Replace: `{date}`
   - Find: `lock}` → Delete (part of the broken tag)

3. **Clean up any remaining broken tags** by searching for `{{` and `}}`

### Method 2: Create New Template
1. **Create a new Word document**
2. **Type your template content** but don't add variables yet
3. **Add variables in one go**:
   - Type each variable completely: `{companyName}`, `{contactName}`, etc.
   - **Important**: Don't edit the variables after typing them

## Required Variables for the Web App

Replace your broken variables with these exact ones:

### Customer Information
- `{companyName}` - Company name
- `{contactName}` - Contact person
- `{fullAddress}` - Complete address (street + postal code + city)
- `{companyId}` - Company ID
- `{date}` - Current date

### One-time Costs Section
```
{#hasOneTimeCosts}
EENMALIGE KOSTEN:
{#oneTimeCosts}
{material} - Aantal: {quantity} x €{unitPrice} = €{total}
{/oneTimeCosts}

Totaal Eenmalig: €{oneTimeTotal}
{/hasOneTimeCosts}
```

### Recurring Costs Section  
```
{#hasRecurringCosts}
JAARLIJKSE KOSTEN:
{#recurringCosts}
{material} - Aantal: {quantity} x €{unitPrice} = €{total}
{/recurringCosts}

Totaal Jaarlijks: €{recurringTotal}
{/hasRecurringCosts}
```

## Prevention Tips

1. **Type variables completely in one go** - don't edit them character by character
2. **Use plain text formatting** when typing variables
3. **Don't copy/paste variables** - type them fresh each time
4. **Save frequently** as .docx format
5. **Test the template** after any changes

## Testing Your Fixed Template

1. Save your template as `standaardofferte Compufit NL.docx`
2. Start the server: `python3 server.py`
3. Open `http://localhost:8000`
4. Fill in the form and try generating a quotation
5. If you get errors, repeat the cleaning process

The web application will now work correctly with your fixed template!