# Word Template Setup Instructions

To make your Word template work with the quotation generator, you need to add template placeholders in your document. These placeholders will be replaced with the data from the web form.

## Required Template Variables

Open your `standaardofferte Compufit NL.docx` file and replace the following fields with these template variables:

### Company Information (Page 1 and Page 3)
Replace the existing company information fields with:
- `{companyName}` - Company name
- `{contactName}` - Contact person name
- `{fullAddress}` - Complete address (street, postal code, city)
- `{companyId}` - Company ID/KvK number
- `{date}` - Current date (automatically filled)

### One-time Costs Section
Create a table or section for one-time costs and use:
```
{#hasOneTimeCosts}
Eenmalige Bijdrage:
{#oneTimeCosts}
- {material} | Aantal: {quantity} | Prijs: €{unitPrice} | Totaal: €{total}
{/oneTimeCosts}
Totaal Eenmalige Bijdrage: €{oneTimeTotal}
{/hasOneTimeCosts}
```

### Recurring Costs Section
Create a table or section for recurring costs and use:
```
{#hasRecurringCosts}
Jaarlijkse Bijdrage:
{#recurringCosts}
- {material} | Aantal: {quantity} | Prijs: €{unitPrice} | Totaal: €{total}
{/recurringCosts}
Totaal Jaarlijkse Bijdrage: €{recurringTotal}
{/hasRecurringCosts}
```

## Example Template Structure

Here's how your Word document should look:

### Page 1 - Header
```
OFFERTE

Aan: {companyName}
T.a.v.: {contactName}
Adres: {fullAddress}
KvK: {companyId}

Datum: {date}
```

### Page 2-3 - Costs
```
KOSTENSPECIFICATIE

{#hasOneTimeCosts}
EENMALIGE KOSTEN:
{#oneTimeCosts}
{material}
Aantal: {quantity} × €{unitPrice} = €{total}
{/oneTimeCosts}

Subtotaal Eenmalige Kosten: €{oneTimeTotal}
{/hasOneTimeCosts}

{#hasRecurringCosts}
JAARLIJKSE KOSTEN:
{#recurringCosts}
{material}
Aantal: {quantity} × €{unitPrice} = €{total}
{/recurringCosts}

Subtotaal Jaarlijkse Kosten: €{recurringTotal}
{/hasRecurringCosts}
```

## Important Notes

1. **Exact Syntax**: Use the exact placeholder syntax `{variableName}` with curly braces
2. **Conditional Sections**: Use `{#sectionName}...{/sectionName}` for conditional content
3. **Loops**: Use `{#arrayName}...{/arrayName}` for repeating content like cost items
4. **Save Format**: Save your template as `.docx` format
5. **Backup**: Keep a backup of your original template before making changes

## Testing Your Template

After modifying your template:
1. Open the web application in a browser
2. Fill in sample data
3. Generate a test quotation
4. Check if all fields are filled correctly
5. Adjust placeholders if needed

## Troubleshooting

- **Missing Data**: Check that placeholder names match exactly
- **Formatting Issues**: Ensure placeholders are not split across multiple paragraphs
- **Loop Problems**: Verify that array sections use `{#arrayName}` and `{/arrayName}` syntax
- **Conditional Sections**: Make sure conditional blocks are properly closed