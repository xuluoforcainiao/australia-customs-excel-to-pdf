---
name: australia-customs-excel-to-pdf
description: Convert Australian customs/ logistics commercial invoice Excel files (.xlsx) into a single-page landscape A4 PDF. Use when the user needs to convert Excel invoices to PDF format, especially for Australian customs clearance documents from Top Logistics, Cainiao, 4PX, or similar freight forwarders. Handles embedded product images and preserves the standard commercial invoice layout.
---

# Australia Customs Excel to PDF

## Purpose

Convert Excel commercial invoice files (typically from Australian logistics providers) into a single-page, landscape A4 PDF with a professional layout matching standard customs documentation formats.

## When to Use

- User asks to convert an Excel invoice to PDF
- Working with Australian customs or freight forwarding documents
- Files contain product images that need to be preserved in the PDF
- Need a single-page landscape A4 output for printing or submission

## Requirements

- Python 3.7+ with `openpyxl` and `reportlab` installed
- CJK font support (Microsoft YaHei on Windows, PingFang on macOS)

## Workflow

### Step 1: Verify Dependencies

Ensure required packages are available:

```bash
python -m pip install openpyxl reportlab
```

### Step 2: Run the Conversion Script

```bash
python scripts/excel_to_pdf.py <input.xlsx> <output.pdf>
```

### Step 3: Manual Fallback (if script fails)

If the script cannot be executed, generate the PDF manually following this layout:

1. **Page setup**: Landscape A4 (29.7cm x 21cm)
2. **Outer border**: 1.5pt black rectangle with 2cm margins
3. **Header** (red, centered):
   - Company name from cell D1
   - Address from cell D2
   - Horizontal line separator
4. **Title**: "COMMERCIAL INVOICE" centered, 13pt
5. **Consignee** (left): From cell E4, label "Consignee :" in black, name in red
6. **Right info block** (aligned with underlines):
   - DATE: from J4
   - INVOICE NO.: from J5
   - TRADE TERM: from J6
   - PORT OF LOADING: from J7
   - DESTINATION: from J8
7. **Table** (grey header, full borders, 7 columns):
   - BUYER P.O. NO. | Package | Quantity | DESCRIPTION OF GOODS | Products' Material | UNIT PRICE (AUD) | TOTAL AMOUNT (AUD)
   - Data from row 10
   - GRAND TOTAL from row 11
8. **WEB**: URL from E12 in blue
9. **Product image**: Extract from Excel zip (`xl/media/`) and place below WEB line

### Step 4: Image Extraction

To extract embedded images from the Excel file manually:

```python
import zipfile
z = zipfile.ZipFile('input.xlsx')
for name in z.namelist():
    if name.startswith('xl/media/') and not name.endswith('/'):
        data = z.read(name)
        with open(name.split('/')[-1], 'wb') as f:
            f.write(data)
```

## Expected Excel Structure

The skill assumes a standard layout commonly used by Australian freight forwarders:

| Content | Typical Cell |
|---------|-------------|
| Company name | D1 |
| Company address | D2 |
| Consignee name | E4 |
| Date | J4 |
| Invoice No. | J5 |
| Trade Term | J6 |
| Port of Loading | J7 |
| Destination | J8 |
| Table header | D9:J9 |
| Product data | D10:J10 |
| Grand total | D11:J11 |
| WEB URL | E12 |
| Product image | Embedded in xl/media/ |

## Output Format

- Single page, landscape A4
- Outer border frame
- Red company header text
- Grey table header background
- Complete table borders (horizontal and vertical lines)
- Blue WEB link text
- Preserved product image
