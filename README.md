# Australia Customs Excel to PDF

Convert Australian customs / logistics commercial invoice Excel files (.xlsx) into a single-page landscape A4 PDF.

## When to Use

- Converting Excel commercial invoices to PDF format
- Working with Australian customs or freight forwarding documents from Top Logistics, Cainiao, 4PX, or similar freight forwarders
- Preserving embedded product images in the output PDF
- Needing a single-page landscape A4 output for printing or submission

## Quick Start

```bash
python -m pip install openpyxl reportlab
python scripts/excel_to_pdf.py <input.xlsx> <output.pdf>
```

## Requirements

- Python 3.7+
- `openpyxl`
- `reportlab`
- CJK font support (Microsoft YaHei on Windows, PingFang on macOS)

## License

MIT
