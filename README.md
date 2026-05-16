# Australia Customs Excel to PDF

Convert Australian customs/logistics commercial invoice Excel files (.xlsx) into a single-page landscape A4 PDF that meets customs clearance submission requirements.

## Use Cases

- Convert Excel commercial invoices to PDF for Australian customs submission
- Handle invoice files from Top Logistics, Cainiao, 4PX and other freight forwarders
- Preserve embedded product images in the output PDF
- Generate single-page landscape A4 output for printing or digital submission

## Core Capabilities

- **Single-page landscape A4**: Auto-layout to standard 29.7cm x 21cm customs invoice format
- **Image preservation**: Automatically extract and embed product images from Excel
- **Professional styling**: Red company header, grey table header, full borders, blue WEB links
- **CLI tool**: Support batch conversion via command line

## Tech Stack

- Python 3.7+
- openpyxl (read Excel)
- reportlab (generate PDF)
- CJK font support (Windows: Microsoft YaHei, macOS: PingFang)

## Quick Start

```bash
pip install openpyxl reportlab
python scripts/excel_to_pdf.py input.xlsx output.pdf
```

## Input Format

The tool auto-parses based on the standard Excel layout commonly used by Australian freight forwarders:

| Content | Typical Cell |
|---------|-------------|
| Company name | D1 |
| Company address | D2 |
| Consignee | E4 |
| Date / Invoice No. / Trade Term / Port / Destination | J4-J8 |
| Product detail table | D9:J10 |
| Product images | xl/media/ embedded |

## Related Projects

- [australia-customs-pdf-tla-suite](https://github.com/xuluoforcainiao/australia-customs-pdf-tla-suite) - Bundle this tool with TLA upload workflow
