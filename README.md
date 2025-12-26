# PDF to Excel Converter

A Python application with GUI that extracts PDF filenames and total amounts from invoices and exports them to Excel.

## Features

- 🎨 Modern GUI interface (Tkinter)
- 📁 Browse and select PDF folders
- 📊 Export to Excel with sheet selection
- 💰 Automatic total amount extraction from PDFs
- 🔗 Clickable hyperlinks to open invoices
- ✅ Duplicate detection and prevention
- 🚀 Standalone .exe available (no Python installation required)

## Requirements

- Python 3.13+
- openpyxl
- pdfplumber
- tkinter (included with Python)

## Installation

1. Clone this repository:
```bash
git clone <your-repo-url>
cd Export
```

2. Install dependencies:
```bash
pip install openpyxl pdfplumber
```

## Usage

### Python Script
```bash
python pdf_to_excel.py
```

### Standalone Executable
Simply double-click `PDF_to_Excel_GUI.exe` (no Python installation needed)

## Building the Executable

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --name "PDF_to_Excel_GUI" --clean pdf_to_excel.py
```

The executable will be created in the `dist/` folder.

## How It Works

1. Select a folder containing PDF invoices
2. Choose an Excel file (new or existing)
3. Select a sheet or create a new one
4. Click "Convert PDFs to Excel"
5. The app extracts:
   - PDF filename
   - Total amount (searches for "Total Amount" field)
   - Creates clickable hyperlink to the invoice

## Output Format

The Excel file will contain:
- **Column A**: PDF Filename
- **Column B**: Total Amount
- **Column C**: Clickable link to open the invoice

## License

MIT License

## Author

Created on December 26, 2025
