# PDF to Excel Converter

A Python application with GUI that extracts custom fields from PDF files and exports them to Excel with configurable field mapping.

## ✨ Features

- 🎨 Modern GUI interface (Tkinter)
- 📁 Browse and select PDF folders
- 📊 Export to Excel with sheet selection
- 🔧 **Custom Field Mapping** - Choose which fields to extract from PDFs
- 🔍 **Automatic Field Detection** - Detects fields from tables and text
- 💾 **Mapping Persistence** - Saves your field configuration for reuse
- 🔗 Clickable hyperlinks to open invoices
- ✅ Duplicate detection and prevention
- 🚀 Standalone .exe available (no Python installation required)

## 🆕 What's New - Major Upgrade

### Field Mapping Configuration
Instead of only extracting "Total Amount", you can now:
1. **Analyze** any PDF to detect available fields
2. **Select** which fields you want to extract
3. **Map** them to Excel columns
4. **Save** your configuration for future use

This means you can extract ANY data from your PDFs - invoice numbers, dates, customer names, line items, etc.

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

### Step-by-Step Guide

1. **Select PDF Folder**: Browse to the folder containing your PDF files
2. **Configure Field Mapping** (Optional but Recommended):
   - Click "Configure Fields" button
   - The app will analyze a sample PDF and show all detected fields
   - Select the fields you want to extract (e.g., Invoice Number, Total Amount, Date, Customer)
   - Arrange them in the order you want them in Excel
   - Click "Save Mapping" - your configuration is saved for future use
3. **Select Excel File**: Choose an existing file or create a new one
4. **Select Sheet**: Choose a sheet or create a new one
5. **Click "Convert PDFs to Excel"**: The app will process all PDFs and extract your configured fields

### Field Detection

The app automatically detects fields from:
- **Tables**: Column headers and their values
- **Text Patterns**: Key-value pairs like "Invoice Number: 12345"
- **Common Invoice Fields**: Total Amount, Invoice Date, Customer, etc.

### Default Behavior

If you don't configure field mapping, the app falls back to the classic mode:
- Extracts "Total Amount" only
- Works like the previous version for backward compatibility

## Output Format

### With Field Mapping
The Excel file will contain:
- **Column A**: PDF Filename
- **Columns B-N**: Your configured fields (as many as you selected)
- **Last Column**: Clickable link to open the invoice

### Example:
| PDF Filename | Invoice Number | Invoice Date | Total Amount | Customer | Path to Invoice |
|--------------|----------------|--------------|--------------|----------|-----------------|
| invoice1.pdf | INV-001       | 2024-01-15   | 239.40      | Acme Corp | Open Invoice |

### Without Field Mapping (Default)
- **Column A**: PDF Filename
- **Column B**: Total Amount
- **Column C**: Clickable link to open the invoice

## Advanced Features

### Field Mapping Persistence
- Your field mapping is automatically saved to `field_mapping.json`
- The configuration persists across app restarts
- You can reconfigure anytime by clicking "Configure Fields"

### Search and Filter
- Use the search box in the field mapping dialog to quickly find fields
- Filter through hundreds of detected fields easily

### Duplicate Prevention
- The app checks for existing entries and skips duplicates
- Only new PDF files are added to Excel

### Sheet Management
- Create multiple sheets for different data sets
- Clear sheet data while keeping headers
- Works with existing Excel files

## Tips for Best Results

1. **Use Consistent PDF Formats**: The field detection works best when PDFs have similar structures
2. **Start with One Sample**: The app uses the first PDF as a sample - make sure it's representative
3. **Check Field Values**: Use the preview in the mapping dialog to verify field extraction
4. **Name Fields Clearly**: Use descriptive field names that match your Excel needs

## Troubleshooting

### No Fields Detected
- Ensure your PDFs have text (not scanned images)
- Try a different sample PDF
- PDFs with complex layouts may need OCR preprocessing

### Wrong Values Extracted
- Reconfigure field mapping with a better sample PDF
- Check if field names are unique in the PDF
- Some PDFs may have inconsistent formatting

### File Permission Errors
- Close the Excel file before running conversion
- Ensure you have write permissions to the output location

## License

MIT License

## Author

Created on December 26, 2025
Major Field Mapping Upgrade - December 27, 2025

