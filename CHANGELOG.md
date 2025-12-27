# Changelog

All notable changes to the PDF to Excel Converter project.

## [2.0.0] - 2025-12-27

### 🎉 Major Feature Release

#### Added
- **Field Mapping System**: Configure which fields to extract from PDFs
- **Field Detection Engine**: Automatically detects available fields in PDFs
- **Interactive Field Selector**: Visual GUI for choosing and ordering fields
- **Field Mapping Dialog**: New window with:
  - Available fields browser with search
  - Selected fields list with reordering
  - Sample value preview pane
  - Add/Remove/Move controls
- **Mapping Persistence**: Save and load field configurations
- **field_mapping.json**: Configuration storage file
- **Multi-field Extraction**: Extract unlimited fields per PDF
- **Custom Excel Columns**: Dynamic column generation based on mapping

#### Enhanced
- **PDF Analysis**: Three extraction methods (tables, text patterns, key-value pairs)
- **Field Preview**: See actual values before committing to field selection
- **Excel Output**: Dynamic columns based on field mapping
- **User Interface**: Added "Configure Fields" button and status indicator
- **Progress Logging**: Detailed extraction information for each field

#### Technical
- New function: `extract_all_fields_from_pdf()` - Comprehensive field extraction
- New function: `extract_field_from_pdf()` - Targeted field extraction
- New function: `write_to_excel_with_mapping()` - Multi-column Excel writer
- New class: `FieldMappingDialog` - Field mapping interface
- Added JSON module for configuration persistence
- Added OrderedDict for field ordering
- Enhanced pattern matching with regex

#### Documentation
- Updated README.md with field mapping guide
- Added UPGRADE_NOTES.md with technical details
- Added QUICK_START.md with step-by-step tutorial
- Added usage examples and troubleshooting

#### Backward Compatibility
- Falls back to "Total Amount" mode if no mapping configured
- Existing Excel files fully compatible
- No breaking changes for current users

---

## [1.0.0] - 2025-12-26

### Initial Release

#### Features
- GUI application using Tkinter
- PDF folder browsing
- Excel file selection and sheet management
- Total Amount extraction from PDFs
- Excel export with:
  - PDF Filename column
  - Total Amount column
  - Hyperlink to invoice
- Duplicate detection and prevention
- Clear sheet functionality
- Progress tracking
- Error handling
- PyInstaller executable support

#### Core Functions
- `get_pdf_files()`: PDF file discovery
- `extract_total_amount()`: Amount extraction with multiple patterns
- `write_to_excel()`: Excel writing with hyperlinks
- `write_to_excel_gui()`: GUI version with logging

#### GUI Features
- Modern interface with color scheme
- Folder and file browsers
- Sheet dropdown with refresh
- Progress text area
- Progress bar for long operations
- Clear sheet data button
- Multi-threaded processing

---

## Version History Summary

| Version | Date       | Key Feature                    |
|---------|------------|--------------------------------|
| 2.0.0   | 2025-12-27 | Field Mapping System           |
| 1.0.0   | 2025-12-26 | Initial Release                |

---

## Upgrade Path

### From v1.0 to v2.0
1. Replace `pdf_to_excel.py` with new version
2. No configuration needed - works as before by default
3. Optional: Click "Configure Fields" to use new features
4. `field_mapping.json` created automatically when you configure fields

---

## Breaking Changes

### v2.0.0
- None! Fully backward compatible

---

## Roadmap

### Future Enhancements (v2.1+)
- [ ] Multiple mapping profiles
- [ ] PDF template recognition
- [ ] Batch processing with different mappings
- [ ] Export to CSV/JSON
- [ ] Field validation rules
- [ ] Custom field formulas
- [ ] OCR integration for scanned PDFs
- [ ] Cloud storage integration
- [ ] Scheduled automation
- [ ] Email notifications

### Under Consideration
- Web interface version
- API for programmatic access
- Machine learning for smart field detection
- Multi-language support
- Database export options

---

## Support & Contributions

- Report issues on GitHub
- Suggest features via GitHub Issues
- Contribute via Pull Requests
- Star the repo if you find it useful!

---

**Current Version**: 2.0.0  
**Last Updated**: December 27, 2025  
**License**: MIT
