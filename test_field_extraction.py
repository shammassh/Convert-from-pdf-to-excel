"""
Test script to verify field extraction functionality
Run this to test the new field mapping features without the GUI
"""

import os
from pdf_to_excel import extract_all_fields_from_pdf, extract_field_from_pdf

def test_field_extraction():
    """Test field extraction on a sample PDF"""
    
    print("=" * 60)
    print("PDF to Excel Converter v2.0 - Field Extraction Test")
    print("=" * 60)
    
    # Get PDF folder
    pdf_folder = input("\nEnter path to folder with PDF files: ").strip().strip('"')
    
    if not os.path.exists(pdf_folder):
        print(f"❌ Error: Folder '{pdf_folder}' does not exist")
        return
    
    # Get PDF files
    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"❌ Error: No PDF files found in '{pdf_folder}'")
        return
    
    print(f"\n✓ Found {len(pdf_files)} PDF file(s)")
    
    # Test with first PDF
    sample_pdf = os.path.join(pdf_folder, pdf_files[0])
    print(f"\n📄 Analyzing: {pdf_files[0]}")
    print("-" * 60)
    
    # Extract all fields
    print("\n🔍 Extracting all fields...")
    fields = extract_all_fields_from_pdf(sample_pdf)
    
    if not fields:
        print("❌ No fields detected in PDF")
        print("   This might be a scanned image without text layer")
        return
    
    print(f"\n✓ Detected {len(fields)} field(s):\n")
    
    # Display fields with values
    for idx, (field_name, field_value) in enumerate(fields.items(), 1):
        print(f"{idx:3d}. {field_name:30s} = {field_value}")
    
    # Test targeted extraction
    print("\n" + "=" * 60)
    print("Testing Targeted Field Extraction")
    print("=" * 60)
    
    # Let user select fields to test
    print("\nEnter field names to extract (comma-separated):")
    print("Example: Total Amount, Invoice Number, Invoice Date")
    
    field_input = input("Fields: ").strip()
    
    if field_input:
        selected_fields = [f.strip() for f in field_input.split(',')]
        print(f"\n🔍 Extracting: {', '.join(selected_fields)}")
        print("-" * 60)
        
        results = extract_field_from_pdf(sample_pdf, selected_fields)
        
        print("\n📊 Extraction Results:\n")
        for field in selected_fields:
            value = results.get(field, "N/A")
            status = "✓" if field in results else "⚠"
            print(f"  {status} {field:30s} = {value}")
    
    # Test on all PDFs (first 5)
    print("\n" + "=" * 60)
    print("Testing Field Consistency Across Multiple PDFs")
    print("=" * 60)
    
    test_count = min(5, len(pdf_files))
    print(f"\nAnalyzing first {test_count} PDF(s)...\n")
    
    common_fields = set(fields.keys())
    
    for i, pdf_file in enumerate(pdf_files[:test_count], 1):
        pdf_path = os.path.join(pdf_folder, pdf_file)
        print(f"{i}. {pdf_file[:40]:<40} ", end="")
        
        pdf_fields = extract_all_fields_from_pdf(pdf_path)
        field_count = len(pdf_fields)
        
        common_fields &= set(pdf_fields.keys())
        
        print(f"→ {field_count} fields")
    
    print(f"\n✓ Common fields across all tested PDFs: {len(common_fields)}")
    
    if common_fields:
        print("\n📋 Recommended fields for mapping:")
        for idx, field in enumerate(sorted(common_fields)[:10], 1):
            print(f"   {idx}. {field}")
    
    print("\n" + "=" * 60)
    print("✓ Test Complete!")
    print("=" * 60)
    print("\n💡 Next Steps:")
    print("   1. Run the GUI: python pdf_to_excel.py")
    print("   2. Click 'Configure Fields'")
    print("   3. Select fields from the list above")
    print("   4. Save mapping and convert your PDFs")
    print("\n")

if __name__ == "__main__":
    try:
        test_field_extraction()
    except KeyboardInterrupt:
        print("\n\n⚠ Test cancelled by user")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
