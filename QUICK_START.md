# Quick Start Guide - Field Mapping

## 🚀 How to Use the New Field Mapping Feature

### Step 1: Launch the Application
```
python pdf_to_excel.py
```

### Step 2: Select Your PDF Folder
Click **"Browse Folder"** and choose the folder containing your PDF files.

### Step 3: Configure Field Mapping (New!)
Click the **"Configure Fields"** button.

#### What happens:
1. The app analyzes the first PDF in your folder
2. A new window opens showing all detected fields
3. You see two panels:
   - **Left**: Available fields (from the PDF)
   - **Right**: Selected fields (will become Excel columns)

#### How to select fields:
- **Double-click** a field in the left panel to add it
- Or select and click **"Add →"** button
- Use **"Search:"** box to filter fields quickly
- Click a field to see its **sample value** in the preview below

#### How to organize:
- Select a field in the right panel
- Click **"Move Up ↑"** or **"Move Down ↓"** to reorder
- Click **"← Remove"** to remove unwanted fields
- The order here = the order in Excel columns

#### Save your configuration:
- Click **"Save Mapping"** when done
- Your configuration is saved automatically
- You won't need to do this again unless you want to change it!

### Step 4: Select Excel File
Click **"Browse File"** and choose where to save the Excel file.

### Step 5: Choose Sheet
- Select an existing sheet from dropdown
- Or choose **"[Create New Sheet]"**
- Click **"Refresh Sheets"** to reload sheet list

### Step 6: Convert!
Click **"Convert PDFs to Excel"**

Watch the progress window as it:
- ✓ Scans all PDFs
- ✓ Extracts your configured fields
- ✓ Writes to Excel
- ✓ Skips duplicates

### Step 7: View Results
Open your Excel file - you'll see:
- Column A: PDF Filename
- Columns B-N: Your selected fields
- Last Column: "Open Invoice" link

---

## 📝 Example Workflow

### Scenario: Processing Monthly Invoices

**Goal**: Extract Invoice Number, Date, Customer, and Total Amount

1. **Browse** to `C:\Invoices\December2024\`
2. **Configure Fields**:
   - Search for "invoice"
   - Add "Invoice Number"
   - Add "Invoice Date"
   - Search for "customer"
   - Add "Customer"
   - Search for "total"
   - Add "Total Amount"
   - Click Save Mapping
3. **Select Excel**: `C:\Reports\December_Invoices.xlsx`
4. **Choose Sheet**: "December 2024" (or create new)
5. **Convert**: Process all PDFs
6. **Result**: Excel with 6 columns:
   - PDF Filename
   - Invoice Number
   - Invoice Date
   - Customer
   - Total Amount
   - Path to Invoice

---

## 💡 Pro Tips

### Tip 1: Choose a Good Sample PDF
The first PDF in your folder is used as a sample. Make sure it:
- Contains all the fields you need
- Is a typical/representative invoice
- Has clear text (not a scanned image)

### Tip 2: Use Search
If you see 50+ fields, use the search box:
- Type "date" to find all date fields
- Type "amount" to find all amount fields
- Type "number" to find ID/reference numbers

### Tip 3: Preview Before Adding
Always check the "Sample Value" preview:
- Make sure it's extracting the right data
- Verify the field name matches the content
- If wrong, try a different field with similar name

### Tip 4: Reorder Strategically
Put the most important fields first:
- Invoice Number → Column B (right after filename)
- Date → Column C
- Amount → Near the end for easy totals

### Tip 5: Save Multiple Configurations
You can:
- Manually edit `field_mapping.json`
- Save different versions (e.g., `field_mapping_invoices.json`)
- Rename the file to switch between configs

---

## ❓ Troubleshooting

### "No fields detected"
- Your PDF might be a scanned image (no text layer)
- Try a different PDF from the folder
- Use OCR software first to add text layer

### "Field value is N/A"
- The field exists but has no value in that PDF
- Normal for optional fields
- Check if other PDFs have values

### "Wrong value extracted"
- Field detection might have matched wrong text
- Try a different field name from the list
- Some PDFs have duplicate field names

### "Configuration not saved"
- Check folder permissions
- Look for `field_mapping.json` in app folder
- Manually verify file was created

---

## 🔄 Updating Your Configuration

To change which fields are extracted:
1. Click **"Configure Fields"** again
2. Modify your selection
3. Click **"Save Mapping"**
4. Next conversion uses new configuration

---

## ⚙️ Advanced: Manual Configuration

You can manually edit `field_mapping.json`:

```json
{
  "fields": [
    "Invoice Number",
    "Invoice Date",
    "Customer",
    "Total Amount"
  ]
}
```

Just make sure field names match exactly what appears in your PDFs!

---

**Happy Converting! 🎉**
