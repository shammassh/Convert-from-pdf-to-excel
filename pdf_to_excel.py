import os
from pathlib import Path
import openpyxl
from openpyxl import Workbook
import pdfplumber
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

def get_pdf_files(folder_path):
    """
    Get all PDF files from the specified folder.
    
    Args:
        folder_path: Path to the folder containing PDF files
        
    Returns:
        List of PDF file paths
    """
    pdf_files = []
    
    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' does not exist.")
        return pdf_files
    
    for file in os.listdir(folder_path):
        if file.lower().endswith('.pdf'):
            pdf_files.append(os.path.join(folder_path, file))
    
    return sorted(pdf_files)

def extract_total_amount(pdf_path):
    """
    Extract the total amount from a PDF file by looking for 'Total Amount' column.
    
    Args:
        pdf_path: Full path to the PDF file
        
    Returns:
        Total amount as string or 'N/A' if not found
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Try to extract tables from all pages
            for page in pdf.pages:
                tables = page.extract_tables()
                
                # Check if tables exist
                if tables:
                    for table in tables:
                        if not table:
                            continue
                        
                        # Look for 'Total Amount' in the table
                        for row_idx, row in enumerate(table):
                            if not row:
                                continue
                            
                            # Check each cell for 'Total Amount'
                            for col_idx, cell in enumerate(row):
                                if cell and 'Total Amount' in str(cell):
                                    # Try to find the value in the same column
                                    for data_row in table[row_idx + 1:]:
                                        if data_row and len(data_row) > col_idx:
                                            value = data_row[col_idx]
                                            if value and str(value).strip():
                                                # Clean and return the value
                                                cleaned = str(value).replace(',', '').replace('USD', '').strip()
                                                if re.match(r'^[0-9.]+$', cleaned):
                                                    return cleaned
                
                # Extract text and look for the pattern
                text = page.extract_text() or ""
                
                # Look for "Total Amount" followed by optional due date and USD amount
                # Pattern: Total Amount ... Due on ... USD 239.40
                pattern = r'Total Amount.*?USD\s*([0-9,]+\.?[0-9]*)'
                match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
                if match:
                    amount = match.group(1).replace(',', '')
                    return amount
                
                # Alternative: Look for USD followed by amount near Total Amount
                lines = text.split('\n')
                for i, line in enumerate(lines):
                    if 'Total Amount' in line:
                        # Check next few lines for USD amount
                        for j in range(i, min(i + 5, len(lines))):
                            usd_match = re.search(r'USD\s*([0-9,]+\.?[0-9]*)', lines[j])
                            if usd_match:
                                amount = usd_match.group(1).replace(',', '')
                                return amount
                
                # Fallback patterns
                patterns = [
                    r'Total Amount[:\s]+USD\s*([0-9,]+\.?[0-9]*)',
                    r'Total Amount[:\s]+([0-9,]+\.?[0-9]*)',
                ]
                
                for pattern in patterns:
                    match = re.search(pattern, text, re.IGNORECASE)
                    if match:
                        amount = match.group(1).replace(',', '')
                        return amount
            
            return 'N/A'
            
    except Exception as e:
        print(f"   ⚠ Error reading {os.path.basename(pdf_path)}: {str(e)}")
        return 'Error'

def write_to_excel(pdf_data, excel_path):
    """
    Write PDF filenames, total amounts, and hyperlinks to an Excel file.
    Prevents duplicate entries by checking existing data.
    
    Args:
        pdf_data: List of tuples (filename, total_amount, full_path)
        excel_path: Path where the Excel file will be saved
    """
    try:
        # Check if Excel file already exists
        existing_files = set()
        if os.path.exists(excel_path):
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
                
                # Collect existing filenames by iterating only through cells with values
                # This completely ignores empty/deleted rows
                for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
                    cell_value = row[0].value
                    if cell_value and str(cell_value).strip():
                        existing_files.add(cell_value)
                
            except PermissionError:
                print(f"\n❌ ERROR: Cannot open '{excel_path}'")
                print("   The file is currently open in another program.")
                print("   Please close the file and try again.\n")
                return False
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "PDF Files"
            # Add headers
            ws['A1'] = "PDF Filename"
            ws['B1'] = "Total Amount"
            ws['C1'] = "Path to Invoice"
            ws['A1'].font = openpyxl.styles.Font(bold=True)
            ws['B1'].font = openpyxl.styles.Font(bold=True)
            ws['C1'].font = openpyxl.styles.Font(bold=True)
        
        # Filter out duplicates
        new_data = [(name, amount, path) for name, amount, path in pdf_data if name not in existing_files]
        duplicates_count = len(pdf_data) - len(new_data)
        
        if duplicates_count > 0:
            print(f"\n⚠ Skipped {duplicates_count} duplicate file(s)")
        
        if not new_data:
            print("\n⚠ No new files to add (all files already exist in Excel)")
            return True
        
        # Find the next empty row
        start_row = ws.max_row + 1 if ws.max_row > 1 else 2
        
        # Write only new PDF data
        for idx, (pdf_name, total_amount, pdf_path) in enumerate(new_data, start=start_row):
            ws[f'A{idx}'] = pdf_name
            ws[f'B{idx}'] = total_amount
            ws[f'C{idx}'].hyperlink = pdf_path
            ws[f'C{idx}'].value = "Open Invoice"
            ws[f'C{idx}'].font = openpyxl.styles.Font(color="0563C1", underline="single")
        
        # Adjust column widths
        ws.column_dimensions['A'].width = 40
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 20
        
        # Save the workbook
        try:
            wb.save(excel_path)
            print(f"\n✓ Successfully wrote {len(new_data)} PDF file(s) to {excel_path}")
            if duplicates_count > 0:
                print(f"  ({duplicates_count} duplicate(s) skipped)")
            return True
        except PermissionError:
            print(f"\n❌ ERROR: Cannot save to '{excel_path}'")
            print("   The file is currently open in another program (Excel, etc.)")
            print("   Please close the file and try again.\n")
            
            # Offer to save with a different name
            retry = input("Would you like to save with a different filename? (yes/no): ").strip().lower()
            if retry in ['yes', 'y']:
                new_path = input("Enter new Excel file path: ").strip()
                if not new_path.lower().endswith('.xlsx'):
                    new_path += '.xlsx'
                return write_to_excel(pdf_data, new_path)
            return False
            
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        return False

def main():
    """
    Main function to run the PDF to Excel application with GUI.
    """
    app = PDFtoExcelApp()
    app.mainloop()

class PDFtoExcelApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("PDF to Excel Converter")
        self.geometry("700x700")
        self.resizable(False, False)
        
        # Configure style
        self.configure(bg="#f0f0f0")
        
        # Variables
        self.folder_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.available_sheets = []
        
        # Create GUI elements
        self.create_widgets()
        
    def create_widgets(self):
        # Header
        header_frame = tk.Frame(self, bg="#2c3e50", height=80)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame, 
            text="PDF to Excel Converter",
            font=("Arial", 24, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        title_label.pack(pady=20)
        
        # Main content frame
        content_frame = tk.Frame(self, bg="#f0f0f0")
        content_frame.pack(fill="both", expand=True, padx=30, pady=30)
        
        # Folder selection
        folder_frame = tk.LabelFrame(
            content_frame,
            text="Select PDF Folder",
            font=("Arial", 12, "bold"),
            bg="#f0f0f0",
            fg="#2c3e50",
            padx=15,
            pady=15
        )
        folder_frame.pack(fill="x", pady=(0, 20))
        
        folder_entry = tk.Entry(
            folder_frame,
            textvariable=self.folder_path,
            font=("Arial", 10),
            width=50
        )
        folder_entry.pack(side="left", padx=(0, 10))
        
        folder_btn = tk.Button(
            folder_frame,
            text="Browse Folder",
            command=self.browse_folder,
            font=("Arial", 10),
            bg="#3498db",
            fg="white",
            cursor="hand2",
            padx=15,
            pady=5
        )
        folder_btn.pack(side="left")
        
        # Excel file selection
        excel_frame = tk.LabelFrame(
            content_frame,
            text="Select Excel File",
            font=("Arial", 12, "bold"),
            bg="#f0f0f0",
            fg="#2c3e50",
            padx=15,
            pady=15
        )
        excel_frame.pack(fill="x", pady=(0, 20))
        
        excel_entry = tk.Entry(
            excel_frame,
            textvariable=self.excel_path,
            font=("Arial", 10),
            width=50
        )
        excel_entry.pack(side="left", padx=(0, 10))
        
        excel_btn = tk.Button(
            excel_frame,
            text="Browse File",
            command=self.browse_excel,
            font=("Arial", 10),
            bg="#3498db",
            fg="white",
            cursor="hand2",
            padx=15,
            pady=5
        )
        excel_btn.pack(side="left")
        
        # Sheet selection
        sheet_frame = tk.LabelFrame(
            content_frame,
            text="Select Excel Sheet",
            font=("Arial", 12, "bold"),
            bg="#f0f0f0",
            fg="#2c3e50",
            padx=15,
            pady=15
        )
        sheet_frame.pack(fill="x", pady=(0, 20))
        
        self.sheet_combo = ttk.Combobox(
            sheet_frame,
            textvariable=self.sheet_name,
            font=("Arial", 10),
            width=47,
            state="readonly"
        )
        self.sheet_combo.pack(side="left", padx=(0, 10))
        self.sheet_combo.set("Select a sheet or create new...")
        
        refresh_btn = tk.Button(
            sheet_frame,
            text="Refresh Sheets",
            command=self.load_sheets,
            font=("Arial", 10),
            bg="#9b59b6",
            fg="white",
            cursor="hand2",
            padx=15,
            pady=5
        )
        refresh_btn.pack(side="left", padx=(0, 10))
        
        clear_btn = tk.Button(
            sheet_frame,
            text="Clear Sheet Data",
            command=self.clear_sheet_data,
            font=("Arial", 10),
            bg="#e74c3c",
            fg="white",
            cursor="hand2",
            padx=15,
            pady=5
        )
        clear_btn.pack(side="left")
        
        # Progress frame
        self.progress_frame = tk.LabelFrame(
            content_frame,
            text="Progress",
            font=("Arial", 12, "bold"),
            bg="#f0f0f0",
            fg="#2c3e50",
            padx=15,
            pady=15
        )
        self.progress_frame.pack(fill="both", expand=True, pady=(0, 20))
        
        self.progress_text = tk.Text(
            self.progress_frame,
            height=10,
            font=("Consolas", 9),
            bg="white",
            fg="#2c3e50",
            state="disabled",
            wrap="word"
        )
        self.progress_text.pack(fill="both", expand=True)
        
        # Scrollbar for progress text
        scrollbar = tk.Scrollbar(self.progress_text)
        scrollbar.pack(side="right", fill="y")
        self.progress_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.progress_text.yview)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            mode='indeterminate',
            length=300
        )
        
        # Convert button
        self.convert_btn = tk.Button(
            content_frame,
            text="Convert PDFs to Excel",
            command=self.start_conversion,
            font=("Arial", 14, "bold"),
            bg="#27ae60",
            fg="white",
            cursor="hand2",
            padx=30,
            pady=10
        )
        self.convert_btn.pack()
        
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Folder with PDF Files")
        if folder:
            self.folder_path.set(folder)
            
    def browse_excel(self):
        file = filedialog.asksaveasfilename(
            title="Select Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.excel_path.set(file)
            self.load_sheets()
            
    def load_sheets(self):
        excel_path = self.excel_path.get()
        if not excel_path:
            messagebox.showwarning("Warning", "Please select an Excel file first!")
            return
            
        self.available_sheets = ["[Create New Sheet]"]
        
        if os.path.exists(excel_path):
            try:
                wb = openpyxl.load_workbook(excel_path)
                self.available_sheets.extend(wb.sheetnames)
                wb.close()
                self.sheet_combo['values'] = self.available_sheets
                self.sheet_combo.current(0)
            except Exception as e:
                messagebox.showerror("Error", f"Could not read Excel file:\n{str(e)}")
        else:
            self.sheet_combo['values'] = self.available_sheets
            self.sheet_combo.current(0)
            
    def clear_sheet_data(self):
        """Clear all data from the selected sheet (keeps headers)"""
        excel_path = self.excel_path.get()
        sheet_name = self.sheet_name.get()
        
        if not excel_path:
            messagebox.showwarning("Warning", "Please select an Excel file first!")
            return
        
        if not sheet_name or sheet_name == "Select a sheet or create new..." or sheet_name == "[Create New Sheet]":
            messagebox.showwarning("Warning", "Please select a sheet to clear!")
            return
        
        if not os.path.exists(excel_path):
            messagebox.showwarning("Warning", "Excel file does not exist yet!")
            return
        
        # Confirm action
        confirm = messagebox.askyesno(
            "Confirm Clear",
            f"Are you sure you want to clear all data from sheet '{sheet_name}'?\n\nThis will keep the headers but remove all entries."
        )
        
        if not confirm:
            return
        
        try:
            wb = openpyxl.load_workbook(excel_path)
            if sheet_name not in wb.sheetnames:
                messagebox.showerror("Error", f"Sheet '{sheet_name}' not found!")
                wb.close()
                return
            
            ws = wb[sheet_name]
            
            # Delete all rows except the header (row 1)
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row)
            
            wb.save(excel_path)
            wb.close()
            
            messagebox.showinfo("Success", f"Sheet '{sheet_name}' has been cleared!")
            self.log_message(f"\n✓ Cleared all data from sheet '{sheet_name}'")
            
        except PermissionError:
            messagebox.showerror("Error", "Cannot modify Excel file. Please close it in Excel and try again.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            
    def log_message(self, message):

        self.progress_text.config(state="normal")
        self.progress_text.insert("end", message + "\n")
        self.progress_text.see("end")
        self.progress_text.config(state="disabled")
        self.update_idletasks()
        
    def start_conversion(self):
        folder = self.folder_path.get()
        excel = self.excel_path.get()
        sheet = self.sheet_name.get()
        
        if not folder:
            messagebox.showerror("Error", "Please select a PDF folder!")
            return
            
        if not excel:
            messagebox.showerror("Error", "Please select an Excel file!")
            return
            
        if not sheet or sheet == "Select a sheet or create new...":
            messagebox.showerror("Error", "Please select a sheet!")
            return
            
        # Clear previous progress
        self.progress_text.config(state="normal")
        self.progress_text.delete(1.0, "end")
        self.progress_text.config(state="disabled")
        
        # Disable button and show progress bar
        self.convert_btn.config(state="disabled", text="Processing...")
        self.progress_bar.pack(pady=(10, 0))
        self.progress_bar.start()
        
        # Run conversion in a separate thread
        thread = threading.Thread(target=self.run_conversion, args=(folder, excel, sheet))
        thread.daemon = True
        thread.start()
        
    def run_conversion(self, folder_path, excel_path, sheet_name):
        try:
            # Ensure Excel file has .xlsx extension
            if not excel_path.lower().endswith('.xlsx'):
                excel_path += '.xlsx'
            
            # Get PDF files
            self.log_message(f"Searching for PDF files in: {folder_path}")
            pdf_files = get_pdf_files(folder_path)
            
            if not pdf_files:
                self.after(0, lambda: messagebox.showwarning("Warning", "No PDF files found!"))
                self.finish_conversion()
                return
            
            self.log_message(f"Found {len(pdf_files)} PDF file(s)\n")
            self.log_message("Extracting total amounts from PDFs...")
            self.log_message("-" * 50)
            
            # Extract data from each PDF
            pdf_data = []
            for i, pdf_path in enumerate(pdf_files, 1):
                filename = os.path.basename(pdf_path)
                self.log_message(f"{i}. Processing: {filename}...")
                total_amount = extract_total_amount(pdf_path)
                self.log_message(f"   Total: {total_amount}")
                pdf_data.append((filename, total_amount, pdf_path))
            
            self.log_message("-" * 50)
            self.log_message(f"\nWriting to Excel file: {excel_path}")
            self.log_message(f"Sheet: {sheet_name}")
            
            # Write to Excel
            success = write_to_excel_gui(pdf_data, excel_path, sheet_name, self.log_message)
            
            if success:
                self.log_message("\n✓ Operation completed successfully!")
                self.after(0, lambda: messagebox.showinfo("Success", f"Successfully processed {len(pdf_data)} PDF files!\n\nExcel file saved at:\n{excel_path}\nSheet: {sheet_name}"))
            else:
                self.log_message("\n⚠ Operation was not completed.")
                
        except Exception as e:
            self.log_message(f"\n❌ Error: {str(e)}")
            self.after(0, lambda: messagebox.showerror("Error", f"An error occurred:\n{str(e)}"))
        finally:
            self.finish_conversion()
            
    def finish_conversion(self):
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.convert_btn.config(state="normal", text="Convert PDFs to Excel")

def write_to_excel_gui(pdf_data, excel_path, sheet_name, log_func):
    """
    Write PDF filenames, total amounts, and hyperlinks to an Excel file (GUI version).
    """
    try:
        existing_files = set()
        create_new_sheet = (sheet_name == "[Create New Sheet]")
        
        if os.path.exists(excel_path):
            try:
                wb = openpyxl.load_workbook(excel_path)
                
                if create_new_sheet:
                    # Generate new sheet name
                    base_name = "PDF Files"
                    counter = 1
                    while base_name in wb.sheetnames:
                        base_name = f"PDF Files {counter}"
                        counter += 1
                    ws = wb.create_sheet(base_name)
                    sheet_name = base_name
                    log_func(f"Creating new sheet: {sheet_name}")
                else:
                    # Use existing sheet
                    if sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
                        # Collect existing filenames by iterating only through cells with values
                        # This completely ignores empty/deleted rows
                        for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
                            cell_value = row[0].value
                            if cell_value and str(cell_value).strip():
                                existing_files.add(cell_value)
                        
                        # Debug: Log what we found
                        if existing_files:
                            log_func(f"Found {len(existing_files)} existing file(s) in sheet")
                        else:
                            log_func("Sheet is empty, will add all files")
                    else:
                        ws = wb.create_sheet(sheet_name)
                
            except PermissionError:
                log_func(f"\n❌ ERROR: Cannot open '{excel_path}'")
                log_func("   The file is currently open in another program.")
                log_func("   Please close the file and try again.")
                return False
        else:
            wb = Workbook()
            ws = wb.active
            if create_new_sheet or not sheet_name:
                ws.title = "PDF Files"
                sheet_name = "PDF Files"
            else:
                ws.title = sheet_name
        
        # Add headers if new sheet or empty
        if ws.max_row == 1 or ws['A1'].value != "PDF Filename":
            ws['A1'] = "PDF Filename"
            ws['B1'] = "Total Amount"
            ws['C1'] = "Path to Invoice"
            ws['A1'].font = openpyxl.styles.Font(bold=True)
            ws['B1'].font = openpyxl.styles.Font(bold=True)
            ws['C1'].font = openpyxl.styles.Font(bold=True)
        
        new_data = [(name, amount, path) for name, amount, path in pdf_data if name not in existing_files]
        duplicates_count = len(pdf_data) - len(new_data)
        
        if duplicates_count > 0:
            log_func(f"\n⚠ Skipped {duplicates_count} duplicate file(s)")
        
        if not new_data:
            log_func("\n⚠ No new files to add (all files already exist in Excel)")
            return True
        
        start_row = ws.max_row + 1 if ws.max_row > 1 else 2
        
        for idx, (pdf_name, total_amount, pdf_path) in enumerate(new_data, start=start_row):
            ws[f'A{idx}'] = pdf_name
            ws[f'B{idx}'] = total_amount
            ws[f'C{idx}'].hyperlink = pdf_path
            ws[f'C{idx}'].value = "Open Invoice"
            ws[f'C{idx}'].font = openpyxl.styles.Font(color="0563C1", underline="single")
        
        ws.column_dimensions['A'].width = 40
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 20
        
        try:
            wb.save(excel_path)
            log_func(f"\n✓ Successfully wrote {len(new_data)} PDF file(s) to Excel")
            if duplicates_count > 0:
                log_func(f"  ({duplicates_count} duplicate(s) skipped)")
            return True
        except PermissionError:
            log_func(f"\n❌ ERROR: Cannot save to '{excel_path}'")
            log_func("   The file is currently open in another program.")
            return False
            
    except Exception as e:
        log_func(f"\n❌ Unexpected error: {e}")
        return False

if __name__ == "__main__":
    main()
