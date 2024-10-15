import os
import pdfplumber
import tkinter as tk
from tkinter import filedialog
import pandas as pd

def extract_table_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        tables = []
        notes = ""
        for page in pdf.pages:
            extracted_tables = page.extract_tables()
            page_text = page.extract_text()
            
            # Check if "Notes" section exists in the page
            if "Notes:" in page_text:
                # Extract text after "Notes:" until the end of the page
                notes += page_text.split("Notes:")[1].strip()
            
            for table in extracted_tables:
                # Find the row containing "STI Target"
                sti_target_row_index = None
                for i, row in enumerate(table):
                    if "STI Target (Depending on individual performan" in row:
                        sti_target_row_index = i
                        break
                
                if sti_target_row_index is not None:
                    # Merge next two cells after the first cell in the "STI Target" row
                    if len(table[sti_target_row_index]) >= 3 and all(cell is not None for cell in table[sti_target_row_index][1:3]):
                        table[sti_target_row_index][0] = " ".join([table[sti_target_row_index][0], table[sti_target_row_index][1], table[sti_target_row_index][2]])
                        del table[sti_target_row_index][1:3]

                    # Remove rows after the "STI Target" row
                    table = table[:sti_target_row_index+1]

                tables.append(table)
        
        # Append notes as a separate row after the last row of the last table
        if tables:
            last_table = tables[-1]
            last_table.append([notes])
        
        return tables

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Open folder dialog to select a folder containing PDF files
    folder_path = filedialog.askdirectory()

    if not folder_path:
        print("No folder selected. Exiting.")
        return

    # Create a DataFrame to hold all tables from all PDFs
    all_dfs = []

    # Process each PDF file in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            print(f"Processing {pdf_path}...")
            
            # Extract tables from the current PDF
            tables = extract_table_from_pdf(pdf_path)

            # Convert tables to DataFrame
            all_rows = []
            for table in tables:
                for row in table:
                    # Filter out "None" values from the row
                    filtered_row = [str(cell) for cell in row if cell is not None]
                    all_rows.append(filtered_row)

            df = pd.DataFrame(all_rows)

            # Append the DataFrame to the list
            all_dfs.append(df)

    # Write all DataFrames to a single Excel workbook
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if output_file_path:
        with pd.ExcelWriter(output_file_path) as writer:
            for idx, df in enumerate(all_dfs, start=1):
                df.to_excel(writer, sheet_name=f"Sheet{idx}", index=False, header=False)
        print(f"Filtered data from all PDFs saved to {output_file_path}")

if __name__ == "__main__":
    main()
