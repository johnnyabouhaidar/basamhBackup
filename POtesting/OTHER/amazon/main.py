import re
import PyPDF2
import tabula
import pandas as pd
import openpyxl

#pdf_path =  r"[?Lookup("attachmentfiles", "[*CURRENT_LOOP_NUMBER]", "Item No", "Full Path")]"
pdf_path =  "6V9LERFX.pdf"
# Set the path to save the Excel file
#excel_path =  r"[%tmpoutputprocessingfolder]\amazonoutput.xlsx"
excel_path =  r"amazonoutput.xlsx"

tables = []

# Extract tables from all pages
all_tables = tabula.read_pdf(pdf_path, pages="all", encoding="latin1")
print(all_tables)

# Iterate over each table, starting from the third table
for i, table in enumerate(all_tables, start=1):
    if i >= 3:
        tables.append(table)

# Combine the extracted tables into a single DataFrame
combined_table = pd.concat(tables, ignore_index=True)



# Save the DataFrame as an Excel file
combined_table.to_excel(excel_path, index=False, header=False)

def extract_purchase_order_details(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)

        page = reader.pages[0]
        text = page.extract_text()

        # Extract Purchase Order Number
        po_number = re.search(r'PO:\s*(\w+)', text)
        po_str = po_number.group(1) if po_number else None

        # Extract Address
        ship_to_location = re.search(r'Ship to location\s*([^\n]+)', text)
        ship_to_location_str = ship_to_location.group(1).strip() if ship_to_location else None

        # Extract Ordered On Date
        ordered_on = re.search(r'Ordered On(\d{2}/\d{2}/\d{4})', text)
        ordered_on_str = ordered_on.group(1) if ordered_on else None

        # Extract Ship Window Dates
        ship_window = re.search(r'Ship window(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', text)
        ship_window_str = ship_window.group(1) + ' - ' + ship_window.group(2) if ship_window else None

        return po_str, ship_to_location_str, ordered_on_str, ship_window_str


# Extract purchase order details
po_number, ship_to_location, ordered_on, ship_window = extract_purchase_order_details(pdf_path)


# Load the Excel file
excel_file = pd.ExcelFile(excel_path)

# Create a new sheet for PO info
with pd.ExcelWriter(excel_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    # Write the purchase order details to the new sheet
    po_info_df = pd.DataFrame({
        'Purchase Order Number': [po_number],
        'Ship to location': [ship_to_location],
        'Ordered On': [ordered_on],
        'Ship Window': [ship_window],
    })
    po_info_df.to_excel(writer, sheet_name='po info', index=False)

def remove_empty_rows_from_sheet(filepath, sheet_name):
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(filepath)

        # Get the specified sheet
        sheet = wb[sheet_name]

        # Iterate through rows in reverse order to avoid index issues after deletion
        for row in reversed(list(sheet.iter_rows(min_row=1, min_col=1, max_col=1))):
            if not row[0].value:
                sheet.delete_rows(row[0].row)

        # Save the modified workbook back to the same file
        wb.save(filepath)

        print(f"Empty rows removed from '{sheet_name}' sheet and saved successfully.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Example usage:
remove_empty_rows_from_sheet(excel_path, "Sheet1")