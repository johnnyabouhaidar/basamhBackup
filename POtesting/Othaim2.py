import re
import pandas as pd
import pdfplumber

# Set the path to your PDF file
pdf_path1 = r"C:\Users\rpauser\Desktop\POtesting\Othaim 2.pdf"

# Initialize an empty list to store the extracted tables
tables = []

# Iterate over each page and extract the tables
with pdfplumber.open(pdf_path1) as pdf:
    for page in pdf.pages:
        table = page.extract_tables()[0]  # Extract the first table from each page
        tables.append(table)

# Combine the extracted tables into a single DataFrame
combined_table = pd.concat([pd.DataFrame(table) for table in tables], ignore_index=True)

# Set the path to save the Excel file
excel_path1 = "Othaim_output.xlsx"

# Save the DataFrame as an Excel file
combined_table.to_excel(excel_path1, sheet_name='Original', index=False, header=False)

def extract_purchase_order_details(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()

        # Extract Purchase Order Number
        po_number = re.search(r'P\.?O\.?\s*:? ?(\d+)', text, re.IGNORECASE)
        po_str = "Purchase Order - " + po_number.group(1) if po_number else None

        # Extract Entry Date
        entry_date = re.search(r'Entry Date\s*:\s*([\d/-]+)', text)
        entry_date_str = entry_date.group(1) if entry_date else None

        # Extract Expiry Date
        exp_date = re.search(r'Exp Date\s*:\s*([\d/-]+)', text)
        exp_date_str = exp_date.group(1) if exp_date else None

        # Extract Cancel Date
        canc_date = re.search(r'Canc Date\s*:\s*([\d/-]+)', text)
        canc_date_str = canc_date.group(1) if canc_date else None

        # Extract Vendor
        vendor = re.search(r'(?i)Vendor\s+No\s*:\s*\d+\s+Vendor:\s*(.*?)\s+City', text, re.DOTALL)
        vendor_str = vendor.group(1).strip() if vendor else None

        # Extract Store
        store = re.search(r'Store\s*:\s*(\d+)', text)
        store_str = store.group(1) if store else None

        return po_str, entry_date_str, exp_date_str, canc_date_str, vendor_str, store_str


# Extract purchase order details
po_number, entry_date, exp_date, canc_date, vendor, store = extract_purchase_order_details(pdf_path1)

# Load the Excel file
excel_file = pd.ExcelFile(excel_path1)

# Create a new sheet for PO info
with pd.ExcelWriter(excel_path1, mode='a', engine='openpyxl') as writer:
    # Write the purchase order details to the new sheet
    po_info_df = pd.DataFrame({
        'Purchase Order Number': [po_number],
        'Entry Date': [entry_date],
        'Expiry Date': [exp_date],
        'Cancel Date': [canc_date],
        'Vendor': [vendor],
        'Store': [store]
    })
    po_info_df.to_excel(writer, sheet_name='po info', index=False,header=False)

# Print the extracted purchase order details
if po_number:
    print(po_number)
else:
    print("Purchase Order Number not found.")
if entry_date:
    print("Entry Date:", entry_date)
else:
    print("Entry Date not found.")
if exp_date:
    print("Expiry Date:", exp_date)
else:
    print("Expiry Date not found.")
if canc_date:
    print("Cancel Date:", canc_date)
else:
    print("Cancel Date not found.")
if vendor:
    print("Vendor:", vendor)
else:
    print("Vendor not found.")
if store:
    print("Store:", store)
else:
    print("Store not found.")
