
import PyPDF2
import tabula
import pandas as pd
import re

# Set the path to your PDF file
pdf_path1 = r"C:\Users\rpauser\Desktop\POtesting\Panda.pdf"

# Initialize an empty list to store the extracted tables
tables = []
# print(len(tabula.read_pdf(pdf_path1, pages="all", lattice=True, multiple_tables=True, encoding="latin1",
#                                          area=[250.848, 29.088, 763.748, 1402.468])))
# Iterate over each page and extract the tables
for page in range(1, len(tabula.read_pdf(pdf_path1, pages="all", lattice=True, multiple_tables=True, encoding="latin1",
                                         area=[250.848, 29.088, 763.748, 1402.468]))):
    print(page)
    # Extract tables from the current page
    page_tables = tabula.read_pdf(pdf_path1, pages=page, lattice=True, multiple_tables=True, encoding="latin1",
                                  area=[250.848, 29.088, 763.748, 1402.468])
    # Append the extracted tables to the list
    tables.extend(page_tables)

# Combine the extracted tables into a single DataFrame
combined_table = pd.concat(tables, ignore_index=True)

# Set the path to save the CSV file
csv_path1 = "Panda4.xlsx"

# Save the DataFrame as an Excel file (CSV format)
combined_table.to_excel(csv_path1, sheet_name='Original', index=False, header=False)


def extract_purchase_order_details(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)

        page = reader.pages[1]
        text = page.extract_text()

        # Extract Purchase Order Number
        po_number = re.search(r'Purchase Order\s*-\s*(\d+)', text)
        po_str = "PO#" + po_number.group(1) if po_number else None

        # Extract Address
        address = re.search(r'Location Name\s*(.*?)\s*Supplier', text, re.DOTALL)
        address_str = address.group(1).strip() if address else None

        # Extract Not Before Date
        not_before_date = re.search(r'Not Before Date\s*([-0-9A-Za-z]+)', text)
        not_before_str = not_before_date.group(1) if not_before_date else None

        # Extract Not After Date
        not_after_date = re.search(r'Not After Date\s*([-0-9A-Za-z]+)', text)
        not_after_str = not_after_date.group(1) if not_after_date else None

        # Extract PO creation Date
        po_creation_date = re.search(r'PO Creation Date\s*([-0-9A-Za-z]+)', text)
        po_creation_str = po_creation_date.group(1) if not_after_date else None        

        #PO context
        PO_Context = re.search(r'PO Context\s*([-0-9A-Za-z]+)', text)
        PO_Context_str = PO_Context.group(1) if PO_Context else None
        return po_str, address_str, not_before_str, not_after_str, PO_Context_str,po_creation_str



# Extract purchase order details
po_number, address, not_before_date, not_after_date, PO_Context,po_creation = extract_purchase_order_details(pdf_path1)

# Load the Excel file
excel_file = pd.ExcelFile(csv_path1)

# Create a new sheet for PO info
with pd.ExcelWriter(csv_path1, mode='a', engine='openpyxl') as writer:
    # Write the purchase order details to the new sheet
    po_info_df = pd.DataFrame({
        'Purchase Order Number': [po_number],
        'Address': [address],
        'Not Before Date': [not_before_date],
        'Not After Date': [not_after_date],
        'PO Context': [PO_Context],
        'PO Creation': [po_creation]
    })
    po_info_df.to_excel(writer, sheet_name='po info', index=False)


# Print the extracted purchase order details
if po_number:
    print(po_number)
else:
    print("Purchase Order Number not found.")

if address:
    print("Address:", address)
else:
    print("Address not found.")

if not_before_date:
    print("Not Before Date:", not_before_date)
else:
    print("Not Before Date not found.")

if not_after_date:
    print("Not After Date:", not_after_date)
else:
    print("Not After Date not found.")

if PO_Context:
    print("PO Context:", PO_Context)
else:
    print("PO Context not found.")
