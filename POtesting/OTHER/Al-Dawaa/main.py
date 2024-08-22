import PyPDF2
import openpyxl
import PyPDF2
import pdfplumber
import tabula
import pandas as pd
import re
from datetime import datetime, timedelta


# Set the path to your PDF file
pdf_path1 = r"4000058139     100689 - Copy.pdf"
csv_path1 = "output.xlsx"

# Initialize pdfplumber object
with pdfplumber.open(pdf_path1) as pdf:
    # Initialize an empty list to store DataFrames for each table
    dfs = []

    # Iterate over each page and extract the tables
    for page in pdf.pages:
        # Extract tables from the current page
        page_tables = page.extract_tables()

        # Convert each table on the page to a DataFrame and append to the list
        for table in page_tables:
            df = pd.DataFrame(table)
            df = df[df[0] != '']
            df = df[df[1] != 'Article No.']
            df[10] = df[10].str.replace('*', '')
            print(df)
            dfs.append(df)

# Set the path to save the Excel file


# Combine all DataFrames into a single DataFrame
combined_table = pd.concat(dfs, ignore_index=True)

# Create an Excel writer
with pd.ExcelWriter(csv_path1, engine='xlsxwriter') as writer:
    # Save the combined DataFrame to a single sheet in the Excel file
    combined_table.to_excel(writer, sheet_name='Sheet1', index=False, header=False)

print(f"Tables extracted and saved to {csv_path1}")
def delete_columns_except(filename, sheet_name):
    # Load the Excel file
    workbook = openpyxl.load_workbook(filename)

    # Select the desired worksheet
    sheet = workbook[sheet_name]

    # Define the columns to keep (2, 4, 6, 7, 8)
    columns_to_keep = [2, 3, 6, 8, 11]

    # Create a list of columns to delete
    columns_to_delete = [col for col in range(1, sheet.max_column + 1) if col not in columns_to_keep]

    # Iterate over the columns to delete in reverse order (to avoid shifting column indexes)
    for col in reversed(columns_to_delete):
        sheet.delete_cols(col)

    # Save the modified workbook
    workbook.save(filename)

    print(f"Columns except {columns_to_keep} have been deleted from {sheet_name} sheet in {filename}.")


# Usage example:
worksheet_name = 'Sheet1'  # Change this to the name of your worksheet
delete_columns_except(csv_path1, worksheet_name)
def swap_columns_in_excel(file_path, sheet_name):
    try:
        # Read the Excel file into a DataFrame, skipping the header row
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

        # Swap the first and second columns for all rows, including the first cell
        if len(df.columns) >= 2:
            first_col_copy = df[df.columns[0]].copy()
            df[df.columns[0]] = df[df.columns[1]]
            df[df.columns[1]] = first_col_copy

        # Save the modified DataFrame back to the same sheet in the Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        print(f"Columns swapped and saved back to sheet '{sheet_name}' in '{file_path}'.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

sheet_name = "Sheet1"
swap_columns_in_excel(csv_path1, sheet_name)

def delete_first_row_from_excel(file_path, sheet_name):
    try:
        # Read the Excel file into a DataFrame, skipping the first row
        df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=[0])

        # Save the modified DataFrame back to the same sheet in the Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"First row deleted from sheet '{sheet_name}' in '{file_path}'.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

sheet_name = "Sheet1"
#delete_first_row_from_excel(csv_path1, sheet_name)

#
def add_Given_days(nbr_of_days, date_string, format_string):
    # Convert the date string to a datetime object using the provided format
    date = datetime.strptime(date_string, format_string)

    # Add 10 days to the date
    new_date = date + timedelta(days=nbr_of_days)

    # Convert the new date back to a string using the provided format
    new_date_string = new_date.strftime(format_string)

    # Return the new date string
    return new_date_string
'''date_string = "[%pocreation]"
format_string = "%d.%m.%Y"
nbr_of_days = [%daysToAddToExpiryDate]
new_date_string = add_Given_days(nbr_of_days,date_string, format_string)
print(new_date_string) '''

def extract_purchase_order_details(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)

        page = reader.pages[0]
        text = page.extract_text()
        print(text)
        # Extract Purchase Order Number
        po_number_match = re.search(r'Purchase Order(\d+)', text)
        po_str = po_number_match.group(1).strip() if po_number_match else None

        # Extract Ship To Location (Buy-from Vendor)
        ship_to_match = re.search(r':([^\n]+) Destination Plant', text)
        ship_to_str = ship_to_match.group(1).strip() if ship_to_match else None

        # Extract Order Date
        order_date_match = re.search(r'Order Date\s*([\d.]+)', text)
        order_date_str = order_date_match.group(1).strip() if order_date_match else None

        return po_str, ship_to_str, order_date_str


po_str, shipto, aprrovalDate = extract_purchase_order_details(pdf_path1)

# Load the Excel file
excel_file = pd.ExcelFile(csv_path1)

# Create a new sheet for PO info
with pd.ExcelWriter(csv_path1, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    # Write the purchase order details to the new sheet
    po_info_df = pd.DataFrame({
        'Purchase Order Number': [po_str],
        'Ship to': [shipto],
        'Order Date': [aprrovalDate],
        'RDD': [add_Given_days(4,aprrovalDate,"%d.%m.%Y")],
        'Expiry Date': [add_Given_days(7,aprrovalDate,"%d.%m.%Y")]

    })
    po_info_df.to_excel(writer, sheet_name='po info', index=False)


print("PO No: " + (po_str or "Not Found"))
print("Ship to: " + (shipto or "Not Found"))
print("Approval Date: " + (aprrovalDate or "Not Found"))