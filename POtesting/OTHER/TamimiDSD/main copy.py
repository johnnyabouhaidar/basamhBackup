import re
import PyPDF2
import openpyxl
import pdfplumber
import pandas as pd
from datetime import datetime, timedelta
import fitz


pdf_path = r"[?Lookup("attachmentfiles", "[*CURRENT_LOOP_NUMBER]", "Item No", "Full Path")]"
xlsx_path = r"[%tmpoutputprocessingfolder]\output.xlsx"

#pdf_path = r"4732_TF01_2024_104377.pdf"
#xlsx_path = 'output.xlsx'
outputmasked = r"tmpmasked.pdf"
pdf_path_for_header=pdf_path


pdf_path_for_header= pdf_path
def mask_region_in_pdf(input_path, output_path, x1, y1, x2, y2):
    pdf_document = fitz.open(input_path)
    

    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        
        # Define a rectangle based on coordinates
        rect = fitz.Rect(x1, y1, x2, y2)
        
        # Create a redaction annotation on the specified region
        redact_annot = page.add_redact_annot(quad=rect,fill=(255,255,255))
        redact_annot.update()
        page.apply_redactions()
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        
    
    pdf_document.save(outputmasked)
    
    pdf_document.close()

# Usage example
input_file_path = pdf_path
output_file_path = outputmasked
x1, y1, x2, y2 = 415, 50, 140, 900  # Replace with your coordinates

mask_region_in_pdf(input_file_path, output_file_path, x1, y1, x2, y2)

pdf_path = outputmasked


def extract_data_from_pdf(pdf_path):
    data = []
    start_flag = False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            
            lines = text.split('\n')
            
            
            for line in lines:
                
                
                if line.startswith('JEDDAH'):
                    start_flag = True
                    continue

                if True:
                    row_data = line.split(' ')
                    data.append(row_data)
                    
                    
    print(data)
    return data


def save_data_to_xlsx(data, xlsx_path, sheet_name='Sheet1'):
    # Write the DataFrame to the specified sheet in the XLSX file
    data_df = pd.DataFrame(data)
    data_df.to_excel(xlsx_path, index=False, header=False, sheet_name=sheet_name)


# Create an empty XLSX file to ensure it exists
#save_data_to_xlsx([], xlsx_path)

# Extract data from the PDF
extracted_data = extract_data_from_pdf(pdf_path)

# Save the extracted data to an XLSX file in the "Sheet1" sheet
save_data_to_xlsx(extracted_data, xlsx_path, sheet_name='Sheet1')
def modify_even_rows(xlsx_path):
    df = pd.read_excel(xlsx_path, header=None)

    modified_data = []
    for i, row in df.iterrows():
        if len(row) > 0 and str(row[0])[0].isdigit() and ' ' in str(row[0]):
            modified_row = str(row[0]).split(' ')
            if modified_row[1].isdigit():
                modified_row = modified_row + row[1:].tolist()
            else:
                modified_row.append(row[1])
            modified_data.append(modified_row)
        else:
            modified_data.append(row.tolist())
        

    df_modified = pd.DataFrame(modified_data)
    df_modified.to_excel(xlsx_path, index=False, header=False)


# Call the function to modify even rows with space-separated cells
#modify_even_rows(xlsx_path)


def remove_rows_tuning(xlsx_path):
    df = pd.read_excel(xlsx_path, header=None)

    modified_data = []
    skip = False
    FirstLine = False

    for i, row in df.iterrows():
        if str(row[0]) == ".":
            skip = True
            FirstLine = True

        if skip != True and "___" not in str(row[0]):
            modified_data.append(row.tolist())

        if "Description" in str(row[0]):
            skip = False
        

    df_modified = pd.DataFrame(modified_data)
    df_modified.to_excel(xlsx_path, index=False, header=False)


#remove_rows_tuning(xlsx_path)

def find_word_before_keyword(sentence, keyword):
    pattern = r"(\w+)\s+" + re.escape(keyword)
    match = re.search(pattern, sentence)
    if match:
        return match.group(1)
    else:
        return None
def remove_non_numeric_rows(xlsx_path):
    df = pd.read_excel(xlsx_path, header=None)

    valid_rows = []
    for i, row in df.iterrows():
        if str(row[0]).isdigit():
            valid_rows.append(row.tolist())
        
        if "Pack Size" in str(row[0]):
            unit=find_word_before_keyword(str(row[0]),"Pack Size")
            
            if len(unit)>3:
                unit = find_word_before_keyword(str(row[0]),"SKU")
            #print(unit)
            #valid_rows[-1]=valid_rows[-1].append(unit)
            valid_rows[-1].append(unit)
                

            
        
    
    df_valid = pd.DataFrame(valid_rows)
    #print(df_valid)
    df_valid.to_excel(xlsx_path, index=False, header=False)


# Call the function to remove non-numeric rows from the XLSX file
remove_non_numeric_rows(xlsx_path)

def delete_columns_except(filename, sheet_name):
    # Load the Excel file
    workbook = openpyxl.load_workbook(filename)

    # Select the desired worksheet
    sheet = workbook[sheet_name]

    # Define the columns to keep
    #columns_to_keep = [1, 2, 4, 8, 9,12]
    columns_to_keep = [1,2,4,5,7]

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
delete_columns_except(xlsx_path, worksheet_name)
def swap_columns_in_excel(file_path, sheet_name):
    try:
        # Read the Excel file into a DataFrame without headers
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        df = df.assign(F=None)
        # Swap the first and second columns
        if len(df.columns) >= 2:
            df[df.columns[0]], df[df.columns[5]] = df[df.columns[5]].copy(), df[df.columns[0]].copy()
            df[df.columns[1]], df[df.columns[5]] = df[df.columns[5]].copy(), df[df.columns[1]].copy()
            df[df.columns[2]], df[df.columns[5]] = df[df.columns[5]].copy(), df[df.columns[2]].copy()
            df[df.columns[3]], df[df.columns[4]] = df[df.columns[4]].copy(), df[df.columns[3]].copy()
            df[0]="-"

        # Reset the index to start from 0
        #df.reset_index(drop=True, inplace=True)

        # Save the modified DataFrame back to the same sheet in the Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        print(f"Columns swapped and saved back to sheet '{sheet_name}' in '{file_path}'.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
sheet_name = "Sheet1"
swap_columns_in_excel(xlsx_path, sheet_name)

def swap_columns_3_and_4_in_excel(file_path, sheet_name):
    try:
        # Read the Excel file into a DataFrame without headers
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

        # Swap the third and fourth columns
        if len(df.columns) >= 4:
            df[df.columns[2]], df[df.columns[3]] = df[df.columns[3]].copy(), df[df.columns[2]].copy()

        # Reset the index to start from 0
        df.reset_index(drop=True, inplace=True)

        # Save the modified DataFrame back to the same sheet in the Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        print(f"Columns 3 and 4 swapped and saved back to sheet '{sheet_name}' in '{file_path}'.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Example usage:
sheet_name = "Sheet1"
#swap_columns_3_and_4_in_excel(xlsx_path, sheet_name)
def extract_purchase_order_details(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        page = reader.pages[0]
        text = page.extract_text()

        page2 = reader.pages[-1]
        text2 = page2.extract_text()
        print(text2)

        # Extract Purchase Order Number
        po_number = re.search(r'DSD ORDER:\s*([\d-]+[A-Z])', text)        
        po_str = po_number.group(1) if po_number else None
        if po_str is None:
            po_number = re.search(r'DSD ORDER:\s*(.*)', text)        
            po_str = po_number.group(1) if po_number else None

        # Extract Address (Ship To)
        ship_to = re.search(r'Store\s*(.*?)\s*(?:Dept\s*:|VAT No\s*\.|[_]+|$)', text, re.DOTALL)
        ship_to_str = ship_to.group(1).strip().split('-')[0] if ship_to else None

        # Extract PO Release Date
        po_release_date = re.search(r'Order Date :\s*([0-9./]+)', text)
        po_release_date_str = po_release_date.group(1) if po_release_date else None

        # Extract Supplier
        supplier = re.search(r'Supplier:\s*(\d+)', text)
        supplier_str = supplier.group(1) if supplier else None

        
        rdd = re.search(r'on\s*([0-9./]+)', text)
        rdd_str = rdd.group(1) if rdd else None

        exp = re.search(r'until\s*([0-9./]+)', text)
        exp_str = exp.group(1) if exp else None
        if exp_str==".":
            expp = datetime.now()+ timedelta(days=4)
            exp_str = expp.strftime("%d.%m.%Y")

        return po_str, ship_to_str, po_release_date_str, supplier_str,rdd_str,exp_str

# Example: Extract purchase order details from the PDF
po_number, ship_to, po_release_date, supplier,rdd,exp = extract_purchase_order_details(pdf_path_for_header)
# Load the Excel file
excel_file = pd.ExcelFile(xlsx_path)

def convert_date_format(input_date, input_format, output_format):
    try:
        # Parse the input date string into a datetime object using the input format
        datetime_obj = datetime.strptime(input_date, input_format)

        # Convert the datetime object to the desired output format
        output_date = datetime_obj.strftime(output_format)
        return output_date

    except ValueError:
        return "Invalid input date format"
def remove_last_dot(text):
    if text.endswith('.'):
        return text[:-1]
    else:
        return text

# Create a new sheet for PO info
with pd.ExcelWriter(xlsx_path, mode='a', engine='openpyxl') as writer:
    # Write the purchase order details to the new sheet
    
    po_info_df = pd.DataFrame({
        'Purchase Order Number': [po_number],
        'Ship To': [ship_to],
        'PO Release Date': [po_release_date],
        'RDD': [rdd],
        'Expiry': [remove_last_dot(exp)],

    })
    po_info_df.to_excel(writer, sheet_name='po info', index=False)

# Print the extracted purchase order details
if po_number:
    print("Purchase Order Number:", po_number)
else:
    print("Purchase Order Number not found.")

if ship_to:
    print("Ship To:", ship_to)
else:
    print("Ship To not found.")

if po_release_date:
    print("PO Release Date:", po_release_date)
else:
    print("PO Release Date not found.")

if rdd:
    print("PO RDD:", rdd)
else:
    print("PO RDD not found.")    
    
if exp:
    print("PO exp:", exp)
else:
    print("PO exp not found.")     

