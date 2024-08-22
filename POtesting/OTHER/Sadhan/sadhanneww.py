import re
import tabula
import PyPDF2
import pandas as pd
from datetime import datetime

# Provide the path to your PDF file
pdf_path = "2311002875094.pdf"
csv_path1="output.xlsx"

area = [10, 10, 300, 500]  # Example values, adjust as needed

# Extract tables
tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)

mainDF=pd.DataFrame()
# Loop through each table and display its contents
for i, table in enumerate(tables, start=1):
    print(f"Table {i}:")
    #print(table)
    if table.columns[0]=='Article code':
        #print(table)
        table = table.drop(table[ pd.isna(table["Article code"])].index)
        mainDF = pd.concat([mainDF, table])
        
    print("\n")
mainDF = mainDF[mainDF.columns.intersection(['Article code','Quantity','Net purchase price','Total net purchase'])]
print(mainDF)


print(tables[0]["Order no."][0])
print(tables[0]["Order date"][0])
print(tables[1]["Unnamed: 1"][0])

mainDF['Total net purchase'] = mainDF['Total net purchase'].apply(lambda x: x.replace(" ",""))
mainDF.insert(0,"tmpp","-")

#mainDF[mainDF.columns[0]], mainDF[mainDF.columns[4]] = mainDF[mainDF.columns[4]].copy(), mainDF[mainDF.columns[0]].copy()


mainDF.to_excel(csv_path1,index=False,header=False)

def extract_purchase_order_details(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        # Extract text from the first page
        first_page = reader.pages[0]
        first_page_text = first_page.extract_text()
        # Extract text from the last page
        last_page = reader.pages[-1]
        last_page_text = last_page.extract_text()
        # Combine the text from both pages
        text = first_page_text + "\n" + last_page_text
        print(text)
        # Extract text after "contractAd. ch."
        contract_ad_match = re.search(r'Order no*Ordered from', text,
                                      re.DOTALL)
        contract_ad_text = contract_ad_match.group(1) if contract_ad_match else None

        # Extract order number, order date, and supplier code


        # Extract text after "Asked delivery date Delivery deadline"
        delivery_date_match = re.search(r'Delivery date(.*)Delivery deadline date ', text, re.DOTALL)
        delivery_date_text = delivery_date_match.group(1) if delivery_date_match else None

        rdd_match = re.search(r'(\d{2}/\d{2}/\d{2} \d{2}:\d{2})', delivery_date_text)
        rdd = rdd_match.group(1) if rdd_match else None

        expiry_date_match = re.search(r'Delivery deadline date(.*)\d\d:\d\d',text,re.DOTALL)
        expiry_date_text = expiry_date_match.group(1) if expiry_date_match else None

        #print(expiry_date_text)
        # Extract RDD (Requested Delivery Date) and Expiry
        expiry_match = re.search(r'(\d{2}/\d{2}/\d{2} \d{2}:\d{2})', expiry_date_text)
        expiry = expiry_match.group(1) if expiry_match else None
        

        return  rdd,expiry



rdd,expiry= extract_purchase_order_details(pdf_path)
print(rdd,expiry)

def convert_date_format(input_date, input_format, output_format):
    try:
        # Parse the input date string into a datetime object using the input format
        datetime_obj = datetime.strptime(input_date, input_format)

        # Convert the datetime object to the desired output format
        output_date = datetime_obj.strftime(output_format)
        return output_date

    except ValueError:
        return "Invalid input date format"

# Create a new sheet for PO info
with pd.ExcelWriter(csv_path1, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    # Write the purchase order details to the new sheet
    po_info_df = pd.DataFrame({
        'Purchase Order Number': [tables[0]["Order no."][0]],
        'WH Address': [tables[1]["Unnamed: 1"][0]],
        'PO Date': [convert_date_format(tables[0]["Order date"][0].split(" ")[0],'%d/%m/%y','%d.%m.%Y')],
        'Not Before Date': [convert_date_format(rdd.split(" ")[0],'%d/%m/%y','%d.%m.%Y')],
        'Not After Date': [convert_date_format(expiry.split(" ")[0],'%d/%m/%y','%d.%m.%Y')],


    })
    po_info_df.to_excel(writer, sheet_name='po info', index=False)