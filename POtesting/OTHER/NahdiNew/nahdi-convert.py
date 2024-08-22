import pandas as pd
from datetime import datetime
import tabula

input_file = r"nahdi_receipt.pdf"
output_file = r"output.xlsx"


df = tabula.read_pdf(input_path=input_file, pages='all', lattice=True)
final_df = pd.DataFrame()
header_df = pd.DataFrame(columns=['c1', 'c2'])

count = 0

for table in df:
    if 'S No.' in table.columns:
        final_df = pd.concat([final_df, table])

    if 'PO Status' in table.columns:
        PONumber = table.iloc[0][1]
        PODate = table.iloc[1][1]
        

    if 'WH Address' in table.columns:
        WHAddr = table.columns[1]
        PORDD = table.iloc[1][1]
        POExpiry = table.iloc[2][1]        

final_df = final_df.drop(['S No.'], axis=1)

final_df = final_df.drop(['Item Description'], axis=1)
final_df = final_df.drop(['Case Size'], axis=1)
final_df = final_df.drop(['Total VAT'], axis=1)

final_df.insert(0,"tmpp","-")
final_df.to_excel(output_file, index=False, header=False)

def convert_date_format(input_date, input_format, output_format):
    try:
        # Parse the input date string into a datetime object using the input format
        datetime_obj = datetime.strptime(input_date, input_format)

        # Convert the datetime object to the desired output format
        output_date = datetime_obj.strftime(output_format)
        return output_date

    except ValueError:
        return "Invalid input date format"


with pd.ExcelWriter(output_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    # Write the purchase order details to the new sheet
    po_info_df = pd.DataFrame({
        'Purchase Order Number': [PONumber],
        'WH Address': [WHAddr],
        'PO Date': [convert_date_format(PODate,'%Y-%m-%d','%d.%m.%Y')],
        'Not Before Date': [convert_date_format(PORDD,'%Y-%m-%d','%d.%m.%Y')],
        'Not After Date': [convert_date_format(POExpiry,'%Y-%m-%d','%d.%m.%Y')],


    })
    po_info_df.to_excel(writer, sheet_name='po info', index=False)
#header_df.to_csv(output_header_file, index=False, header=False)