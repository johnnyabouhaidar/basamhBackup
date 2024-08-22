import pandas as pd
from datetime import datetime
input  = "M30201224000379.txt"
output = "output.xlsx"

def convert_date_format(input_date, input_format, output_format):
    try:
        # Parse the input date string into a datetime object using the input format
        datetime_obj = datetime.strptime(input_date, input_format)

        # Convert the datetime object to the desired output format
        output_date = datetime_obj.strftime(output_format)
        return output_date

    except ValueError:
        return "Invalid input date format"

items = [
    #[3,11], #PO
    #[17,23], #deliverydate
    #[62,68], #orderdate  
    [118,126],#basamhsku
    [109,118],#custsku

    [241,249],#qty
    [261,270],#unit price
    [272,285],#totalpricewoVAT
]




all_items=[]
f = open(input, "r")
txt = f.read()

info={}

info['PO Number']=txt[3:11]
info['WH Address']=txt[0:3]
info['Order Date']=convert_date_format(txt[62:68],'%y%m%d','%d.%m.%Y')
print(info)

for line in txt.split("\n"):
    print("----")
    tmp_line = []
    for item in items:
        #print(line[item[0]:item[1]])
        tmp_line.append(line[item[0]:item[1]])
    all_items.append(tmp_line)
    
    print("----")
print(all_items)
df = pd.DataFrame(all_items)
print(df)
df.to_excel(output,header=False,index=False)

with pd.ExcelWriter(output, mode='a', engine='openpyxl') as writer:
    # Write the purchase order details to the new sheet
    po_info_df = pd.DataFrame(info)
    po_info_df.to_excel(writer, sheet_name='po info', index=False)

