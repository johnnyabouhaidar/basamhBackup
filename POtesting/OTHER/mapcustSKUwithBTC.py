import pandas as pd
import numpy as np

master_data_file=r"[*MY_DOCUMENTS_DIRECTORY]Item Master Date.xlsx"
customer_SKU="[%tmpitemcode]"
mapped_BTC_code=""
division_factor=""
FOUND="FALSE"
sheet_name = '[%POvendorcurrent]'


# Replace 'your_excel_file.xlsx' with the path to your Excel file
xls = pd.ExcelFile(master_data_file)

# Create a dictionary of DataFrames, where each key is a sheet name
dfs = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}



# Specify the sheet name where you want to perform the lookup


# Lookup a value on the specified sheet

if sheet_name in dfs:
    sheet_df = dfs[sheet_name]
    print(sheet_df.iloc[:,0])
    try:
        result = sheet_df[sheet_df.iloc[:,0] == int(customer_SKU)]
    except:
        result = sheet_df[sheet_df.iloc[:,0] == customer_SKU]
    
    if not result.empty:
        mapped_BTC_code = int(result.iloc[0][1])
        if np.isnan(result.iloc[0][2]):
            division_factor = ""
        else:
            division_factor = result.iloc[0][2]
        FOUND="TRUE"
        
        print(str(mapped_BTC_code),division_factor)
    else:
        print("not found on {sheet_name}.")
else:
    print(f"{sheet_name} not found in the Excel file.")
