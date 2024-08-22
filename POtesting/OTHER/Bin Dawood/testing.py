import pandas as pd
import math

input_material_number = "10000568"
input_quantity = "2"

#df= pd.read_excel(r"[*MY_DOCUMENTS_DIRECTORY]MasterConversionData.xlsx")
df= pd.read_excel(r"C:\Users\rpauser\Documents\MasterConversionData.xlsx")

final_quantity = None
final_unit = None


#print(PC_df.iloc[0]['Material'])
PC_df = df.loc[(df['Material'] == input_material_number) & (df['Alternative Unit of Measure'] =="PC")]
TR_df = df.loc[(df['Material'] == input_material_number) & (df['Alternative Unit of Measure'] =="TR")]
print(PC_df)
print(TR_df)

tmp_quantity=float(input_quantity.replace(",",""))*(PC_df.iloc[0]['Denominator']/TR_df.iloc[0]['Denominator'])
print(tmp_quantity)
final_quantity = tmp_quantity
RPAEngine.SetVar("quantity", final_quantity)
try:
	
    '''
    PC_denom=PC_df.iloc[0]['Denominator']
    RPAEngine.SetVar("conversion_rate", PC_denom)
    #print(tmp_quantity)
    CAR_quantity=tmp_quantity
    #if isinstance(tmp_quantity,int):
    if tmp_quantity%1==0:
        #print(True)
        final_quantity=tmp_quantity
        final_unit="CAR"
    else:
        try:
            TR_df = df.loc[(df['Material'] == input_material_number) & (df['Alternative Unit of Measure'] =="TR")]
            TR_denom=TR_df.iloc[0]['Denominator']
            tmp_quantity = float(input_quantity.replace(",",""))/int(PC_denom/TR_denom)
            
            if tmp_quantity%1==0:
                
                final_quantity=tmp_quantity
                final_unit="TR"
            else:
                #print(final_quantity)
                
                final_quantity = math.floor(CAR_quantity)
                final_unit="CAR"
        except:
            final_quantity = math.floor(CAR_quantity)
            final_unit="CAR"
 '''
except:
    pass




'''print("Final Quantity:",final_quantity)
print("Final Unit:",final_unit)'''
'''
if input_material_number == None:
    new_quantity = int(input_quantity)
    RPAEngine.SetVar("newquantity", new_quantity)
    RPAEngine.SetVar("newunit", "CAR")
else:
    new_quantity=int(final_quantity)
    RPAEngine.SetVar("newquantity", new_quantity)
    RPAEngine.SetVar("newunit", final_unit)
'''
