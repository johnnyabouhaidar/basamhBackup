import pandas as pd
import math

input_material_number = "20000400"
input_quantity = "108"

df= pd.read_excel(r"C:\Users\rpauser\Documents\MasterConversionData.xlsx")

final_quantity = None
final_unit = None


#print(PC_df.iloc[0]['Material'])
try:
    PC_df = df.loc[(df['Material'] == input_material_number) & (df['Alternative Unit of Measure'] =="PC")]

    tmp_quantity=float(input_quantity.replace(",",""))/PC_df.iloc[0]['Denominator']
    print(tmp_quantity)
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
 
except:
    pass



'''print("Final Quantity:",final_quantity)
print("Final Unit:",final_unit)'''
if input_material_number == None:
    #RPAEngine.SetVar("newquantity", input_quantity)
    print(input_quantity)
    #RPAEngine.SetVar("newunit", "CAR")
    print("CAR")
else:
    #RPAEngine.SetVar("newquantity", final_quantity)
    print(final_quantity)
    #RPAEngine.SetVar("newunit", final_unit)
    print(final_unit)

