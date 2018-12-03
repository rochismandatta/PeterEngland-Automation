import pandas as pd
def Curr_Inv(i):
    df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/Updated Dumps/Consolidatin/consolidated_CurrInv_Aug.xlsx',sheet_name='Sheet1')
    df1 = df.pivot_table(index = 'Site',columns = ['Brand 2','Product'], aggfunc='sum')
    print(df1.head())
    df1.loc[i].fillna(0).to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrInv collection/CurrInv_{}.xlsx'.format(i))

Curr_Inv(2097)
