import pandas as pd
from SUM_SB import SB_CurrInv

df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/May Opening Inventory.xlsx',sheet_name='Sheet1')
df1 = df.pivot_table(index = 'Site',columns = ['Brand1','Product'], aggfunc='sum')
print(df1.head())

storecode = [4507,2048]
for i in storecode:
    df2=df1.loc[i]
    print(df2)
    ##    df2.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrInv collection/CurrInv_{}.xlsx'.format(i))
    SB_CurrInv(i)


