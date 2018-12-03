import pandas as pd

df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/Sales_Data_NonBinary.xlsx',sheet_name='Sheet1')
df1 = df.pivot_table(index = 'Partner',columns = ['Brand','Material group'], aggfunc='sum')
print(df1.head())

storecode = [4507,2048]
for i in storecode:
    df2=df1.loc[i]
    df2.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrSales collection/CurrSales_{}.xlsx'.format(i))
    
