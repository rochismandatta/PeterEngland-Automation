import pandas as pd

df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/SOH Feb Mar Apr.xlsx',sheet_name='Sheet1')
df1 = df.pivot_table(index = 'Store',columns = ['Brand ','Month','Product'], aggfunc='sum') #Brand has a space in the column name
print(df1.head())

storecode = [4507,2048]
for i in storecode:
    df2=df1.loc[i]
    df2.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/SOH collection/SOH_{}.xlsx'.format(i))
    
