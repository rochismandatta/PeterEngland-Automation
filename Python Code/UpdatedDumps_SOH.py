import pandas as pd

df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Data_DS/Updated Dumps/SOH Feb Mar Apr May.xlsx',sheet_name='Sheet1')
df1 = df.pivot_table(index = 'Store',columns = ['Brand ','Month','Product'], aggfunc='sum')
df1.loc[2404].to_excel('2404_SOH_updated.xlsx')
##print(df1.loc[2404])
