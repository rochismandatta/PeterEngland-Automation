import pandas as pd

df=pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/Copy of DS_SS18_120 Days Sell Thru_BASE FILE.xlsx', sheet_name='Base Data')
df.set_index('Store Code', inplace = True)
df1 = df.pivot_table(index = ['Store Code','Product'],columns = ['Category','Fabric Design Type'],values = ['Sales.', 'Dispatch'],aggfunc= 'sum')

print(df.head())
df1.loc[2048].to_excel('TestingSS120.xlsx')
