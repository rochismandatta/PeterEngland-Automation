## Stock Productivity Code compiled


import pandas as pd
storecode = [4292,4257]

## Curr Inv code
df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/May Opening Inventory.xlsx',sheet_name='Sheet1')
df1 = df.pivot_table(index = 'Site',columns = ['Brand1','Product'], aggfunc='sum')
print(df1.head())


for i in storecode:
    df2=df1.loc[i]
    print(df2)
    df2.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrInv collection/CurrInv_{}.xlsx'.format(i))
    
## Curr Sales code

df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/Sales_Data_NonBinary.xlsx',sheet_name='Sheet1')
df1 = df.pivot_table(index = 'Partner',columns = ['Brand','Material group'], aggfunc='sum')
print(df1.head())


for i in storecode:
    df2=df1.loc[i]
    df2.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrSales collection/CurrSales_{}.xlsx'.format(i))
    

## Proj Sales code

df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/Sales_Data_NonBinary.xlsx',sheet_name='Proj')
df1 = df.pivot_table(index = 'Partner',columns = ['Brand1','Cal. year / month','MaterialGroup12'], aggfunc='sum')
print(df1.head())

for i in storecode:
    try:
        df2=df1.loc[i]
        print(df2)
        writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/ProjSl collection/Proj_{}.xlsx'.format(i),engine='xlsxwriter')
        df2.to_excel(writer)
        writer.save()
    except KeyError:
        print('{} is missing data'.format(i))
        
## Average Inventory

df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/SOH Feb Mar Apr.xlsx',sheet_name='Sheet1')
df1 = df.pivot_table(index = 'Store',columns = ['Brand ','Month','Product'], aggfunc='sum') #Brand has a space in the column name
print(df1.head())


for i in storecode:
    df2=df1.loc[i]
    df2.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/SOH collection/SOH_{}.xlsx'.format(i))
    
## DC code



df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/DC Breakup for Billing.xlsx',sheet_name= 'Sheet1')
df.set_index('Site Code',inplace=True)
df1= pd.DataFrame(df, columns= ['PS','PT','RK','SNB','NS','NT','NK','Jeans'])
#df1.set_index('Site Code',inplace=True)


for i in storecode:
    df2=df1.loc[i]
    print(df2)
    df2.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/DS collection/DS_{}.xlsx'.format(i))
    



