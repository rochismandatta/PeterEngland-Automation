import pandas as pd

df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/Sales_Data_NonBinary.xlsx',sheet_name='Proj')
df1 = df.pivot_table(index = 'Partner',columns = ['Brand1','Cal. year / month','MaterialGroup12'], aggfunc='sum')
print(df1.head())


##storecode = [2048,'P001']
##for i in storecode:
##    try:
##        df2=df1.loc[i]
##        df2.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/ProjSl collection/Proj_{}.xlsx'.format(i))
##    except KeyError:
##        print('{} is missing data'.format(i))
##        continue
##
##
##storecode = [2048,'P001',4507]
##for i in storecode:
##    df2=df1.loc[i]
##    print(df2)
##    writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/ProjSl collection/Proj_{}.xlsx'.format(i),engine='xlsxwriter')
##    df2.to_excel(writer)
##    writer.save()


storecode = [2048,'P001',4507]
for i in storecode:
    try:
        df2=df1.loc[i]
        print(df2)
        writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/ProjSl collection/Proj_{}.xlsx'.format(i),engine='xlsxwriter')
        df2.to_excel(writer)
        writer.save()
    except KeyError:
        print('{} is missing data'.format(i))
        

