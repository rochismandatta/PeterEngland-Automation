import pandas as pd

df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/DC Breakup for Billing.xlsx',sheet_name= 'Sheet1')
df.set_index('Site Code',inplace=True)
df1= pd.DataFrame(df, columns= ['PS','PT','RK','SNB','NS','NT','NK','Jeans'])
#df1.set_index('Site Code',inplace=True)
storecode=[4507,2048,'P001']

for i in storecode:
    df2=df1.loc[i]
    print(df2)
##    df2.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/DS dump files/DS_{}.xlsx'.format(i))
    
