import pandas as pd
from CopyExcelFormat import CopyExcel
storecode = [156,254,166,158,265,150,146,158,186,265,254,312]

def Curr_Sales(i):

    df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/Updated Dumps/Sales Data for python.xlsx',sheet_name='Current')
    df1 = df.pivot_table(index = 'Partner',columns = ['Brand','Material group12'], values ='Qty',aggfunc='sum')
    df1.loc[i].fillna(0).to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrSales collection/CurrSales_{}.xlsx'.format(i))


def Proj_Sales(i):
    df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/Updated Dumps/Sales Data for python.xlsx',sheet_name='Proj Aug-Oct')
    df1 = df.pivot_table(index = 'Partner',columns = ['Brand1','Cal. year / month','Material group2'], aggfunc='sum')
    print(df1.head())
    try:
        df2=df1.loc[i].fillna(0)
        print(df2)
        writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/ProjSl collection/Proj_{}.xlsx'.format(i),engine='xlsxwriter')
        df2.to_excel(writer)
        writer.save()
    except KeyError:
        print('{} is missing data'.format(i))
        

def Curr_Inv(i):
    df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/Updated Dumps/Consolidatin/consolidated_CurrInv_Aug.xlsx',sheet_name='Sheet1')
    df1 = df.pivot_table(index = 'Site',columns = ['Brand 2','Product'], aggfunc='sum')
    print(df1.head())
    df1.loc[i].fillna(0).to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrInv collection/CurrInv_{}.xlsx'.format(i))


       
def SOH(i):
    df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/Updated Dumps/SOH Feb Mar Apr May.xlsx',sheet_name='Sheet1')
    df1 = df.pivot_table(index = 'Store',columns = ['Brand ','Month','Product'], aggfunc='sum') #Brand has a space in the column name
    print(df1.head())
    df1.loc[i].fillna(0).to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/SOH collection/SOH_{}.xlsx'.format(i))



def DC(i):
    df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/DC Breakup for Billing.xlsx',sheet_name= 'Sheet1')
    df.set_index('Site Code',inplace=True)
    df1= pd.DataFrame(df, columns= ['PS','PT','RK','SNB','NS','NT','NK','Jeans'])
    df1.loc[i].to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/DS collection/DS_{}.xlsx'.format(i))


CopyExcel(storecode)

for i in storecode:
    DC(i)
    SOH(i)
    Curr_Inv(i)
    Proj_Sales(i)
    Curr_Sales(i)
