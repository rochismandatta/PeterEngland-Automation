import pandas as pd

storecode = [312,168,156,254,166,158,265,150,146,158,186,265,254]
def StoreName(i):
    df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/Pantaloons Diagnostic fill data sheets/DC Breakup for Billing.xlsx',sheet_name= 'Sheet1')
    df.set_index('Site Code',inplace=True)
    df = pd.DataFrame(df, columns= ['Store'])
    df = df.loc[i]
    print(df)
    df.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/StoreName collection/{}_storename.xlsx'.format(i))
    print('End of StoreName')
##storecode = [2422,2408]
##
##for i in storecode:
##    StoreName(i)
##

#END OF STORENAME
###############################################################################

def Budget(i):
    df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/new data files/1.Data/BANGLORE Final_FY-19 Sec budget.xlsx',sheet_name = 'Sheet1')
    df.set_index('Row Labels', inplace = True)
    df= df.loc[i]
    print(df)
    df.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/Budget Collection/{}_budget.xlsx'.format(i))
    print('End of Budget')
##storecode = [2422,2408,'P097']
##for i in storecode:
##    Budget(i)
    
#END OF BUDGET
################################################################################


def YTDSale(i):
    df= pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/new data files/1.Data/YTD1819.xlsx', sheet_name = 'Sheet1')
    df.set_index('Partner',inplace = True)
    df =df.loc[i]
    print(df)
    df.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/YTD collection/{}_YTD.xlsx'.format(i))
    print('End of YTD1819')
##storecode = [2422,2408,'P097']
##for i in storecode:
##    YTDSale(i)
#END OF YTD1819
################################################################################


def STD_table(i):
    df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/new data files/1.Data/DS_SS18_120 Days Sell Thru_BASE FILE_for python.xlsx', sheet_name = 'Base Data')
    df1 = df.pivot_table(index = 'Store Code', columns = ['ProdCat'], aggfunc='sum')
##    print(df1.loc[2422])
    df1.loc[i].to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/STD collection/{}_STD.xlsx'.format(i))
    print('End of STD Table')
##storecode = [2422,2408,'P097']
##for i in storecode:
##    STD_table(i)
#END OF STD
##################################################################################


def PE(i):
    if (isinstance(i, str) != True and i<1000):
        df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/new data files/PE/PE_SS SALE DUMP FY_17 AND FY_18.xlsx',sheet_name = 'Sheet1')
        df1 = df.pivot_table(index = 'SITE', columns = ['Season12','YR','BRAND_Real','Product'], aggfunc = 'sum')
##        df1.loc[i].to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/PE_collection/{}_PE.xlsx'.format(i),sheet_name = 'NewSeason')
        df2 = df.pivot_table(index = 'SITE', columns = ['All Seasons','YR','BRAND_Real','Product'], aggfunc = 'sum')
##        df2.loc[i].to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/PE_collection/{}_PE.xlsx'.format(i),sheet_name = 'FullYear')
        writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/PE_collection/{}_PE.xlsx'.format(i), engine = 'xlsxwriter')
        df1.loc[i].to_excel(writer, sheet_name = 'NewSeason')
        df2.loc[i].to_excel(writer, sheet_name = 'FullYear')
        writer.save()
        writer.close()
    else:
        df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/new data files/PE/PE Final L2L dump (CT, MX, PT, RL) FY-18 for python.xlsx', sheet_name = 'base file')
        df1 = df.pivot_table(index = 'Partner', columns = ['remakrs-season','FY','BRANDDD','product-nitya'], aggfunc = 'sum')
        df2 = df.pivot_table(index = 'Partner', columns = ['All seasons','FY','BRANDDD','product-nitya'], aggfunc = 'sum')
        writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/PE_collection/{}_PE.xlsx'.format(i), engine = 'xlsxwriter')
        df1.loc[i].to_excel(writer, sheet_name = 'NewSeason')
        df2.loc[i].to_excel(writer, sheet_name = 'FullYear')
        writer.save()
        writer.close()
    print('End of PE')
               

##storecode = [123,2422]
##for i in storecode:
##    PE(i)
#END of PE
#######################################################################################
    
def AS(i):
    df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/new data files/AS/AS_YTD 1st April to 31st Mar  FY17  FY18.xlsx',sheet_name = 'BASE FILE')
    df1 = df.pivot_table(index = 'Site Code', columns = ['SEASON REMARKS','NITYA BRAND','New Grp'], aggfunc = 'sum')
    df2 = df.pivot_table(index = 'Site Code', columns = ['All Seasons','NITYA BRAND','New Grp'], aggfunc = 'sum')
    writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/AS_collection/{}_AS.xlsx'.format(i), engine = 'xlsxwriter')
    try:
        df1.loc[i].to_excel(writer, sheet_name = 'NewSeason')
        df2.loc[i].to_excel(writer, sheet_name = 'FullYear')
        writer.save()
        writer.close()
    except KeyError:
        print('Store Code '+str(i)+' doesnt have AS')
    print('End of AS')
##AS(123)
##END OF AS
##########################################################################################


def LP(i):
    df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/new data files/LP/LP_YTD Sec Sales -Mar -2018 (FY 14 - FY 18) for python.xlsx',sheet_name = 'Data')
    df1 = df.pivot_table(index = 'Store Code', columns = ['REMARKS-SEASON','FY','NITYA BRAND','Product Category'], aggfunc = 'sum')
    df2 = df.pivot_table(index = 'Store Code', columns = ['All Seasons','FY','NITYA BRAND','Product Category'], aggfunc = 'sum')
    writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/LP_collection/{}_LP.xlsx'.format(i), engine = 'xlsxwriter')
    try:
        df1.loc[i].to_excel(writer, sheet_name = 'NewSeason')
        df2.loc[i].to_excel(writer, sheet_name = 'FullYear')
        writer.save()
        writer.close()
    except KeyError:
        print('StoreCode '+ str(i) +' doesnt have LP')
    print('End of LP')

##LP(123)
##END OF LP
##########################################################################################



def VH(i):
    df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/new data files/VH/VH_MERA SECONDARY YTD MAR 2018-N for python.xlsx',sheet_name = 'BASE SHEET')
    df1 = df.pivot_table(index = 'Store Code', columns = ['REMARKS-SEASON','NITYA BRAND','PRODUCT'], aggfunc = 'sum')
    df2 = df.pivot_table(index = 'Store Code', columns = ['All Seasons','NITYA BRAND','PRODUCT'], aggfunc = 'sum')
    writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/VH_collection/{}_VH.xlsx'.format(i), engine = 'xlsxwriter')
    try:
        df1.loc[i].to_excel(writer, sheet_name = 'NewSeason')
        df2.loc[i].to_excel(writer, sheet_name = 'FullYear')
        writer.save()
        writer.close()
    except KeyError:
        print('Store Code '+ str(i)+ ' doesnt have VH')
    print('End of VH')
##VH(123)
#END OF VH
############################################################################################


for i in storecode:
    StoreName(i)
    Budget(i)
    YTDSale(i)
    STD_table(i)
    PE(i)
##    VH(i)
##    LP(i)
##    AS(i)


