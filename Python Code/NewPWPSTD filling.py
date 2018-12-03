import pandas as pd

##storecode = [8020,106,107,123,161,162,177,188,192,272,2006,2071,2097,
##             2119,2410,2416,2427,2436,2441,4132,8013,8032,8058,8077,
##             8086,8094,8107,8109,8118,8137,8306,"P006","P016","P025",
##             "P027","P033","P058","P059","P067","P073","P078","P090",
##            "P146", "P169", "P187", "P233","P234"]




storecode = [312,168,156,254,166,158,265,150,146,186]

df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/PwP new dumps/PE-FY-17-18-19 TILL AUG for python.xlsx', sheet_name = '1')
df.set_index('Partner', inplace = True)

def PE_STD(i,df):
    df1 = df.pivot_table(index = 'Partner', columns = ['SEASON','BRANDDD','Material group'], values = 'GSV', aggfunc = 'sum')
    df1.loc[i].fillna(0).to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/NewPE_collection/{}_newPE.xlsx'.format(i))

for i in storecode:
    PE_STD(i,df)
    print("Pass " + str(i))



##df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/PwP new dumps/LP-FY-17-18-19 TILL AUG for python.xlsx',sheet_name = '1')
##def LP(i,df):
##    
##    df1 = df.pivot_table(index = 'Partner', columns = ['YR','Brand-NITYA','Material group'], values= 'NSV', aggfunc = 'sum')
##    df2 = df.pivot_table(index = 'Partner', columns = ['SEASON','Brand-NITYA','Material group'],values = 'NSV', aggfunc = 'sum')
##    writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/LP_new_collection/{}_LP_new.xlsx'.format(i), engine = 'xlsxwriter')
##    try:
##        df2.loc[i].to_excel(writer, sheet_name = 'NewSeason')
##        df1.loc[i].to_excel(writer, sheet_name = 'FullYear')
##        writer.save()
##        writer.close()
##    except KeyError:
##        print('StoreCode '+ str(i) +' doesnt have LP')
##    print('End of LP')
##
##for i in storecode:
##    LP(i,df)


##df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/PwP new dumps/VH-FY-17-18-19 TILL AUG for python.xlsx',sheet_name = '1')
##
##def VH(i,df):
##    
##    df1 = df.pivot_table(index = 'Partner', columns = ['SEASON','Branddd','Material group'],values= 'NSV', aggfunc = 'sum')
##    df2 = df.pivot_table(index = 'Partner', columns = ['FY','Branddd','Material group'],values='NSV', aggfunc = 'sum')
##    writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/VH_new_collection/{}_VH_new.xlsx'.format(i), engine = 'xlsxwriter')
##
##    
##    try:
##        df1.loc[i].to_excel(writer, sheet_name = 'NewSeason')
##        df2.loc[i].to_excel(writer, sheet_name = 'FullYear')
##        writer.save()
##        writer.close()
##    except KeyError:
##        print('Store Code '+ str(i)+ ' doesnt have VH')
##    print('End of VH')
##
##
##for i in storecode:
##    VH(i,df)


##df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Data_DS/PwP new dumps/AS-FY-17-18-19 TILL AUG for python.xlsx',sheet_name = '1')
##def AS(i,df):
##    
##    df1 = df.pivot_table(index = 'Partner', columns = ['SEASON','Branddd','Material group'], values = 'NSV', aggfunc = 'sum')
##    df2 = df.pivot_table(index = 'Partner', columns = ['FY','Branddd','Material group'], values = 'NSV', aggfunc = 'sum')
##    writer = pd.ExcelWriter('C:/Users/rochisman.datta/Desktop/Python/Python code/AS_new_collection/{}_AS_new.xlsx'.format(i), engine = 'xlsxwriter')
##    try:
##        df1.loc[i].to_excel(writer, sheet_name = 'NewSeason')
##        df2.loc[i].to_excel(writer, sheet_name = 'FullYear')
##        writer.save()
##        writer.close()
##    except KeyError:
##        print('Store Code '+str(i)+' doesnt have AS')
##    print('End of AS')
##
##
##for i in storecode:
##    AS(i,df)
