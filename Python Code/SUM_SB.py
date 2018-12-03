import pandas as pd


def SB_CurrInv(i):
    df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrInv collection/CurrInv_{}.xlsx'.format(i))
    df.fillna(0, inplace=True)
    ##int Jeans = df.iloc[0][3]
    ##int Shorts = df.iloc[0][5]
    ##print(Jeans+ Shorts)
    df['S&B']=((df.iloc[12]).get(2)+(df.iloc[19].get(2)))
    df.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrInv collection/CurrInv_{}.xlsx'.format(i))
##    df3.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrInv collection/CurrInv_2048.xlsx', sheet_name='Sheet2')
    print(df)


#SB_CurrInv()


def SOH_Calc(i): #includes S&B and Normal 3 month average
    df=pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/SOH collection/SOH_{}.xlsx'.format(i))
    df.fillna(0, inplace=True)
    df['ShirtAvgFormal']=(df.iloc[28].get(3)+ df.iloc[39].get(3) + df.iloc[49].get(3))/3
    df['TrouserAvgFormal']=(df.iloc[31].get(3)+ df.iloc[42].get(3) + df.iloc[52].get(3))/3
    df['T-ShirtAvgFormal']=(df.iloc[32].get(3)+ df.iloc[43].get(3) + df.iloc[53].get(3))/3
    df['SNBAvgFormal']=(df.iloc[24].get(3)+ df.iloc[36].get(3) + df.iloc[46].get(3)+ df.iloc[29].get(3) + df.iloc[40].get(3) + df.iloc[50].get(3))/6

    df['ShirtAvgCasual']=(df.iloc[2].get(3)+ df.iloc[10].get(3) + df.iloc[17].get(3))/3
    df['TrouserAvgCasual']=(df.iloc[6].get(3)+ df.iloc[13].get(3) + df.iloc[20].get(3))/3
    df['T-ShirtAvgCasual']=(df.iloc[7].get(3)+ df.iloc[14].get(3) + df.iloc[21].get(3))/3
    df['JeansAvgCasual']=(df.iloc[1].get(3)+ df.iloc[9].get(3) + df.iloc[16].get(3))/3
    df.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/SOH collection/SOH_{}.xlsx'.format(i))
    print(df)
##SOH_Calc()


def SB_CurrSales(i):
    df=pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrSales collection/CurrSales_{}.xlsx'.format(i))
    df.fillna(0,inplace=True)
    df['SBAvgSales']= df.iloc[10].get(2)+ df.iloc[14].get(2)
    df.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrSales collection/CurrSales_{}.xlsx'.format(i))
    print(df)

##SB_CurrSales()


def SB_Proj(i):
    df=pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/ProjSl collection/Proj_{}.xlsx'.format(i))
    df.fillna(0, inplace=True)
    df['SB_May']=df.iloc[25].get(3)+ df.iloc[28].get(3)
    df['SB_Jun']=df.iloc[32].get(3)+ df.iloc[35].get(3)
    df['SB_Jul']=df.iloc[40].get(3)+ df.iloc[44].get(3)
    df.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/ProjSl collection/Proj_{}.xlsx'.format(i))
    print(df)
##SB_Proj()
