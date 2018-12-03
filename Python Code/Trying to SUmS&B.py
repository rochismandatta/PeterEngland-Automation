import pandas as pd


def SB_CurrInv():
    df = pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrInv collection/CurrInv_2048.xlsx')
    df.fillna(0, inplace=True)
    ##int Jeans = df.iloc[0][3]
    ##int Shorts = df.iloc[0][5]
    ##print(Jeans+ Shorts)
    df['S&B']=((df.iloc[12]).get(2)+(df.iloc[19].get(2)))
    df.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrInv collection/CurrInv_2048.xlsx')
##    df3.to_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrInv collection/CurrInv_2048.xlsx', sheet_name='Sheet2')
    print(df)


SB_CurrInv()


def SOH_Calc(): #includes S&B and Normal 3 month average
    df=pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/SOH collection/SOH_2048.xlsx')
    df.fillna(0, inplace=True)
    df['ShirtAvg']=(df.iloc[28].get(3)+ df.iloc[39].get(3) + df.iloc[49].get(3))/3
    df['TrouserAvg']=(df.iloc[31].get(3)+ df.iloc[42].get(3) + df.iloc[52].get(3))/3
    df['T-ShirtAvg']=(df.iloc[32].get(3)+ df.iloc[43].get(3) + df.iloc[53].get(3))/3
    df['SNBAvg']=(df.iloc[24].get(3)+ df.iloc[36].get(3) + df.iloc[46].get(3)+ df.iloc[29].get(3) + df.iloc[40].get(3) + df.iloc[50].get(3))/6
    print(df)
##SOH_Calc()


def SB_CurrSales():
    df=pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/CurrSales collection/CurrSales_2048.xlsx')
    df.fillna(0,inplace=True)
    df['SBAvgSales']= df.iloc[10].get(2)+ df.iloc[14].get(2)
    print(df)

##SB_CurrSales()


def SB_Proj():
    df=pd.read_excel('C:/Users/rochisman.datta/Desktop/Python/Python code/ProjSl collection/Proj_2048.xlsx')
    df.fillna(0, inplace=True)
    df['SB_May']=df.iloc[25].get(3)+ df.iloc[28].get(3)
    df['SB_Jun']=df.iloc[32].get(3)+ df.iloc[35].get(3)
    df['SB_Jul']=df.iloc[40].get(3)+ df.iloc[44].get(3)
    print(df)
##SB_Proj()
