import shutil,sys
def CopyHope(storecode):
    
    for i in storecode:
        shutil.copy("C:/Users/rochisman.datta/Desktop/Python/Excel Macros/template_HOPE.xlsx", "C:/Users/rochisman.datta/Desktop/Python/Excel Macros/{}_HOPE.xlsx".format(i))

storecode = [312,168,156,254,166,158,265,150,146,186]
CopyHope(storecode)

