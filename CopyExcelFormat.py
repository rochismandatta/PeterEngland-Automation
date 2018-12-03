import shutil,sys
def CopyExcel(storecode):
    
    for i in storecode:
        shutil.copy("C:/Users/rochisman.datta/Desktop/Python/Excel Macros/template_formal.xlsx", "C:/Users/rochisman.datta/Desktop/Python/Excel Macros/{}_formal.xlsx".format(i))
        shutil.copy("C:/Users/rochisman.datta/Desktop/Python/Excel Macros/template_casual.xlsx", "C:/Users/rochisman.datta/Desktop/Python/Excel Macros/{}_casual.xlsx".format(i))



