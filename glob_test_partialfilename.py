from glob import glob


storecode = [2422,"P067","P059",107]

for i in storecode:
    filename = glob('C:/Users/rochisman.datta/Desktop/Python/Excel Macros/{}*.xlsx'.format(i))
    print(filename)
