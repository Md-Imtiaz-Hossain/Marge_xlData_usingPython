
import pandas as pd
import os
files = os.listdir('pandas/')
datasets = []
for f in files:
    datasets.append(pd.read_csv('pandas/' + f, index_col=None))

main = pd.DataFrame()
for i in datasets:
    main = main.append(i)

main.to_excel('pandas/Final.xlsx', index= False)
print(main)