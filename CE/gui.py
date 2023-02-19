import numpy as np
import pandas as pd
import os
from pandas.api.types import CategoricalDtype
import tkinter as tk
from tkinter import filedialog as fd

df = pd.DataFrame()
# dirname = 'data/kis/'
# dirfiles = os.listdir(dirname)
# fullpaths = map(lambda name: os.path.join(dirname, name), dirfiles)
# pd.options.display.float_format = '{:,.0F}'.format

# for file in fullpaths:
# 	if file == 'data/kis/.DS_Store': os.remove('data/kis/.DS_Store')
# 	df1 = pd.read_excel(file, header=2, sheet_name=None)
# 	df1 = pd.concat(df1, axis=0).reset_index(drop=True)
# 	df = pd.concat([df, df1], axis=0)

def callback():
    name = fd.askopenfilename()
    print(name)
    df = pd.read_excel(os.path.abspath(name), header=2,engine="xlrd")
    print(df.info())


errmsg = 'Error!'
tk.Button(text='Click to Open File', command=callback).pack(fill=tk.X)
tk.mainloop()






