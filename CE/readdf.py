import pandas as pd
import os
pd.options.display.float_format = '{:,.0F}'.format

class Ce:
    df = pd.DataFrame()
    dirname = 'data/kis/'
    
    def readdf(self):
        dirfiles = os.listdir(self.dirname)
        fullpaths = map(lambda name : os.path.join(self.dirname, name), dirfiles)
        for file in fullpaths:
            if file == 'data/kis/.DS_Store' : os.remove('data/kis/.DS_Store')
            self.df1 = pd.read_excel(file, header=2, sheet_name=None)
            self.df1 = pd.concat(self.df1, axis=0).reset_index(drop=True)
            self.df = pd.concat([self.df, self.df1], axis=0)


a = Ce()
a.readdf()
print(a.df.info())






