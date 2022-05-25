import pandas as pd
import numpy as np
import glob
import os
import re
import sys
import csv


path =r'converti_excel'
filenames = glob.glob(path + "/*.xlsx")

dfs = []

for df in dfs: 
    xl_file = pd.ExcelFile(filenames)
    df=xl_file.parse('Sheet1')
    dfs.concat(df, ignore_index=True)

print(df)    

