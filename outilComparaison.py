import pandas as pd
import numpy as np
import glob
import os

fileTest = r"comparaisons/Fichier_modifie.xlsx"

file_list = glob.glob("converti_excel/*.xlsx")
files = []
for filename in file_list:
    df = pd.read_excel(filename)
    files.append(df)  
frame = pd.concat(files, axis=0, ignore_index=True)
frame.to_excel("comparaisons/Fichier_modifie.xlsx",index=False)

df2 = pd.read_excel("comparaisons/Fichier_modifie.xlsx")
df2 = df2.rename(columns={"matiere": "Description"})          
df2.to_excel("comparaisons/Fichier_modifie.xlsx",index=False)  

df1 = pd.read_excel('base/Fichier-Base.xlsx',sheet_name ='EPace Data')
df2 = pd.read_excel('comparaisons/Fichier_modifie.xlsx')

data = pd.concat([df1, df2], sort=False, ignore_index=True)

data = pd.merge(df1, df2, on=['Description'], how='outer')

header = ['ID Type de produit','Description','Désignation']
data = data.loc[:, header]

data['Désignation'].replace('', np.nan, inplace=True)
data.dropna(subset=['Désignation'], inplace=True)

data['ID Type de produit'].replace('', np.nan, inplace=True)
data.dropna(subset=['ID Type de produit'], inplace=True)

data.to_excel("comparaisons/Fichier_final.xlsx",index=False)


try:
    os.remove(fileTest)
except OSError as e:
    print(e)
else:
    print("File is deleted successfully")
print('Félicitations! Veuillez trouver le fichier "Fichier_final.xlsx" dans le dossier comparaisons.')
