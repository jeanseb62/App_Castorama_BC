import pandas as pd
import os
import PyPDF2
import re
import openpyxl
import numpy as np
from openpyxl.styles import numbers


#gros_df = pd.read_excel('prix/réfs_prixNew.xlsx')            
#gros_df.replace('€',str("NaN"), inplace=True)
#gros_df.to_excel("prix/réfs_prixNew2.xlsx",index=False)


wb = openpyxl.load_workbook("prix/Price list finale - RFP ILV & Signaletique Riou_1er MAI_2022.xlsx")

#ws = wb.create_sheet("Feuil1")
wb_sheet = wb.active

#wb_sheet ['E1'] = wb_sheet['E1'].apply(lambda x: x.replace('€', '')
 #                               if isinstance(x, str) else x).astype(str)
wb_sheet['C1'] = 'Impression'
wb_sheet['D1'] = 'Détail'
wb_sheet['E1'] = 'Option1'
wb_sheet['F1'] = 'Option2'
wb_sheet['G1'] = 'Option3'

for r in range(12,999):
    wb_sheet[f'E{r}'].number_format ='##.00##'
    #wb_sheet[f'E{r}'] = str(wb_sheet[f'E{r}'].value)
for r in range(12,999):
    wb_sheet[f'F{r}'].number_format ='##.00##'
for r in range(12,999):
    wb_sheet[f'G{r}'].number_format ='##.00##'        

wb.save("prix/réfs_prixNew2.xlsx")
wb.close()


print('Félicitations!')

#prix = pd.read_excel('prix/réfs_prixNew2.xlsx')

#prix.to_excel('prix/réfs_prixNew2.xlsx',index=False)

#print('Encore félicitations!')