import tabula
import pandas as pd
import os
import PyPDF2
import re
import openpyxl
import glob


chemin = os.path.dirname('c:/Users/Utilisateur/App_Castorama_BC/')

liste_fichiers = os.listdir(chemin)


for fichier in liste_fichiers:
    #on ne traite que les pdf
    if fichier.split(".")[-1]=="pdf":
        nom = fichier[:-4]

 
        #df_list = tabula.read_pdf(chemin+"/"+nom+".pdf", lattice = True, pages = 'all',multiple_tables=True)

        #gros_df = pd.concat(df_list[1::2], ignore_index=True, sort=False)
        
        gros_df = pd.concat(tabula.read_pdf(chemin+"/"+nom+".pdf",pages='all',multiple_tables=True))

        #gros_df.replace('',float("NaN"), inplace=True)
        gros_df.replace('Sous-total','', inplace=True)
        gros_df.replace('Transport (Franco)','', inplace=True)
        gros_df.replace('Total','', inplace=True)
        del gros_df["Réf Frn"]
        del gros_df["Réf DTM"]
        del gros_df["Qté"]
        del gros_df["Longueur"]
      #  del gros_df["Largeur"]
        del gros_df["Total €"]

gros_df.to_excel("prix/prix.xlsx",index=False)        


gros_df  = gros_df.rename(columns={"Désignation": "Détail", "Prix\rUnitaire €": "Option1"})  

gros_df.to_excel("prix/prix.xlsx",index=False)   

df = pd.read_excel("prix/prix.xlsx")

matches = ["150 microns","300 microns","300microns","400 microns","500 microns","600 microns","700 microns","700microns","3mm","5mm","10mm"]
pattern = "|".join(f"\\b{i}\\b" for i in matches)

df["mots-clés"] = df['Détail'].str.findall(pattern,flags=re.IGNORECASE).str.join(",")

matches2 = ["recto/verso","recto"]
pattern2 = "|".join(f"\\b{i}\\b" for i in matches2)

df["impression"] = df['Détail'].str.findall(pattern2,flags=re.IGNORECASE).str.join(",")

df['impression'] = df['impression'].replace({'recto/verso' : 'Recto-Verso'}, regex=True)
df['impression'] = df['impression'].replace({'recto' : 'Recto'}, regex=True)
df['impression'] = df['impression'].replace({'' : 'Sans Impression'}, regex=True)


header = ['impression','mots-clés']
final = df.loc[:, header]

df.to_excel("prix/prix.xlsx",index=False)  

df2 = pd.read_excel("prix/prix.xlsx")

#df2['mots-clés'] = df2['mots-clés'].replace({'recto' : 'Recto'}, regex=True)
df2.groupby(['mots-clés']).groups

df2.loc[df2["mots-clés"] == '150 microns', "microns"] = "150µ"
df2.loc[df2["mots-clés"] == '150 microns', "matiere"] = "Coala, Re-Stick 150, mat, blanc, adhésif repositionnable, Polymeric PVC, 300g/m2"

df2.loc[df2["mots-clés"] == '200 microns', "microns"] = "200µ"
df2.loc[df2["mots-clés"] == '200 microns', "matiere"] = "Coala Easy Stick Floor granulé, Floor graphics R10, jet d'encre solvant, UV, latex, grainé, blanc, colle acrylique trans.enlevab., Monomeric PVC, 274g/m2, 200 µm"

df2.loc[df2["mots-clés"] == '300 microns', "microns"] = "300µ"
df2.loc[df2["mots-clés"] == '300 microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 420g/m2, 300µm"

df2.loc[df2["mots-clés"] == '300microns', "microns"] = "300µ"
df2.loc[df2["mots-clés"] == '300microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 420g/m2, 300µm"

df2.loc[df2["mots-clés"] == '400 microns', "microns"] = "400µ"
df2.loc[df2["mots-clés"] == '400 microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 563g/m2, 400µm"

df2.loc[df2["mots-clés"] == '500 microns', "microns"] = "500µ"
df2.loc[df2["mots-clés"] == '500 microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 682g/m2, 500µm"

df2.loc[df2["mots-clés"] == '600 microns', "microns"] = "600µ"
df2.loc[df2["mots-clés"] == '600 microns', "matiere"] = "Coala, Magnetic Printable PVC, mat, blanc, PVC, 2000g/m2, 600 µm"

df2.loc[df2["mots-clés"] == '700microns', "microns"] = "500µ"
df2.loc[df2["mots-clés"] == '700microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 682g/m2, 500µm"
df2.loc[df2["mots-clés"] == '700 microns', "microns"] = "500µ"
df2.loc[df2["mots-clés"] == '700 microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 682g/m2, 500µm"

df2.loc[df2["mots-clés"] == '2mm', "microns"] = "2mm"
df2.loc[df2["mots-clés"] == '2mm', "matiere"] = "Forex Print mat, pellicule protect. 1 face, blanc, 2.000mm, 900g/m2"
df2.loc[df2["mots-clés"] == '3mm', "microns"] = "3mm"
df2.loc[df2["mots-clés"] == '3mm', "matiere"] = "Akyprint, mat, polypropylène, blanc, 900g/m2"

#df2['mat']=df2['Détail'].str.extract(r'(^w{5})')
df2[df2['Détail'].str.match('^01_.*')== True]

print(df2)
df2.to_excel("prix/prix.xlsx",index=False)
