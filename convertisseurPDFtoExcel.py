import tabula
import pandas as pd
import os
import PyPDF2
import re
import openpyxl

matches = ["150 microns","300 microns","300microns","400 microns","500 microns","600 microns","700 microns","700microns","3mm","5mm","10mm"]
pattern = "|".join(f"\\b{i}\\b" for i in matches)
matches2 = ["recto/verso","recto"]
pattern2 = "|".join(f"\\b{i}\\b" for i in matches2)

chemin = os.path.dirname('c:/Users/Utilisateur/App_Castorama_BC/')

liste_fichiers = os.listdir(chemin)



for fichier in liste_fichiers:
    #on ne traite que les pdf
    if fichier.split(".")[-1]=="pdf":
        nom = fichier[:-4]

 
        df_list = tabula.read_pdf(chemin+"/"+nom+".pdf", lattice = True, pages = 'all')

        #le [1::2] permet de ne pas prendre en compte les en-tête de chaque page du pdf de casto
        gros_df = pd.concat(df_list[1::2], ignore_index=True, sort=False)  

        gros_df['surface m2'] = (gros_df['Largeur'] / 1000)*(gros_df['Longueur'] / 1000) *gros_df['Qté']   

        gros_df["mots-clés"] = gros_df['Désignation'].str.findall(pattern,flags=re.IGNORECASE).str.join(",")
        gros_df["impression"] = gros_df['Désignation'].str.findall(pattern2,flags=re.IGNORECASE).str.join(",")

        gros_df['impression'] = gros_df['impression'].replace({'recto/verso' : 'Recto-Verso'}, regex=True)
        gros_df['impression'] = gros_df['impression'].replace({'recto' : 'Recto'}, regex=True)

        gros_df.loc[gros_df["mots-clés"] == '150 microns', "matiere"] = "Coala, Re-Stick 150, mat, blanc, adhésif repositionnable, Polymeric PVC, 300g/m2"

        gros_df.loc[gros_df["mots-clés"] == '200 microns', "matiere"] = "Coala Easy Stick Floor granulé, Floor graphics R10, jet d'encre solvant, UV, latex, grainé, blanc, colle acrylique trans.enlevab., Monomeric PVC, 274g/m2, 200 µm"

        gros_df.loc[gros_df["mots-clés"] == '300 microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 420g/m2, 300µm"

        gros_df.loc[gros_df["mots-clés"] == '300microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 420g/m2, 300µm"

        gros_df.loc[gros_df["mots-clés"] == '400 microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 563g/m2, 400µm"

        gros_df.loc[gros_df["mots-clés"] == '500 microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 682g/m2, 500µm"

        gros_df.loc[gros_df["mots-clés"] == '600 microns', "matiere"] = "Coala, Magnetic Printable PVC, mat, blanc, PVC, 2000g/m2, 600 µm"

        gros_df.loc[gros_df["mots-clés"] == '700microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 682g/m2, 500µm"
        gros_df.loc[gros_df["mots-clés"] == '700 microns', "matiere"] = "Pentaprint, PR-M 419/58, traité Corona 2 faces, mat, 05/9450, PVC, blanc, 682g/m2, 500µm"

        gros_df.loc[gros_df["mots-clés"] == '2mm', "matiere"] = "Forex Print mat, pellicule protect. 1 face, blanc, 2.000mm, 900g/m2"
        gros_df.loc[gros_df["mots-clés"] == '3mm', "matiere"] = "Akyprint, mat, polypropylène, blanc, 900g/m2"


        #on supprime ensuite les colonnes inutiles
        col_suppr = []
        for col in gros_df.columns :
            if 'Unnamed' in col :
                col_suppr.append(col)

        for c in col_suppr:
            gros_df.drop(c, inplace = True, axis = 1)

        gros_df.insert(0,'Type', "P")
        gros_df.loc[0,'Type'] = "E"
        del gros_df["mots-clés"]
        indexNames = gros_df[ gros_df['Désignation'] == 'Sous-total'].index
        indexNames2 = gros_df[ gros_df['Désignation'] == 'Transport (Franco)'].index
        indexNames3 = gros_df[ gros_df['Désignation'] == 'Total'].index
        gros_df.drop(indexNames, inplace=True)
        gros_df.drop(indexNames2, inplace=True)
        gros_df.drop(indexNames3, inplace=True)
        gros_df.to_excel("converti_excel/"+nom+".xlsx",index=False)


        #La partie suivante gère l'entête du fichier casto pour récupérer les infos souhaitées (n° de commande, date de livraison, magasin à livrer)

        #On ouvre à nouveau le pdf
    
        pdfFileObj = open(chemin +"/"+ nom+'.pdf', 'rb') 
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
        #on se contente de la première page
        
        pageObj = pdfReader.getPage(0) 
        Texte = pageObj.extractText()
        #on commence après le '.fr' qui est à la fin de la page mais apparait ici au début de la string (je ne sais pas pourquoi). Les infos avant ce '.fr' sont inutiles
        
        Texte = Texte.split('.fr')[-1]

        #On considère que toute majuscule précédée d'une minuscule (et non d'un espace) marque le début d'un nouveau "champ". 
        
        cases = re.split(r"([a-z][A-Z])", Texte)


        #Cette boucle permet la bonne séparation des cases et remet la minuscule à la fin de la précédente, la majuscule au début de la suivante
        
        i = 0
        while i < len(cases) :
            if len(cases[i])==2:
                cases[i+1] = cases[i][1] + cases[i+1]
                cases[i-1] = cases[i-1] + cases[i][0]

                cases.pop(i)
                i = i-1
            i+=1
        
        
        #On répète cette opération pour les majuscules précédées d'un chiffre. Attention, si jamais une référence du type "AZER65QDSRR" était dans le fichier, elle serait coupée. 
        #Cependant ça n'était pas le cas jusqu'à présent et à priori cela pourrait ne pas géner le fonctionnement de l'outil suivant la position de cette référence
        
        for i in range(len(cases)):
            cases[i] = re.split(r"([0-9][A-Z])", cases[i])



        cases = [item for sublist in cases for item in sublist]

        j = 0
        while j < len(cases) :
            if len(cases[j])==2:
                cases[j+1] = cases[j][1] + cases[j+1]
                cases[j-1] = cases[j-1] + cases[j][0]

                cases.pop(j)
                j = j-1
            j+=1

    
        pdfFileObj.close()

        #enfin on parcourt toutes les cases pour récupérer les infos que l'on cherche 
        for case in cases:
            if 'Numéro de commande' in case :
                num_com = case.split(':')[-1]
            if 'Magasin :' in case :
                magasin = case.split(':')[-2][:-5]
            if 'Date :' in case :
                date = case.split(':')[-1] 
 
           
        #dernière partie : on ouvre l'excel du devis et on y ajoute une feuille nommée 'Informations Client' contenant ces informations
        wb = openpyxl.load_workbook("converti_excel/"+nom+".xlsx")

        wb_sheet = wb.active

        wb_sheet['M1'] = 'Numéro de commande'
        wb_sheet['M2'] = num_com
        wb_sheet['N1'] = 'Magasin'
        wb_sheet['N2'] = magasin
        wb_sheet['O1'] = 'Date'
        wb_sheet['O2'] = date


        wb.save("converti_excel/"+nom+".xlsx")
        wb.close()
        print('Félicitations! Veuillez trouver le fichier converti dans le dossier "converti_excel"')
