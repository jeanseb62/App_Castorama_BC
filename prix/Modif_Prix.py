import tabula
import pandas as pd
import os
import PyPDF2
import re
import openpyxl

chemin = os.path.dirname('c:/Users/Utilisateur/App_Castorama_BC/')

liste_fichiers = os.listdir(chemin)

#prix = pd.read_excel('prix/réfs_prixNew2.xlsx',sheet_name ='Price list finale réf_fév 2022')
#prix = pd.read_excel('prix/Classeur2.xlsx')

prix = pd.read_excel('prix/réfs_prix_Mai2022.xlsx',sheet_name ='Feuil1')
#prix = pd.read_excel('prix/réfs_prix.xlsx',sheet_name ='Feuil1')


for fichier in liste_fichiers:
    #on ne traite que les pdf
    if fichier.split(".")[-1]=="pdf":
        nom = fichier[:-4]

 
        df_list = tabula.read_pdf(chemin+"/"+nom+".pdf", lattice = True, pages = 'all')

        #le [1::2] permet de ne pas prendre en compte les en-tête de chaque page du pdf de casto
        gros_df = pd.concat(df_list[1::2], ignore_index=True, sort=False)
        
        matiere = []
        impression = []
        
        #ici on corrige le bug de casto qui conduit à ne pas avoir la matière
        #pour chaque ligne du tableau on compare le prix unitaire à celui de nos matières premières pour retrouver la matière correspondante
        
        for index, produit in gros_df.iterrows():
            price = produit["Prix\rUnitaire €"]

            mat = ''
            imp = ''
            for indexe, ligne in prix.iterrows():
                if price == ligne['Option1']:
                    mat = ligne['Détail'] 
                    #imp = 'Recto'
                    imp = ligne['Impression']
                elif price == ligne['Option2']:
                    mat = ligne['Détail'] 
                    #imp = 'Recto-Verso'
                    imp = ligne['Impression']
                elif price == ligne['Option3']:
                    mat = ligne['Détail'] 
                    #imp = 'Sans Impression'
                    imp = ligne['Impression']
            matiere.append(mat)
            impression.append(imp)

        #on ajoute les 2 colonnes matière et impression qui contiennent les infos manquantes
       
        gros_df['surface m2'] = (gros_df['Largeur'] / 1000)*(gros_df['Longueur'] / 1000) *gros_df['Qté']


        gros_df['matiere'] = matiere
        gros_df['impression'] = impression

        

        #on supprime ensuite les colonnes inutiles
        col_suppr = []
        for col in gros_df.columns :
            if 'Unnamed' in col :
                col_suppr.append(col)

        for c in col_suppr:
            gros_df.drop(c, inplace = True, axis = 1)

        #on sauvegarde le tableau dans le fichier excel du nom de notre choix
        
        
        gros_df.replace('',float("NaN"), inplace=True)
        gros_df.dropna(thresh=3,inplace=True)
        gros_df.reset_index(drop = True, inplace = True)
        gros_df.insert(0,'Type', "P")
        gros_df.loc[0,'Type'] = "E"
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


        #regex_Designation_no = re.compile(r"Désignation(INV-\d+)")
        #Designation = re.search(regex_Designation_no, Texte).group(1)
    
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

        #wb_sheet.append(["Désignation"])


        wb_sheet['M1'] = 'Numéro de commande'
        wb_sheet['M2'] = num_com
        wb_sheet['N1'] = 'Magasin'
        wb_sheet['N2'] = magasin
        wb_sheet['O1'] = 'Date'
        wb_sheet['O2'] = date
        
        #data=[num_com,magasin,Type,Designation,ref_Frn,ref_dtm,Qté,surface,matiere,impression,Longueur, Largeur,Total,Prix]
        #df = pd.DataFrame(data)
       # print(df)

        wb.save("converti_excel/"+nom+".xlsx")
        wb.close()
        print('Félicitations! Veuillez trouver le fichier converti dans le dossier "converti_excel"')
