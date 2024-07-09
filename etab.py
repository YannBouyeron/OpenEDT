import pandas as pd
import os
import random
import numpy as np
import json
import time
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows



pd.set_option("max_rows", None)




class Etab:


    def __init__(self, path_name, **dico):


        self.name = path_name

        self.dico = dico



        self.six = pd.read_excel("base/sixieme.xlsx")
        self.cin = pd.read_excel("base/cinquieme.xlsx")
        self.qua = pd.read_excel("base/quatrieme.xlsx")
        self.tro = pd.read_excel("base/troisieme.xlsx")
        self.sec = pd.read_excel("base/seconde.xlsx")
        self.pre = pd.read_excel("base/premiere.xlsx")
        self.ter = pd.read_excel("base/terminale.xlsx")

        self.ph = [111, 112, 113, 114, 121, 122, 123, 124, 211, 212, 213, 214, 221, 222, 223, 224, 311, 312, 313, 314, 411, 412, 413, 414, 421, 422, 423, 424, 511, 512, 513, 514, 521, 522, 523, 524, 611, 612, 613, 614]


        self.lc = [(self.six,"sixieme"),(self.cin,"cinquieme"),(self.qua,"quatrieme"),(self.tro,"troisieme"),(self.sec,"seconde"),(self.pre,"premiere"),(self.ter,"terminale")]



        if os.path.isfile(self.name):

            try: 
                
                df = pd.read_excel(self.name)

            except:

                print("Impossible d'ouvrir le fichier")

            else:

                # on remplace tous les NaN des la colonne ph par "[0,0,0]"
        
                df["ph"]= df["ph"].fillna("[0,0,0]")

                # on remplace tous les NaN de la colonne prof par " "

                df["prof"] = df["prof"].fillna(" ")

                # on remplace tous les NaN de la colonne matiere par ""

                df["matiere"] = df["matiere"].fillna("")

                # on remplace tous les NaN de la colonne semaine par ""

                df["semaine"] = df["semaine"].fillna("")

                # on remplace tous les NaN de la colonne salle par ""

                df["salle"] = df["salle"].fillna("")

        else:

            df = self.creat(**self.dico)


        self.df = df


    def service(self, prof):

        """retourne le service hebdo en heure d'un prof , et son DataFrame"""

        p = self.df.loc[self.df.prof == prof]

        sut = p.ut.sum()

        for i in p.itertuples():

            if i.semaine in ["SP","SI"]:

                sut -= i.ut/2

            if i.regroup in p.index and i.semaine not in ["SP","SI"]:

                sut -= i.ut

        return sut, p



                




    def get_reg(self):
    
        """ retourne un df des seances en regroupement (seances maitresses)"""

        r = self.df[(self.df.regroup > 0)]
        
        r = r.regroup.tolist()
        
        r  = list(set(r))
        
        #r = self.df.loc[r] c'est déprécié !

        r = self.df.reindex(r)

        return r




    def get_reg_classe(self, classe):

        r = self.get_reg()

        sm_in_classe = r.loc[r.classe == classe]

        sr_to_sm_in_classe = self.df.loc[self.df.regroup.isin(sm_in_classe.id.tolist())]

        sm_out_classe = self.df.loc[(self.df.classe != classe) & (self.df.id.isin(self.df.loc[self.df.classe == classe, "regroup"].tolist()))]

        sr_to_sm_out_classe = self.df.loc[self.df.regroup.isin(sm_out_classe.id.tolist())]



        txt = ""

        for k in sm_in_classe.itertuples():
        
            txt += "id: " + str(k.id) + " classe: " + str(k.classe) + " matière: " + str(k.matiere) + " " + str(k.semaine) + " prof: " + str(k.prof) + "\n"
        
            x = self.df[(self.df.regroup == k.id)]
        
            for i in x.itertuples():

                txt += "id: " + str(i.id) + " classe: " + str(i.classe) + " matière: " + str(i.matiere) + " " + str(i.semaine) + " prof: " + str(i.prof)  + "\n"
        
            txt += "\n \n"

        for k in sm_out_classe.itertuples():
        
            txt += "id: " + str(k.id) + " classe: " + str(k.classe) + " matière: " + str(k.matiere) + " " + str(k.semaine) + " prof: " + str(k.prof) + "\n"
        
            x = self.df[(self.df.regroup == k.id)]
        
            for i in x.itertuples():

                txt += "id: " + str(i.id) + " classe: " + str(i.classe) + " matière: " + str(i.matiere) + " " + str(i.semaine) + " prof: " + str(i.prof)  + "\n"
        
            txt += "\n \n"


        return txt


    def analyse(self):

        plages_horraires = [111, 112, 113, 114, 121, 122, 123, 124, 211, 212, 213, 214, 221, 222, 223, 224, 311, 312, 313, 314, 411, 412, 413, 414, 421, 422, 423, 424, 511, 512, 513, 514, 521, 522, 523, 524, 611, 612, 613, 614]


        lc = list(set(self.df.classe.tolist()))

        lp = list(set(self.df.prof.tolist()))

        if " " in lp:
            lp.remove(" ")



        r = "Analyse de vos classes: \n \n"

        r += "Vous avez {0} classes et {1} professeurs \n \n".format(len(lc), len(lp))


        for i in lc:

            r += "\n\n"

            r += "Classe {0}: \n\n".format(i)

            j = self.df[(self.df.classe == i)]

            h_requise = j.loc[j.matiere != "", "ut"].sum() - j.loc[j.regroup > 0, "ut"].sum()

            h_dispo = len(plages_horraires) - j.loc[j.prof == " ", "ut"].sum()

       
            if h_requise > h_dispo:

                r += "Classe {0}: [Impossible] Cette classe recquiert un total de {1} heures pour un total de {2} heures disponibles.\n \n".format(i, h_requise, h_dispo)

            else:
        
                r += "Classe {0}: Cette classe recquiert un total de {1} heures pour un total de {2} heures disponibles.\n \n".format(i, h_requise, h_dispo)

            lm = list(set(j.matiere.values.tolist()))

            for m in lm:

                if m != "":

                    r += "Matière: {0}   professeur: {1} \n".format(m, j.loc[j.matiere==m, "prof"].tolist()[0])

            r += "\n\n"
            r += "Analyse des regroupements: \n \n" 
            r += self.get_reg_classe(i)

        html = "<p>" + r.replace("\n", "</br>") + "</p>"


        return r, html






    def show_reg(self):

        """
        Affiche tous les regrouppements.
        """

        r = self.get_reg()

        for k in r.itertuples():

            
            print("id: " + str(k.id) + " classe: " + k.classe + " matière: " + k.matiere + " prof: " + k.prof)
        
            x = self.df[(self.df.regroup == k.id)]
        
            for i in x.itertuples():

                print("id: " + str(i.id) + " classe: " + i.classe + " matière: " + i.matiere + " prof: " + i.prof)

            print("\n")
        

    def delete_row(self, i):

        dff = self.df.drop(i)

        dff.index = range(len(dff))
        dff.id = range(len(dff))

        dff.loc[dff.regroup > i, "regroup"] -= 1

        self.df = dff


    def insert_row(self, r, i):

        """
        Insert une ligne (au format dataframe) r à l' index i, puis decale les regroupements supérieur à l'index  i, puis ré indexe.
        """

        a = self.df.iloc[:i]
        b = self.df.iloc[i:]

        dff = pd.concat([a,r], ignore_index=True, sort=False)
        dff = pd.concat([dff, b], ignore_index=True, sort=False)

        dff.index = range(len(dff))
        dff.id = range(len(dff))

        dff.loc[dff.regroup > i, "regroup"] += 1

        self.df = dff


    def duplic_row(self, i):

        """
        Copie une ligne à l'identique et l'insert sous la ligne originale, puis decale les regroupements et ré indexe.
        """

        r = self.df.loc[i]
        r = pd.DataFrame([r])

        self.insert_row(r, i)


    def divise_block2(self, i):

        """
        Divise une seance de 2 ut en 2 seances de 1 ut, insere la nouvelle ligne, re indexe
    
        i: index de la seance à diviser
        """

        if self.df.at[i, "ut"] == 2:

            r = self.df.loc[i]
            r = pd.DataFrame([r])

            self.insert_row(r, i)

            self.df.at[i, "ut"] = 1
            self.df.at[i+1, "ut"] = 1

    

    def add_prof(self, prof, matiere, *classes):

        """
        Ajoute un prof pour la matiere indiquée en argument, et pour les classes indiquées.
        """

        self.df.loc[(self.df.matiere.str.contains(matiere)) & (self.df.classe.isin(classes)), "prof"] = prof


    
    def add_random_prof(self, matiere, *name):

        """
        Ajoute des profs au hasard...
        """

        lc = list(set(self.df.classe.tolist())) 

        for n in name:

            hprof = self.df.loc[self.df.prof == n, "ut"].sum()

            while hprof < 17 and len(lc) > 0:

                c = random.choice(lc)

                self.df.loc[(self.df.classe == c) & (self.df.matiere.str.contains(matiere)), "prof"] = n

                lc.remove(c)

                hprof = self.df.loc[self.df.prof == n, "ut"].sum()

    


    def add_random_all_prof(self):

        lm = ["Fr", "Math", "SVT", "SES", "SNT", "PC", "HG", "Phylo", "LV1", "LV2-1", "LV2-2", "Huma", "Langues", "Art", "Musique", "EPS", "Techno"]

        c = 0

        for m in lm:

            self.random_prof(m, *["Prof_" + str(i) for i in range(c, c+20)])
            c += 20



    def add_regroup(self, ref, *reg):

        for r in reg:

            self.df.at[r, "regroup"] = ref



    def add_contrainte(self, name, *contr):

        """        
        Ajoutes les contriantes 1 par 1 (pas en block) pour un prof ou pour une classe de nom name
        """



        if name in self.df.classe.values:

            for c in contr:

                c = "[" + str(c) + "]"

                r = pd.DataFrame({"id":(len(self.df)+1,), "classe":(name,), "matiere": (np.nan,), "semaine":(np.nan,), "prof":("",), "ut":(1,), "regroup":(np.nan,), "salle":(np.nan,), "ph":(c,)})  

                self.df = pd.concat([self.df, r], ignore_index=True, sort=False)



        if name in self.df.prof.values:

            for c in contr:

                c = "[" + str(c) + "]"

                r = pd.DataFrame({"id":(len(self.df)+1,), "classe":(np.nan,), "matiere": (np.nan,), "semaine":(np.nan,), "prof":(name,), "ut":(1,), "regroup":(np.nan,), "salle":(np.nan,), "ph":(c,)})

                self.df = pd.concat([self.df, r], ignore_index=True, sort=False)



        self.df.index = range(len(self.df))
        self.df.id = range(len(self.df))



    def save(self):

        """
        Sauvegarde le dataframe etab au format xlsx en remplacant les regroup par des formules relatives pour excel.
        
        Retourne le df modifié avec les formules
        """

        df = self.df.copy()

        df.regroup = df.regroup.astype(object)
            
        # on remplace les regroup préconfigurés par des formules excel de lien entre cellule

        for z in range(len(df)):

            try:

                df.at[z, "regroup"] = "=A" + str(int(df.at[z, "regroup"]) + 2) # decalage de +2 entre les index et id de pd et les index excel

            except:

                pass
        
        

        df.to_excel(self.name + ".xlsx", index=False)

        return df





    def creat(self, **dico):

        """ 
        Assemble des dataframes des différents modèles de classe passés en argument.
        
        dico est un dictionnaire du type: {"sixième":3, "quatrième":3, "samedi":True, "mercrdedi":True, "m4":True, "a3":True, "a4":True}
        """ 



        # dictionnaire donnant acces aux df modèles

        d = {"sixième":self.six, "cinquième":self.cin, "quatrième":self.qua, "troisième":self.tro, "seconde":self.sec, "première":self.pre, "terminale":self.ter}




        # liste des classes

        lc = []


        # assemblage des df des classes selectionnées 

        for k in dico.keys() - ["samedi","mercredi","m1","a3","a4"]:

            for i in range(dico[k]):
            
            
                df2 = d[k].copy()
                df2.classe = k.capitalize() + " " + str(i+1)
            
                lc.append(k.capitalize() + " " + str(i+1)) 

                try:
                    df2.regroup += len(df)
                    df = pd.concat([df, df2], sort=False)
                    df.index = range(len(df))
                    df.id = range(len(df))

                except:

                    df = df2.copy()
   


        # on block les heures libres eventuelles du samedi, mercredi, m1, a3, a4 généralisé à toutes les classes

        for h in ["samedi","mercredi","m1","a3","a4"]:

            try:

                x = dico[h]
   
            except:

                pass

            else:

                if h == "samedi":

                    r = pd.DataFrame({"id":[np.nan]*len(lc), "classe":lc, "matiere": [np.nan]*len(lc), "prof": [""]*len(lc), "ut":[4]*len(lc), "regroup": [np.nan]*len(lc), "salle": [np.nan]*len(lc), "ph":["[611,612,613,614]"]*len(lc)})
                    df = pd.concat([df,r],ignore_index=True, sort=False)


                elif h == "mercredi":
                   
                    r = pd.DataFrame({"id":[np.nan]*len(lc), "classe":lc, "matiere": [np.nan]*len(lc), "prof": [""]*len(lc), "ut":[4]*len(lc), "regroup": [np.nan]*len(lc), "salle": [np.nan]*len(lc), "ph":["[311,312,313,314]"]*len(lc)})
                    df = pd.concat([df,r],ignore_index=True, sort=False)


                elif h == "m1":

                    for t in ["[111]", "[211]", "[311]", "[411]", "[511]", "[611]"]:

                        r = pd.DataFrame({"id":[np.nan]*len(lc), "classe":lc, "matiere": [np.nan]*len(lc), "prof": [""]*len(lc), "ut":[1]*len(lc), "regroup": [np.nan]*len(lc), "salle": [np.nan]*len(lc), "ph":[t]*len(lc)})
                        df = pd.concat([df,r],ignore_index=True, sort=False)


                elif h == "a3":

                    for t in ["[123]", "[223]", "[423]", "[523]", "[124]", "[224]", "[424]", "[524]"]:
                        r = pd.DataFrame({"id":[np.nan]*len(lc), "classe":lc, "matiere": [np.nan]*len(lc), "prof": [""]*len(lc), "ut":[1]*len(lc), "regroup": [np.nan]*len(lc), "salle": [np.nan]*len(lc), "ph":[t]*len(lc)})
                        df = pd.concat([df,r],ignore_index=True, sort=False)


                elif h == "a4":
                    for t in ["[124]", "[224]", "[424]", "[524]"]:
                        r = pd.DataFrame({"id":[np.nan]*len(lc), "classe":lc, "matiere": [np.nan]*len(lc), "prof": [""]*len(lc), "ut":[1]*len(lc), "regroup": [np.nan]*len(lc), "salle": [np.nan]*len(lc), "ph":[t]*len(lc)})
                        df = pd.concat([df,r],ignore_index=True, sort=False)


        # on ré index

        df.index = range(len(df))
        df.id = range(len(df))


        # on remplace tous les NaN des la colonne ph par "[0,0,0]"
        
        df["ph"]= df["ph"].fillna("[0,0,0]")

        # on remplace tous les NaN de la colonne prof par " "

        df["prof"] = df["prof"].fillna(" ")

        # on remplace tous les NaN de la colonne matiere par ""

        df["matiere"] = df["matiere"].fillna("")

        # on remplace tous les NaN de la colonne semaine par ""

        df["semaine"] = df["semaine"].fillna("")

        # on remplace tous les NaN de la colonne salle par ""

        df["salle"] = df["salle"].fillna("")
        

        return df








