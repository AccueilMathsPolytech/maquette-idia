#import os
import pandas as pd
#from mdutils.mdutils import MdUtils
#from mdutils import Html
#from texttable import Texttable
#import latextable

from pylatex import Document, Section, Subsection, Tabular, Math, TikZ, Axis, Plot, Figure, Matrix, Alignat, LongTable
from pylatex.utils import italic


geometry_options = {"tmargin": "1cm", "lmargin": "10cm"}
doc = Document(geometry_options=geometry_options)

#
#  Lecture des tables de données depuis la source XLS
#

data_source_file = "schema-donnees-IDIA.xlsx"
dfDescriptionUE = pd.read_excel(
    data_source_file, sheet_name="description UE")
dfCompositionSemestre = pd.read_excel(
    data_source_file, sheet_name="composition semestre")
dfCompositionUE = pd.read_excel(
    data_source_file, sheet_name="composition UE")
dfDescriptionMatiere = pd.read_excel(
    data_source_file, sheet_name="Description matière")

#
# Creation de la table jointe Semestre-UE-Matiere
#

dfUEmatieres = pd.merge(dfCompositionSemestre, dfCompositionUE)
dfDescriptionUE.set_index("NomUE", inplace=True)

print(dfUEmatieres)

# Pour le tableau récapitulatif affiché en debut de semestre

docSyntheseSemestre = Section("Semestre")
with docSyntheseSemestre.create(LongTable("c c c c")) as syntheseSemestre:

    # Extraire la liste des noms de semestres
    nomsSemestres = dfCompositionSemestre['Semestre'].drop_duplicates()
    print(nomsSemestres)

    # Boucle sur l'ensemble des semestres

    UE_number = 0

    for semestre in nomsSemestres:
        dfUEmatieres_dans_semestre = dfUEmatieres[dfUEmatieres['Semestre'] == semestre]

        print(dfUEmatieres_dans_semestre)

        nomsUE = dfUEmatieres_dans_semestre['NomUE'].drop_duplicates()

        print(nomsUE)

        #dfDescriptionUE.set_index("NomUE", inplace=True)

        TempsTotalSemestreMaquette = 0
        TempsTotalSemestrePerso = 0
        TempsTotalSemestre = 0
        ECTSTotalSemestre = 0

        #
        # Boucle sur l'ensemble des UE dans le semestre courant
        #

        # sert à générer les noms de fichiers (un fichier par UE)

        for nom_UE in nomsUE:

            # sert à générer les noms de fichiers (un fichier par UE)
            UE_number = UE_number+1

            descriptionUE = dfDescriptionUE.loc[nom_UE, "Description UE"]
            ECTSUE = dfDescriptionUE.loc[nom_UE, "ECTS"]

            # Ecrire le descriptif d'UE dans le fichier d'UE

            docUE = Section("UE " + nom_UE)
            docUE.create('')
            docUE.append(descriptionUE)

            #
            # Ajouter l'UE dans le tableau récapitulatif qui ira en début de semestre
            #
            syntheseSemestre.add_row(("", "", "", ""))
            syntheseSemestre.add_row(("UE", "", "", "Crédits ECTS"))
            syntheseSemestre.add_row((nom_UE, "", "", ECTSUE))
            syntheseSemestre.add_row(("", "", "", ""))
            syntheseSemestre.add_row(
                ("Matière", "Présentiel", "Personnel", "Total"))

            matieresDansUE = dfUEmatieres_dans_semestre[dfUEmatieres["NomUE"] == nom_UE]
            matieresDansUE.reset_index(drop=True, inplace=True)
            matieresDansUE = pd.merge(matieresDansUE, dfDescriptionMatiere)

            TempsTotalUEMaquette = 0
            TempsTotalUEPerso = 0

            #
            # Boucle sur l'ensemble des matières dans l'UE courante
            #
            for mat in range(len(matieresDansUE)):

                nom_mat = matieresDansUE.loc[mat, "Nom_Matiere"]
                HeuresCM = matieresDansUE.loc[mat, "CM"]
                HeuresTD = matieresDansUE.loc[mat, "TD"]
                HeuresProjet = matieresDansUE.loc[mat, "Projet"]
                HeuresTP = matieresDansUE.loc[mat, "TP"]
                HeuresEval = matieresDansUE.loc[mat, "Eval"]
                HeuresMaquette = HeuresCM+HeuresTD+HeuresTP+HeuresProjet+HeuresEval
                TempsPersonnel = matieresDansUE.loc[mat, "TravailPersonnel"]
                TempsTotalMatiere = HeuresMaquette+TempsPersonnel
                TempsTotalUEMaquette = TempsTotalUEMaquette+HeuresMaquette
                TempsTotalUEPerso = TempsTotalUEPerso+TempsPersonnel
                TempsTotalUE = TempsTotalUEMaquette+TempsTotalUEPerso

                #
                # ajouter les informations descriptives de matière dans le document d'UE
                #

                docUE.append(Subsection(nom_mat))
                docUE.append(matieresDansUE.loc[mat, "Description_Matiere"])

                with docUE.create(LongTable("c c c c c c c")) as table:
                    table.add_hline()
                    table.add_row(("Cours", "TD", "TP", "Projet",
                                   "Eval", "Personnel", "Total"))
                    table.add_row((str(HeuresCM)+" h", str(HeuresTD)+" h", str(HeuresTP)+" h", str(HeuresProjet) +
                                   " h", str(HeuresEval)+" h", str(TempsPersonnel)+" h", str(HeuresMaquette+TempsPersonnel)+" h"))
                    table.add_hline()

                #
                # ajouter les informations d'entête sur l'UE dans le tableau de récap
                #

                syntheseSemestre.add_row((nom_mat, str(HeuresCM+HeuresTD+HeuresTP+HeuresProjet+HeuresEval) +
                                          " h", str(TempsPersonnel)+" h", str(HeuresMaquette+TempsPersonnel)+" h"))
            docUE.append(Subsection("Total UE :"))

            #
            # ajouter le récap des infos d'UE dans le document d'UE
            #

            with docUE.create(LongTable("c c c c")) as table:
                table.add_hline()
                table.add_row(("Travail maquette", "Travail personel",
                               "Travail total", "Crédits ECTS"))
                table.add_row((str(TempsTotalUEMaquette)+" h", str(TempsTotalUEPerso) +
                               " h", str(TempsTotalUEMaquette+TempsTotalUEPerso)+" h", str(ECTSUE)))
                table.add_hline()

            # Creation du fichier latex contenant une description d'UE
            docUE.generate_tex("generated/"+str(UE_number))

            syntheseSemestre.add_hline()

            syntheseSemestre.add_row(("", str(TempsTotalUEMaquette)+" h", str(
                TempsTotalUEPerso)+" h", str(TempsTotalUEMaquette+TempsTotalUEPerso)+" h"))
            syntheseSemestre.add_hline()

            # Creation du fichier latex contenant une description d'UE

            syntheseSemestre.generate_tex("generated/"+str("semestre"))

            TempsTotalSemestreMaquette = TempsTotalSemestreMaquette+TempsTotalUEMaquette
            TempsTotalSemestrePerso = TempsTotalSemestrePerso+TempsTotalUEPerso
            TempsTotalSemestre = TempsTotalSemestre + TempsTotalUE

            ECTSTotalSemestre = ECTSTotalSemestre+ECTSUE
