#!/usr/bin/python
# -*- coding: utf-8 -*-

# from qgis.core import Qgis, QgsProject, QgsFeatureRequest, QgsExpression, NULL
import os
# import openpyxl
import numpy as np
import pandas as pd
# from tabulate import tabulate

left = pd.DataFrame(
    {
        "key1": ["K0", "K0", "K1", "K2"],
        "key2": ["K0", "K1", "K0", "K1"],
        "A": ["A0", "A1", "A2", "A3"],
        "B": ["B0", "B1", "B2", "B3"],
    }
)
# print(f"left {left}")
right = pd.DataFrame(
    {
        "key1": ["K0", "K1", "K1", "K2"],
        "key2": ["K0", "K0", "K0", "K0"],
        "C": ["C0", "C1", "C2", "C3"],
        "D": ["D0", "D1", "D2", "D3"],
    })
# print(f"right {right}")

# join_func = {
#     "inner": libjoin.inner_join,
#     "left": libjoin.left_outer_join,
#     "right": _right_outer_join,
#     "outer": libjoin.full_outer_join,
# }

result = pd.merge(left, right, how="left", on=["key1", "key2"])
# print(result)

# ['372301', np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, '372300',np.NaN, np.NaN, np.NaN, np.NaN, '372299', 'BT-454', 'BT-5454', 'NC']

df = pd.DataFrame({'Time': ['372301', np.nan, np.nan, np.nan, np.nan, np.nan, 372300, np.nan, np.nan, np.nan, np.nan, 372299, 'BT-454', 'BT-5454', 'NC'], "Autre":"Rien"})
# df = pd.DataFrame({'Time': [66, '91', 1.23, 'a', 0.0, '', np.nan, 77], "Autre":"Rien"})
print(f"df :\n{df}")

new_df = df[(df.loc[:, 'Time'].map(type) == int)]
# non_int_df = df[~df['Time'].map(pd.api.types.is_integer)]
print("Tout\n", new_df)


def LectureFichiersExcelsCap_ft(df):  # , repertoire):
    """Fonction pour parcourir les fichiers Excel pour renseigner la référence des appuis."""
    dicoPoteauFt_SousTraitant = {}

    # df = pd.DataFrame(columns=["N° appui", "Nature des travaux", "Fichier"])
    # print(f"df \n{df}")
    repertoire = f"C:/Users/asoumare/Documents/Teletravail/Developpement/CAP_COMAC/UNique/"
    # repertoire = f"C:/Users/asoumare/Documents/Teletravail/Developpement/CAP_COMAC/test/"
    for subdir, dirs, files in os.walk(repertoire):
        dossier_parent = ""
        for name in files:
            # print(f"name {name}")
            # print(f"name 00 : {name}")

            if name.endswith('.xlsx') and str("~$") not in str(name):
                cheminComplet = subdir + name
                # df1 = pd.datetime
                # print(f"name : {name}")
                # try:
                with pd.ExcelFile(cheminComplet) as xls:
                    # On prend les données à partir de la ligne 7 (la colonne).
                    df1 = pd.read_excel(xls, "Export 1", header=7, index_col=None)  # , index_col=0,  skiprows=[15], na_values=["NA"],)

                # pd.read_excel(cheminComplet, sheet_name="Export 1")
                # print(tabulate(df1, headers='keys', tablefmt='psql'))
                # print(df1.T)
                # print(df1)
                # print("ligne : \n", df1.iloc[:, 0:32:31].dropna(thresh=1))  #, df1.iloc[:, 31])

                # Creation d'un Dataframe à partir des données collectées dans Excel
                # En supprimant les lignes qui ont une plus de un np.NaN
                # Les colonnes qui nous intéresse sont dans la
                print(f"test : \n{df1.iloc[:, 0].values}")
                resultat = df1.iloc[:, 0:32:31].dropna(thresh=1)
                print("resultat 0 :\n", resultat)

                # df1.iloc[0, :].map(type) != str

                # print(f"AUTRE  :\n{resultat.loc[:, 'N° appui'].map(type) == int}")
                #
                # resultat2 = resultat[(resultat.loc[:, 'N° appui'].map(type) != str)]
                # # resultat = resultat[~resultat['N° appui'].map(pd.api.types.is_integer)]
                # print("resultat2  :\n", resultat2)

                # print("ok", df1.loc[~df['N° appui'].str.isdigit(), 'N° appui'])
                # df1.loc[~df['N° appui'].astype(str).str.isdigit(), 'N° appui'].tolist()
                # df1.loc[~df['N° appui'].astype(str).str.isdigit(), 'N° appui'].tolist()
                # (df1['Time'].map(type) != int)
                # On ajoute le nom du fichier dans une nouvelle colonne
                resultat.insert(2, "Fichier", name.replace(".xlsx", ""), True)
                df = df.append(resultat)
                # print(f"df \n{df}")
                # print(tabulate(df, headers='keys', tablefmt='psql'))
                # print(tabulate(resultat, headers='keys', tablefmt='psql'))

                # print("columns : ", df1.columns)
                # print("ligne : \n", df1.loc[:, 'N° appui'] != np.NaN)
                # print("ligne : \n", df1.loc[:, 'Nature des travaux'])

                # tableau = np.DataFrame(df1.iloc[:, 0], df1.iloc[:, 31])
                # print(f"tableau {tableau}")
                # except Exception as e:
                #     print(f"Une erreur a été trouvée : \n{e}")
            #
            #         listePoteauFt.append(name)
            #
            #         dossier_parent = os.path.split(os.path.split(os.path.split(cheminComplet)[0])[0])[1]
            #         # print(f"cheminComplet {cheminComplet}", "Export 1")
            #
            # if listePoteauFt:
            #     dicoPoteauFt_SousTraitant[dossier_parent]=listePoteauFt

    # print(tabulate(df, headers='keys', tablefmt='psql'))
    # print("df0\n", df)
    return df


# df = pd.DataFrame(columns=["N° appui", "Nature des travaux", "Fichier"], dtype="float64")
# LectureFichiersExcelsCap_ft(df)
#
