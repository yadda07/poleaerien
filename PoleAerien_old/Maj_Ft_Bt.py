#!/usr/bin/python
# -*- coding: utf-8 -*-

from qgis.core import Qgis, QgsProject, QgsFeatureRequest, QgsExpression, NULL
import os
import re
import numpy as np
import pandas as pd
from tabulate import tabulate

from .fonctions_utiles import fonctions_utiles
# from openpyxl.styles import PatternFill
# from functools import reduce


class MajFtBt:
    """ Comparaison des données C6 Par rapport à la base de données"""

    def __init__(self):
        """Le constructeur de ma classe
        Il prend pour attribut de classe les *** """

    def LectureFichiersExcelsFtBtKo(self, fichier_Excel):
        """Fonction pour parcourir les fichiers Excel pour renseigner la référence des appuis."""

        # fichier_Excel = f"C:/Users/asoumare/Documents/Teletravail/Developpement/CAP_COMAC/D17-10/FT-BT KO D17-02.xlsx"
        names_ft = ["Nom Etudes", "N° appui", "Action", "inf_mat_replace", "Etiquette jaune", "Zone privée", "Transition aérosout"]
        names_bt = ["Nom Etudes", "N° appui", "Action", "typ_po_mod", "Zone privée", "Portée molle"]

        # name = os.path.basename(fichier_c6)
        try:
            with pd.ExcelFile(fichier_Excel) as xls:
                # On prend les données à partir de la ligne 7 (la colonne).
                df_ft = pd.read_excel(xls, "FT", header=0, index_col=None, names=names_ft)
                df_bt = pd.read_excel(xls, "BT", header=0, index_col=None, names=names_bt)

            ######################################### FT #######################################
            # On donne la liste des colonnes que l'on souhaite garder
            df_ft = df_ft.loc[:, ["Nom Etudes", "N° appui", "Action", "inf_mat_replace"]]

            df_ft['etat'] = df_ft['Action'].apply(lambda x: "A RECALER" if "recalage" in str(x).lower()
            else ("A REMPLACER" if "remplacement" in str(x).lower()
                  else ("A RENFORCER" if "renforcement" in str(x).lower() else '')))

            df_ft['action'] = df_ft['Action'].apply(lambda x: "RECALAGE" if "recalage" in str(x).lower()
            else ("REMPLACEMENT" if "remplacement" in str(x).lower()
                  else ("RENFORCEMENT" if "renforcement" in str(x).lower() else '')))

            df_ft["Nom Etudes"] = df_ft["Nom Etudes"].fillna(method='ffill', axis=0).str.upper()

            ######################################### BT #######################################
            df_bt = df_bt.loc[:, ["Nom Etudes", "N° appui", "Action", "typ_po_mod"]]
            df_bt['etat'] = df_bt['Action'].apply(lambda x: "BT KO" if "implantation" in str(x).lower() else "")
            df_bt['inf_type'] = df_bt['Action'].apply(lambda x: "POT-AC" if "implantation" in str(x).lower() else "")
            df_bt['action'] = df_bt['Action'].apply(lambda x: "IMPLANTATION" if "implantation" in str(x).lower() else "")
            df_bt['inf_propri'] = df_bt['Action'].apply(lambda x: "SMPN" if "implantation" in str(x).lower() else "")
            # df_ft.fillna(method='bfill', axis="Etude CAPFT", inplace=True)  # columns=["Etude CAPFT"],
            # df_ft[["Etude CAPFT"]] = df_ft[["Etude CAPFT"]].fillna(method='ffill', axis=0)
            # Remplacer les champs de la colonne vide par les valeurs pércédentes.
            df_bt["Nom Etudes"] = df_bt["Nom Etudes"].fillna(method='ffill', axis=0).str.upper()

        except AttributeError as lettre:
            print(f"FICHIER : {fichier_Excel}\nLettre : {lettre}")
            df_ft = pd.DataFrame({})
            df_bt = pd.DataFrame({})

        except Exception as e:
            print(f"Erreur avec ce fichier {fichier_Excel}\nPrécision : {e}")
            df_ft = pd.DataFrame({})
            df_bt = pd.DataFrame({})

        print(f"df_ft :\n", tabulate(df_ft, headers="keys", tablefmt="psql"))
        print(f"df_bt :\n", tabulate(df_bt, headers="keys", tablefmt="psql"))

        return df_ft, df_bt

    def liste_poteau_etudes(self, table_poteau, table_cap_ft, table_comac):
        """Récupérer les données liées à la base de données"""
        infra_pt_pot = QgsProject.instance().mapLayersByName(table_poteau)[0]
        spindex_infra_pt_pot = fonctions_utiles.create_qgsspatialindex(infra_pt_pot)
        decoupage_cap_ft = QgsProject.instance().mapLayersByName(table_cap_ft)[0]
        decoupage_comac = QgsProject.instance().mapLayersByName(table_comac)[0]

        listePoteaux = []
        listePoteauxComplet = []
        # listeEtat = []
        nomEtudesFt = []
        gidFt = []

        # requete = QgsExpression("inf_type LIKE 'POT-FT'")
        # request = QgsFeatureRequest(requete)
        # clause = QgsFeatureRequest.OrderByClause('inf_type', ascending=True)
        # orderby = QgsFeatureRequest.OrderBy([clause])
        # request.setOrderBy(orderby)

        for feat_cap_ft in decoupage_cap_ft.getFeatures():
            etudes = str(feat_cap_ft["nom_etudes"]).upper()
            geom_cap_ft = feat_cap_ft.geometry()
            inter = spindex_infra_pt_pot.intersects(geom_cap_ft.boundingBox())
            if inter:
                for id in inter:
                    feat_pot = infra_pt_pot.getFeature(id)
                    if feat_pot["inf_type"] == "POT-FT":

            #for feat_pot in infra_pt_pot.getFeatures(request):
                        if feat_cap_ft.geometry().contains(feat_pot.geometry()):

                            # ATTENTION CHANGEMENT DE FORMALISME POUR inf_num !!
                            inf_num = feat_pot["inf_num"]

                            nomEtudesFt.append(etudes)
                            gidFt.append(feat_pot["gid"])
                            listePoteauxComplet.append(inf_num)
                            position = inf_num.find('FT-')
                            inf_num = int(inf_num[position + 3:])

                            listePoteaux.append(str(inf_num))
                            # listeEtat.append(np.NaN if feat_pot["etat"] == NULL else feat_pot["etat"])

        df_ft = pd.DataFrame({'gid': gidFt, 'N° appui': listePoteaux,  # "etat": listeEtat,
                              'inf_num': listePoteauxComplet, "Nom Etudes": nomEtudesFt},
                             dtype="float64")

        # print(f"df_ft :\n", tabulate(df_ft, headers="keys", tablefmt="psql"))

        # requete = QgsExpression("inf_type LIKE 'POT-BT'")
        # request = QgsFeatureRequest(requete)
        # clause = QgsFeatureRequest.OrderByClause('inf_type', ascending=True)
        # orderby = QgsFeatureRequest.OrderBy([clause])
        # request.setOrderBy(orderby)

        listePoteauxBT = []
        listePoteauxCompletBT = []
        listeEtat = []
        nomEtudesBt = []
        gidBt = []

        for feat_comac in decoupage_comac.getFeatures():
            etudesBT = str(feat_comac["nom_etudes"]).upper()
            geom_comac = feat_comac.geometry()
            inter = spindex_infra_pt_pot.intersects(geom_comac.boundingBox())
            if inter:
                for id in inter:
                    feat_pot = infra_pt_pot.getFeature(id)
                    if feat_pot["inf_type"] == "POT-BT":

            #for feat_pot in infra_pt_pot.getFeatures(request):
                        if feat_comac.geometry().contains(feat_pot.geometry()):
                            # ATTENTION CHANGEMENT DE FORMALISME POUR inf_num !!
                            inf_num = feat_pot["inf_num"]

                            nomEtudesBt.append(etudesBT)
                            gidBt.append(feat_pot["gid"])
                            listePoteauxCompletBT.append(inf_num)
                            position = inf_num.find('BT-')
                            inf_numBT = inf_num[position:]

                            listePoteauxBT.append(str(inf_numBT))
                            listeEtat.append(np.NaN if feat_pot["etat"] == NULL else feat_pot["etat"])

        df_bt = pd.DataFrame({"gid": gidBt, 'N° appui': listePoteauxBT, 'inf_num': listePoteauxCompletBT,
                              "Nom Etudes": nomEtudesBt}, index=gidBt,
                             dtype="float64")

        # print(f"df_bt QGIS :\n", tabulate(df_bt, headers="keys", tablefmt="psql"))

        return df_ft, df_bt

    def comparerLesDonnees(self, excel_df_ft, excel_df_bt, bd_df_ft, bd_df_bt):
        # print("bd_ft", bd_df_ft)
        # print("excel_ft", excel_df_ft)
        """Compare les données Dataframe du fichier Excel par rapport à la base de données"""
        ######################################### FT #######################################
        liste_valeur_introuvbl_ft = pd.DataFrame({})
        liste_valeur_trouve_ft = pd.DataFrame({})

        # df3 = (excel_df_ft != bd_df_ft).any(1)
        df_ft = pd.merge(excel_df_ft, bd_df_ft, how="left", on=["N° appui", "Nom Etudes"], indicator=True)
        df_ft.fillna({"etat": "", "Action": "", "inf_mat_replace": ""}, inplace=True)

        tt_valeur_introuvable_ft = np.sum(df_ft['_merge'] == "left_only")
        tt_valeur_trouve_ft = np.sum(df_ft['_merge'] == "both")

        # print(f"tt_valeur_introuvable_ft : {tt_valeur_introuvable_ft}")
        # print(f"tt_valeur_trouve_ft : {tt_valeur_trouve_ft}")
        # Pas de correspondance existants : Présence dans Excel, mais absent de QGIS
        if tt_valeur_introuvable_ft > 0:
            liste_valeur_introuvbl_ft = df_ft.loc[:, ["Nom Etudes", "N° appui", "Action", "inf_mat_replace"]].loc[(df_ft["_merge"] == "left_only")]
            liste_valeur_introuvbl_ft.fillna("")

        # Nbre de correspondance trouve
        if tt_valeur_trouve_ft > 0:
            liste_valeur_trouve_ft = df_ft.loc[:, ["gid", "inf_num", "action", "etat", "inf_mat_replace"]].loc[
                (df_ft["_merge"] == "both")]
            liste_valeur_trouve_ft = liste_valeur_trouve_ft.set_index('gid')

            # print(f"resultat :\n", tabulate(liste_valeur_trouve_ft, headers="keys", tablefmt="psql"))
        liste_ft = [tt_valeur_introuvable_ft, liste_valeur_introuvbl_ft, tt_valeur_trouve_ft, liste_valeur_trouve_ft]

        ######################################### BT #######################################
        liste_valeur_introuvbl_bt = pd.DataFrame({})
        liste_valeur_trouve_bt = pd.DataFrame({})

        df_bt = pd.merge(excel_df_bt, bd_df_bt, how="left", on=["N° appui", "Nom Etudes"], indicator=True)
        df_bt.fillna({"etat": "", "Action": "", "typ_po_mod": ""}, inplace=True)

        tt_valeur_introuvable_bt = np.sum(df_bt['_merge'] == "left_only")
        tt_valeur_trouve_bt = np.sum(df_bt['_merge'] == "both")

        # Pas de correspondance existants : Présence dans Excel, mais absent de QGIS
        if tt_valeur_introuvable_bt > 0:
            liste_valeur_introuvbl_bt = df_bt.loc[:, ["Nom Etudes", "N° appui", "Action", "typ_po_mod"]].loc[
                (df_bt["_merge"] == "left_only")]
            # print(f"APRES\n: {tt_valeur_introuvable_bt.dtype}")
            # print(f"liste_valeur_introuvbl_bt :\n", tabulate(liste_valeur_introuvbl_bt, headers="keys", tablefmt="psql"))

        # Nbre de correspondance trouve
        if tt_valeur_trouve_bt > 0:
            df_bt['inf_num'] = df_bt['inf_num'].str.replace('BT', 'PN')
            # df.replace({'A': r'^ba.$'}, {'A': 'new'}, regex=True)
            liste_valeur_trouve_bt = df_bt.loc[:, ["gid", "inf_num", "inf_propri", "inf_type", "action", "etat", "typ_po_mod"]].loc[
                (df_bt["_merge"] == "both")]
            # set column as index
            liste_valeur_trouve_bt = liste_valeur_trouve_bt.set_index('gid')
            # print(f"df_bt trouve :\n", tabulate(df_bt, headers="keys", tablefmt="psql"))
            # print(f"resultat trouve :\n", tabulate(liste_valeur_trouve_bt, headers="keys", tablefmt="psql"))

        liste_bt = [tt_valeur_introuvable_bt, liste_valeur_introuvbl_bt, tt_valeur_trouve_bt, liste_valeur_trouve_bt]

        return liste_ft, liste_bt

    def miseAjourFinalDesDonnees(self, table_poteau, df):
        """Mise à jour des données de la base table infra_pt_pot"""
        infra_pt_pot = QgsProject.instance().mapLayersByName(table_poteau)[0]

        # On récupére la liste des columns qui seront modifié
        listesColumns = list(df.columns.values)  # this will always work in pandas

        changementNomColumns = {}
        for column in listesColumns:
            # On récupère la position du champs (colonne) des appuis à remplacer
            idx_inf_num = infra_pt_pot.dataProvider().fields().indexFromName(str(column))
            changementNomColumns[column] = idx_inf_num

        # print(f"changementNomColumns : {changementNomColumns}")
        # On change le nom des colonnes par rapport à leurs positions.
        df = df.rename(columns=changementNomColumns, errors="raise")

        # Mise à jour des données.
        infra_pt_pot.dataProvider().changeAttributeValues(df.to_dict('index'))
        infra_pt_pot.updateExtents()
        infra_pt_pot.updateFields()
