#!/usr/bin/python
# -*-coding:Utf-8-*-

# Second pour contenir quelques fonctions qui seront utilisées dans le fichier principale Police.py
from qgis.core import (Qgis, QgsProject, QgsExpression, QgsFeatureRequest, QgsVectorLayer, QgsFeature, QgsGeometry,
                       QgsLayerTreeLayer, QgsPointXY, NULL)

import xlrd
import os
import re
# from shapely.geometry import *
# import csv
# import math


class PoliceC6:
    """ Classe qui fait tout le travail pour accéder à la base de données,
    Réaliser des requêtes dans la base de données,
    Exporter le résultat des requêtes vers le fichier Excel """

    def __init__(self):
        """Le constructeur de ma classe
        Il prend pour attribut de classe les *** """

    def barreP(self, numero):
        """ Pour faire appel à la barre de progression """
        return f"self.dlg.progressBar.setValue({numero})"

    def insertLayerInGroupGraceTHD(self, couche):
        """Fonction qui permet d'ajout une couche dans un groupe d'erreur"""

        root = QgsProject.instance().layerTreeRoot()
        group = root.findGroup(u"GRACETHD")

        if not group:
            groupName = u"GRACETHD"
            root = QgsProject.instance().layerTreeRoot()
            group = root.addGroup(groupName)

        else:
            pass

        group.insertChildNode(1, QgsLayerTreeLayer(couche))

    def layerGroupError(self, couche, nomEtude):
        """Fonction qui permet d'ajout une couche dans un groupe d'erreur"""
        root = QgsProject.instance().layerTreeRoot()
        group = root.findGroup(f"ERROR_{nomEtude}")

        if not group:
            groupName = f"ERROR_{nomEtude}"
            root = QgsProject.instance().layerTreeRoot()
            group = root.addGroup(groupName)

        else:
            pass

        group.insertChildNode(1, QgsLayerTreeLayer(couche))

    def removeGroup(self, name):
        """Suppression de groupe de couche du nom de GraceTHD"""
        root = QgsProject.instance().layerTreeRoot()
        group = root.findGroup(name)
        if not group is None:
            for child in group.children():
                dump = child.dump()
                id = dump.split("=")[-1].strip()
                QgsProject.instance().removeMapLayer(id)
            root.removeChildNode(group)

    def lectureFichierExcel(self, fname):
        """Fonction qui permet de lire le contenu du fichier Excel et de renvoyer ses contenus"""
        # Les valeurs retenues commencent à partir de la colonne A et ligne 8  soit (A8).
        # On créé une liste qui contient toutes valeurs des appuis à remplacer
        fname = fname
        # Le nom du fichier à analyser

        document = xlrd.open_workbook(fname)
        # La premier feuille du fichier

        # Feuille 4 "Export 1"
        feuille_1 = document.sheet_by_index(3)

        # cols = feuille_1.ncols
        rows = feuille_1.nrows
        champs_xlsx = []
        liste_cable_appui_OD = []  #
        self.liste_appui_ebp = []  # Liste des appuis ayant des boîtes
        num_appui_o = ''

        # Les valeurs retenues commencent à partir de la colonne A et ligne 8  soit (A8).
        # On créé une liste qui contient toutes valeurs des appuis à remplacer

        for r in range(8, rows):
            num_appui_start = str(feuille_1.cell_value(rowx=r, colx=0))  # num_appui origine
            capa_cable = str(feuille_1.cell_value(rowx=r, colx=18))  # capacité du cable
            num_appui_end = str(feuille_1.cell_value(rowx=r, colx=24))  # Appui destination

            # Valeur associée à la présence ou non de la boîte.  Colonne (AJ) du Excel
            pbo = str(feuille_1.cell_value(rowx=r, colx=35))

            # Si le numéro n'appui se terminer par '.0', float
            regexp = re.compile(r'\.0')
            if regexp.search(num_appui_start):
                num_appui_start = num_appui_start[:-2]

            regexp = re.compile(r'^0')
            if regexp.search(num_appui_start):
                num_appui_start = num_appui_start[1:]

            num_appui_origine = num_appui_start

            # Si le numéro de l'appui existe (non vide)
            if num_appui_start:
                if num_appui_start.islower():
                    pass

                # PRESENCE DES APPUIS
                # Pour les appuis
                # On ne tient pas compte des BT (Appuis d'Orange) et des FAC (Façades)
                if re.match('BT', num_appui_start) is None and re.match('FAC', num_appui_start) is None:
                    champs_xlsx.append(num_appui_start)

                # Sauvegarder de la valeur du num_appui non vide, servira à remplacer les num_appuis vide
                num_appui_o = num_appui_origine

            # EXTREMITES CABLES - APPUIS :
            # On ne tient pas compte des façades (FAC)
            if re.match('FAC', num_appui_o) is None and re.match('FAC', num_appui_end) is None:

                # Il faut que la valeur associée au champs du cable soit supérieur ou égal à 10
                # On ne tient pas compte de l'espace vide
                if len(str(capa_cable.replace(' ', ''))) >= 12:
                    # Les valeurs retenues sont les 3 premières lettres avant le mot 'F-'
                    positionf = capa_cable.find('F-')
                    extrait = capa_cable[positionf-3:positionf].replace(' ', '')

                    # Si le num appui est vide, on prend la valeur du num appui précédent qui n"était pas nul
                    extrait_capa_cab = int(extrait)  # La capacité du cable (valeur en entier)

                    # ='Saisies terrain'!$F$9:$N$4000

                    # Si la colonne du num appui est vide, on récupère la précédente valeur associée au num_appui (ligne A)
                    if not num_appui_origine:
                        num_appui_o = num_appui_o

                    liste_cable_appui_OD.append([(r+1), num_appui_o, int(extrait_capa_cab), str(num_appui_end), str(pbo)])

                # PRESENCE EBP (boites)
                # Pour sauvegarder les valeurs liées à la présence des boîtes (PBO)
                # A préciser, les valeurs en prendre en compte (PBO, PEO, etc.) voire Fabien
                if str(pbo).upper() == "PB" or str(pbo).upper() == "PEO":
                    self.liste_appui_ebp.append([(r+1), num_appui_o, str(pbo)])

        self.presence_liste_appui_ebp = True if self.liste_appui_ebp else False

        return champs_xlsx, liste_cable_appui_OD

    def lireFichiers(self, fname, table, colonne, valeur, t_bpe, t_attaches):
        """Fonction pour parcourir les fichiers Excel pour renseigner la référence des appuis."""

        bpe = QgsProject.instance().mapLayersByName(t_bpe)[0]
        # La table (Linestring) qui servira à intersecter avec les bpe
        attaches = QgsProject.instance().mapLayersByName(t_attaches)[0]

        self.nb_appui_corresp = 0  # Comptabilise les appuis présent dans les deux données (infra_pt_pot et Annexe C6)
        self.nb_pbo_corresp = 0  # Comptabilise les pbo présent dans les deux données (infra_pt_pot et Annexe C6)
        self.bpo_corresp = []  # Sauvegarde des pbo présent dans les deux données (infra_pt_pot et Annexe C6)
        self.nb_appui_absent = 0  # Comptabilise les appuis absents dans les fichiers C6 mais présent dans infra_pt_pot
        self.nb_appui_absentPot = 0  # Comptabilise les appuis absents dans infra_pt_pot mais présent dans C6
        self.potInfNumPresent = []
        self.infNumPotAbsent = []
        infNumPoteauAbsent = []
        self.ebp_non_appui = []  # EBP qui sont dans la zone d'étude mais n'itersecte aucun appui
        valeursmanquant = []
        self.absence = []

        # Résultat de la fonction qui réalise la lecture du fichier Excel et renvoie ses contenus
        champs_xlsx, liste_cable_appui_OD = self.lectureFichierExcel(fname)

        colonne = colonne
        condition = valeur
        requete = QgsExpression(f"{colonne} LIKE \'{condition}\'")

        self.idPotPresent = []  # id des poteaux présents dans la zone géographique et dans  le fichier Excel
        self.idPotAbsent = []  # id des poteaux présents dans la zone géographique mais pas dans le fichier Excel.

        # Les valeurs retenues commencent à partir de la colonne A et ligne 8  soit (A8).
        # On créé une liste qui contient toutes valeurs des appuis à remplacer

        infra_pt_pot = ""  # La table (point) concernant les infra_pt_pot
        infra_pt_chb = ""  # La table (chambre) pour intersecte avec les EBP qui n'intersecte pas les appuis
        etude_cap_ft = ""  # La table (polygone) qui servira à intersecter avec les infra_pt_pot
        bpe_pot = []
        t_cheminement_copy = ""

        self.bpe_pot_cap_ft = []
        self.ebp_appui_inconnu = []  # Pour sauvegarder les appuis avec ebp qui ne serait pas dans Annexe C6

        bufferDist = 0.5  # Buffer servira à intersecter les tables bpe et infra_pt_pot. unité en mètre (0,1m soit 10cm)

        # Toutes les couches qui sont dans le projet QGIS
        mesCouches = QgsProject.instance().mapLayers()
        for couche in mesCouches.values():
            # Les appuis
            if str(couche.name()).lower() == "infra_pt_pot":
                infra_pt_pot = couche

            # Les chambres
            if str(couche.name()).lower() == "infra_pt_chb":
                infra_pt_chb = couche

            if couche.name() == table:
                etude_cap_ft = couche

            if couche.name() == "t_cheminement_copy":
                t_cheminement_copy = couche

        # On récupère le champs id qui sera à extraire les valeurs dans la table EBP

        # Colonne id de EBP
        field_index = bpe.fields().indexFromName("gid")
        # Si le champs gid n'existe pas, on prend, on suppose que c'est 'id'
        self.fied_id_Ebp = "id" if field_index == -1 else "gid"

        # On enlève d'abord la sélection qui serait en cours.
        infra_pt_pot.selectByIds([])
        etude_cap_ft.selectByIds([])

        # pour renseigner les appuis à remplacer
        idx_inf_num = infra_pt_pot.dataProvider().fields().indexFromName('inf_num')

        # Pour les appuis adjacents, hors de la zone d'étude
        for feat_t_chem in t_cheminement_copy.getFeatures():
            adjacent = False  # Par défaut pas d'adjacent de trouvé

            for feat_cap_ft in etude_cap_ft.getFeatures(QgsFeatureRequest(requete)):
                # Le cheminement doit intersecter la zone d'étude ...
                if feat_t_chem.geometry().intersects(feat_cap_ft.geometry()):

                    if feat_t_chem.geometry().contains(feat_cap_ft.geometry()):
                        print("Impossible")
                        pass

                    # ... mais il ne doit pas être contenu dans la zone d'étude
                    else:
                        geom = feat_t_chem.geometry().asMultiPolyline()  # .asMultiPolyline  ou  asPolyline
                        start_point = QgsPointXY(geom[0][0])  # l'origine du cable
                        chem_geom_start_point = QgsGeometry.fromPointXY(start_point).buffer(bufferDist, 0)

                        # Si
                        if chem_geom_start_point.intersects(feat_cap_ft.geometry()):
                            pass

                        else:
                            for feat_pot_origine in infra_pt_pot.getFeatures():
                                # Si infra_pt_pot intersecte avec la table polygones
                                # Il faut que la géométrie existe

                                if feat_pot_origine.geometry():

                                    if feat_pot_origine.geometry().intersects(chem_geom_start_point):
                                        # Cette partie se sert potentiellement à rien.

                                        if feat_pot_origine[idx_inf_num] != NULL and re.match('POT', feat_pot_origine[idx_inf_num]) is None:
                                            chaine = feat_pot_origine[idx_inf_num]

                                            # print("chaine : ", chaine)
                                            if re.match('POT', chaine):
                                                position = chaine.find('-FT-') + 1
                                                chaine = chaine[position:]
                                                if re.match('^0', chaine):
                                                    chaine = chaine[1:]
                                                    # print("chaine près :", chaine)
                                                valeursmanquant.append(chaine)

                                            else:
                                                chaine = feat_pot_origine[idx_inf_num]
                                                # print("chaine : ", chaine)
                                                if re.match('^0', chaine):
                                                    chaine = chaine[1:]
                                                    # print("chaine près :", chaine)
                                                valeursmanquant.append(chaine)

                                            if chaine in champs_xlsx and chaine not in self.potInfNumPresent:
                                                self.nb_appui_corresp += 1
                                                # Pour éviter des doublons lors du renseignement des appuis à remplacer
                                                self.potInfNumPresent.append(chaine)
                                                self.idPotPresent.append(feat_pot_origine.id())
                                                adjacent = True

                        if not adjacent:
                            # L'intermité destination du cheminement
                            end_point = QgsPointXY(geom[-1][0])  # l'origine du cable
                            chem_geom_end_point = QgsGeometry.fromPointXY(end_point).buffer(bufferDist, 0)

                            if chem_geom_end_point.intersects(feat_cap_ft.geometry()):
                                pass

                            else:
                                for feat_pot_destination in infra_pt_pot.getFeatures():
                                    if feat_pot_destination.geometry():

                                        #if not adjacent:
                                            # Si infra_pt_pot intersecte avec la table polygones
                                        if feat_pot_destination.geometry().intersects(chem_geom_end_point):
                                            if feat_pot_destination[idx_inf_num] != NULL and re.match('POT', feat_pot_destination[idx_inf_num]) is None:
                                                chaine = feat_pot_destination[idx_inf_num]
                                                # Toutes les valeurs trouvées sont ici stockées pour servir de comparaison aux valeurs qui n'existent pas
                                                valeursmanquant.append(chaine)

                                                if chaine in champs_xlsx and chaine not in self.potInfNumPresent:
                                                    self.nb_appui_corresp += 1
                                                    # Pour éviter des doublons lors du renseignement des appuis à remplacer
                                                    self.potInfNumPresent.append(chaine)
                                                    self.idPotPresent.append(feat_pot_destination.id())
                                                    adjacent = True

        # Réquete sur la table géométrie
        # Filtrage de la table polygone (etude_cap_ft) pour ne choisir que la zone géographique qui nous concerne.
        for feat_cap_ft in etude_cap_ft.getFeatures(QgsFeatureRequest(requete)):
            cands = infra_pt_pot.getFeatures(QgsFeatureRequest().setFilterRect(feat_cap_ft.geometry().boundingBox()))

            for feat_pot in cands:
                # Si infra_pt_pot intersecte avec la table polygones
                if feat_pot.geometry().intersects(feat_cap_ft.geometry()):

                    # Je parcours ligne par ligne la colonne inf_num qui ne contient pas 'POT' et dont inf_num non vide
                    if feat_pot[idx_inf_num] != NULL and 'FT' in feat_pot[idx_inf_num]:

                        chaine = feat_pot[idx_inf_num]

                        # if len(chaine)>= 12:
                        regexp = re.compile(r'\-FT\-')
                        if regexp.search(chaine):
                            position = chaine.find('FT-') + 3
                            chaine = chaine[position:]
                            # valeursmanquant.append(chaine)

                        # Toutes les valeurs trouvées sont ici stockées pour servir de comparaison aux valeurs qui n'existent pas
                        valeursmanquant.append(chaine)

                        adjacent = False

                        # Test si une valeur du point technique se trouve dans la liste des valeurs du fichier Excel
                        if chaine in champs_xlsx and chaine not in self.potInfNumPresent:
                            self.nb_appui_corresp += 1
                            # Pour éviter des doublons lors du renseignement des appuis à remplacer
                            self.potInfNumPresent.append(chaine)
                            self.idPotPresent.append(feat_pot.id())
                            adjacent = True

                        else:
                            pass

                        # Si aucune correspondance n'a été trouvé
                        if not adjacent:
                            self.nb_appui_absent += 1
                            self.infNumPotAbsent.append(feat_pot.id())
                            infNumPoteauAbsent.append(chaine)
                            self.idPotAbsent.append(feat_pot.id())

                        # Intersection pour vérifier présence d'ebp ou non
                        for feat_bpe in bpe.getFeatures():

                            # il faut que les bpe soient d'abord dans la zone étude
                            if feat_bpe.geometry().intersects(feat_cap_ft.geometry()):

                                # Sauvegarde des EBP qui sont dans la zone d'étude
                                bpe_pot.append(feat_bpe.id())

                                # buffer autour du bpe
                                feat_bpe_buffer = feat_bpe.geometry().buffer(bufferDist, 0)

                                # Intersection des EBP avec les appuis
                                if feat_bpe_buffer.intersects(feat_pot.geometry()):

                                    self.ebp_appui_inconnu.append([feat_pot.id(), feat_pot[idx_inf_num], feat_bpe['noe_type']])
                                    self.bpe_pot_cap_ft.append(feat_bpe.id())

                                    # Pour sauvegarder les ebp qgis qui ne sont pas dans Annexe C6
                                    # On vérifie que le fichier Excel a déjà des données EBP
                                    if self.liste_appui_ebp:
                                        # Parcours des données pbo récupérer dans le fichier Annexe C6
                                        for compte, valeur in enumerate(self.liste_appui_ebp):

                                            # str(feat_bpe['noe_type'])[:2] pour récupérer juste les 2 premiers valeurs
                                            # Ex : PB sur PBO
                                            # Comparaison des données pbo de l'Annexe C6 avec les données QGIS
                                            if valeur[1] == chaine:
                                                self.nb_pbo_corresp += 1
                                                self.bpo_corresp.append(self.liste_appui_ebp[compte])

                                                # On supprime de liste des pbo Annexe C6 celles dont les correspondances ont été trouvées
                                                del self.liste_appui_ebp[compte]

                                                # Suppression des appuis-EBP qui serait trouvé
                                                for iteration, [_, inf_num, _] in enumerate(self.ebp_appui_inconnu):
                                                    # 'inf_num, noe_type
                                                    if inf_num == feat_pot[idx_inf_num]:
                                                        del self.ebp_appui_inconnu[iteration]
                                                break

                                else:
                                    # S'il y a des attaches
                                    for feat_attaches in attaches.getFeatures():
                                        geom_attaches = feat_attaches.geometry().asPolyline()  # .asMultiPolyline  ou  asPolyline
                                        att_start_point = QgsPointXY(geom_attaches[0])  # l'origine du cable
                                        att_geom_start_point = QgsGeometry.fromPointXY(att_start_point).buffer(bufferDist, 0)

                                        # L'intermité destination du attaches
                                        att_end_point = QgsPointXY(geom_attaches[-1])  # l'origine du cable
                                        att_geom_end_point = QgsGeometry.fromPointXY(att_end_point).buffer(bufferDist, 0)

                                        if feat_bpe.geometry().intersects(att_geom_start_point) or feat_bpe.geometry().intersects(att_geom_end_point):
                                            if feat_pot.geometry().intersects(att_geom_start_point) or feat_pot.geometry().intersects(att_geom_end_point):

                                                self.ebp_appui_inconnu.append([feat_pot.id(), feat_pot[idx_inf_num], feat_bpe['noe_type']])
                                                self.bpe_pot_cap_ft.append(feat_bpe.id())

                                                # Pour sauvegarder les ebp qgis qui ne sont pas dans Annexe C6
                                                # On vérifie que le fichier Excel a déjà des données EBP
                                                if self.liste_appui_ebp:
                                                    # Parcours des données pbo récupérer dans le fichier Annexe C6
                                                    for compte, valeur in enumerate(self.liste_appui_ebp):

                                                        # str(feat_bpe['noe_type'])[:2] pour récupérer juste les 2 premiers valeurs
                                                        # Ex : PB sur PBO
                                                        # Comparaison des données pbo de l'Annexe C6 avec les données QGIS
                                                        if valeur[1] == chaine:
                                                            self.nb_pbo_corresp += 1
                                                            self.bpo_corresp.append(self.liste_appui_ebp[compte])

                                                            # On supprime de liste des pbo Annexe C6 celles dont les correspondances ont été trouvées
                                                            del self.liste_appui_ebp[compte]

                                                            # Suppression des appuis-EBP qui serait trouvé
                                                            for iteration, [_, inf_num, _] in enumerate(self.ebp_appui_inconnu):
                                                                # 'inf_num, noe_type
                                                                if inf_num == feat_pot[idx_inf_num]:
                                                                    del self.ebp_appui_inconnu[iteration]
                                                            break
                                                break

        # EBP qui n'intersecte pas d'appuis. EBP qui sont seuls, chose qui ne devrait pas exister.
        for gid_bpe in bpe_pot:
            if gid_bpe not in self.bpe_pot_cap_ft and gid_bpe not in self.ebp_non_appui:
                self.ebp_non_appui.append(gid_bpe)

        # Si EBP pas sur un appui, vérifie s'il n'intersecte pas une chambre
        if self.ebp_non_appui:
            bufferEbpChambre = 3
            # condition2 = ", ".join(self.ebp_non_appui)
            condition2 = tuple(self.ebp_non_appui) if len(self.ebp_non_appui) > 1 else f"({self.ebp_non_appui[0]})"
            # print(f"condition2 : {condition2}")
            requeteEBP = QgsExpression(f"{self.fied_id_Ebp} IN {condition2}")

            for feat_bpe in bpe.getFeatures(QgsFeatureRequest(requeteEBP)):

                for feat_cap_ft in etude_cap_ft.getFeatures(QgsFeatureRequest(requete)):

                    for feat_chb in infra_pt_chb.getFeatures():
                        if feat_chb.geometry().intersects(feat_cap_ft.geometry()):
                            bufffer_feat_chb = feat_chb.geometry().buffer(bufferEbpChambre, 0)

                            if bufffer_feat_chb.intersects(feat_bpe.geometry()):
                                try:
                                    self.ebp_non_appui.remove(feat_bpe.id())
                                except ValueError:
                                    pass
                                break

        # On vérifie si l'une des données existantes dans les fichiers Excels n'existent pas dans infra_pt_pot
        for absence in champs_xlsx:
            if absence not in valeursmanquant:
                self.nb_appui_absentPot += 1
                # Save numéros des appuis absents dans C3A
                # idée : abord, vérifier que les valeurs absence dans le fichier sont des entiers. Erreur connu "NoneType"
                self.absence.append(absence)

        infra_pt_pot.select(self.idPotPresent)

        return liste_cable_appui_OD, infNumPoteauAbsent

    def appliquerstyle(self, nomcouche):
        """Fonction qui permet d'appliquer un style au fichier erreur qui sera généré """
        # On applique un style au fichier C3A qui aurait été généré
        cheminstyle = ''

        cheminAbsolu = "~/AppData/Roaming/QGIS/QGIS3/profiles/default/python/plugins/PoliceC6/styles/"

        if re.match(u'error_infra_pt_pot_', nomcouche):
            cheminstyle = os.path.expanduser(cheminAbsolu + "infra_pt_pot.qml")

        if re.match(u'error_infra_pt_pot_ebp', nomcouche):
            cheminstyle = os.path.expanduser(cheminAbsolu + "infra_pt_pot_ebp.qml")

        if re.match(u'error_appui_capa_appui', nomcouche):
            cheminstyle = os.path.expanduser(cheminAbsolu + "t_cable.qml")

        if re.match(u'error_bpe', nomcouche):
            cheminstyle = os.path.expanduser(cheminAbsolu + "bpe.qml")

        verifstyle = True
        self.messagestyle = ''
        # On vérifie si le fichier contenant le style existe pour l'appliquer au fichier C3A

        if verifstyle:
            mesCouches = QgsProject.instance().mapLayers()
            for coucheId, couche in mesCouches.items():
                if couche.name() == nomcouche:
                    try:
                        with open(cheminstyle):
                            couche.loadNamedStyle(cheminstyle)
                            self.messagestyle = u"Le style a bien été appliquée au fichier C3A"
                            # pass
                    except IOError:
                        verifstyle = False
                        self.messagestyle = u"Le style n'a pu été trouvé. A faire manuellement"

    def verificationnsDonnees(self):
        """Fonction qui vérifie que dans QGIS toutes les tables nécessaires existent déjà.
        Sinon, le programme ne s'éxecutera pas"""
        # Temps de démarrage
        liste_data_present = []  # Lister des données déjà présentes dans QGIS

        liste_import_data = ["bpe", "infra_pt_pot", "infra_pt_chb", "attaches"]  # "cables"
        mesCouches = QgsProject.instance().mapLayers()

        for couche in mesCouches.values():
            liste_data_present.append(str(couche.name()).lower())

        # Test si parmi les données attendues, certaines sont absent.
        # Stocker des valeurs absentes dans QGIS
        liste_absent = [data for data in liste_import_data if data not in liste_data_present]

        return liste_absent

    def analyseAppuiCableAppui(self, liste_cable_appui_OD, t_etude_cap_ft, champs,  valeur):
        """Fonction principale pour l'analyse des relations appuis - cables - appuis dans QGIS et dans Annexe C6 """
        cable_corresp = 0
        # self.non_cable_corresp = 0
        # self.id_non_cable_corresp = []  # Pour enregistrer les cables et leurs appuis qui n'ont pas trouvé de correspondance

        self.listeCableAppuitrouve = []
        self.listeAppuiCapaAppuiAbsent = []  # Liste des appuis et leurs capa qui sont absent d'Annexe C6

        dicoPointsTechniquesCompletes = self.listePointTechniquesPoteaux("t_ptech_copy", "t_noeud_copy", "t_sitetech_copy")
        dicoCableLine = self.listeInfoCablesLines("t_cable_copy", "t_cableline_copy")

        t_chem = QgsProject.instance().mapLayersByName("t_cheminement_copy")[0]
        etude_cap_ft = QgsProject.instance().mapLayersByName(t_etude_cap_ft)[0]  # le lovage associé à chaque cable

        id_cable_chem_trouve = []  # Liste id des correspondances trouvés
        bufferDistExtremite = 0.5

        requete = QgsExpression(f"{champs} LIKE \'{valeur}\'")

        epsg = t_chem.crs().postgisSrid()  # SRID
        uri = (f"LineString?crs=epsg:{epsg}&field=id_cable_chem:string&field=appui_start:string&field=cab_capa:integer&"
               f"field=appui_end:string&field=cb_typelog:string&field=erreur:string&&index=yes")
        new_table_appui_capa_appui = QgsVectorLayer(uri, 'error_appui_capa_appui', 'memory')
        feat = QgsFeature()
        pr = new_table_appui_capa_appui.dataProvider()

        # self.barreP(35)

        # Parcourir d'abord la table des t_cable (GraceTHD) :
        # for feat_cb_line in t_cableline.getFeatures():
        for [_, cl_cb_code, cb_typelog, cb_capafo, feat_cb_line_geom] in dicoCableLine.values():

            # Parcourir  la table des etude_cap_ft en filtrant les données
            for feat_etude_cap_ft in etude_cap_ft.getFeatures(QgsFeatureRequest(requete)):
                # On prend les géométries des t_cableline qui intersectent la zone d'étude "etude_cap_ft"
                if feat_cb_line_geom.intersects(feat_etude_cap_ft.geometry()):

                    # Créer un buffer de 2 mètres autour de la table t_cableline
                    buffer_t_cableline = feat_cb_line_geom.buffer(bufferDistExtremite, 0)

                    # Parcourir  la table des t_cheminement
                    for feat_t_chem in t_chem.getFeatures():

                        # Le buffer de la géométrie doit contenir la table t_cheminement, pour être valide
                        if buffer_t_cableline.contains(feat_t_chem.geometry()):

                            # Le t_cheminement doit intersecter la zone d'étude
                            # for feat_etude_cap_ft_2 in etude_cap_ft.getFeatures(QgsFeatureRequest(requete)):
                            # On prend les géométries des t_cheminement qui intersecte la zone d'étude
                            if feat_t_chem.geometry().intersects(feat_etude_cap_ft.geometry()):

                                [inf_num_o, pt_typephy_o] = dicoPointsTechniquesCompletes[feat_t_chem['cm_ndcode1']]
                                # print(f"inf_num_o : {inf_num_o} pt_nd_code : {pt_nd_code} pt_typephy : {pt_typephy} ")
                                # print(f"inf_num_o : {inf_num_o}")
                                [inf_num_d, pt_typephy_d] = dicoPointsTechniquesCompletes[feat_t_chem['cm_ndcode2']]

                                # Les chambres ne sont prise en compte dans les C6, dont on ne les comparent pas.
                                if pt_typephy_o != "C" and pt_typephy_d != "C":
                                    # il faut l'extremité du cheminement au moins au poteau de type FT
                                    if (("POT" in inf_num_o and "FT" in inf_num_o) or
                                            ("POT" in inf_num_d and "FT" in inf_num_d)):
                                        id_cable_chem = f"{cl_cb_code} {feat_t_chem['cm_code']}"

                                        # MAJ de la nouvelle géométrie qui doit être créer
                                        feat.setAttributes([id_cable_chem, inf_num_o, cb_capafo, inf_num_d, cb_typelog, u"Introuvable dans C6"])
                                        feat.setGeometry(feat_t_chem.geometry())
                                        pr.addFeatures([feat])

        # self.barreP(40)

        # Liste des appuis et de leurs capa qui ont été trouvé dans la zone d'étude
        # Ssi la table 'résultat' contient des données
        if new_table_appui_capa_appui.featureCount() > 0:
            existant = []

            for [ligne, origC6, capaC6, destC6, _] in liste_cable_appui_OD:
                # ligne, num_appui_o, extrait_capa_cab, num_appui_end, pbo
                # ligne = v[0]  # La ligne correspondante dans l'annexe C6
                # origC6 = str(v[1])  # Origine de l'appui dans l'annexe C6
                # capaC6 = int(v[2])  # Capacité du cable entre les deux appuis dans l'annexe C6
                # destC6 = str(v[3])  # Destination de l'appui dans l'annexe C6

                for feat_resultat in new_table_appui_capa_appui.getFeatures():

                    ide = feat_resultat['id_cable_chem']
                    pt_orig = str(feat_resultat['appui_start'])
                    capa = int(feat_resultat['cab_capa'])
                    pt_dest = str(feat_resultat['appui_end'])

                    # Si le champs inf_num contient FT, on le prend que les valeurs situées après FT-
                    regexp = re.compile(r'\-FT\-')
                    if regexp.search(pt_orig):
                        position = pt_orig.find('FT-') + 3
                        pt_orig = pt_orig[position:]

                    if regexp.search(pt_dest):
                        position = pt_dest.find('FT-') + 3
                        pt_dest = pt_dest[position:]

                    regexp_bt = re.compile(r'BT*')
                    # Les BT sont pris en charges uniquement si cela concerne uniquement une des deux extremités
                    # du câble.
                    # Si l'extremité d'origine contient 'BT'
                    if regexp_bt.search(origC6) and not regexp_bt.search(destC6):
                        #  on vérifie que les l'autre extrémité et capacité du cable correspondent aux données QGIS
                        if (capaC6 == capa and destC6 == pt_orig) or (capaC6 == capa and destC6 == pt_dest):

                            id_cable_chem_trouve.append(ide)

                            cable_corresp += 1
                            # Sauvegarde les valeurs qui correspondent dans les deux fichiers QGIS et (Annexe C6)
                            self.listeCableAppuitrouve.append([ligne, origC6, capaC6, destC6])
                            existant.append([ligne, origC6, capaC6, destC6])
                            break

                    # Si l'extremité de destination contient 'BT'
                    elif not regexp_bt.search(origC6) and regexp_bt.search(destC6):
                        #  on vérifie que les l'autre extrémité et capacité du cable correspondent aux données QGIS
                        if (capaC6 == capa and origC6 == pt_orig) or (capaC6 == capa and origC6 == pt_dest):

                            id_cable_chem_trouve.append(ide)

                            cable_corresp += 1
                            # Sauvegarde les valeurs qui correspondent dans les deux fichiers QGIS et (Annexe C6)
                            self.listeCableAppuitrouve.append([ligne, origC6, capaC6, destC6])
                            existant.append([ligne, origC6, capaC6, destC6])
                            break

                    # Si les trois vs du fichier Annexe C6 correspondant au cable
                    # et ses deux intersections d'appuis
                    elif ((origC6 == pt_orig and capaC6 == capa and destC6 == pt_dest) or
                          (destC6 == pt_orig and capaC6 == capa and origC6 == pt_dest)):
                        id_cable_chem_trouve.append(ide)

                        # myIndex.append(i)

                        cable_corresp += 1
                        # cm_code_corresp.append(contenu[3])
                        # Sauvegarde les valeurs qui correspondent dans les deux fichiers QGIS et (Annexe C6)
                        self.listeCableAppuitrouve.append([ligne, origC6, capaC6, destC6])
                        existant.append([ligne, origC6, capaC6, destC6])
                        break

            # Suppressions des correspondances trouvées dans Annexe C6
            for x in range(10):
                for a, b in enumerate(existant):
                    for i, v in enumerate(liste_cable_appui_OD):
                        if v[1] == b[1] and v[2] == b[2] and v[3] == b[3]:
                            del liste_cable_appui_OD[i]
                            del existant[a]
                            break

            # Suppression des correspondances trouvées dans QGIS
            for feat_resultat in new_table_appui_capa_appui.getFeatures():
                if id_cable_chem_trouve:
                    for v in id_cable_chem_trouve:
                        if v == feat_resultat['id_cable_chem']:
                            # Suppression des valeurs trouvés couche error
                            new_table_appui_capa_appui.startEditing()
                            new_table_appui_capa_appui.deleteFeature(feat_resultat.id())
                            new_table_appui_capa_appui.commitChanges()
        # self.barreP(50)

        # Enregistrer les id du t_cheminement qui sont dans la zone d'étude
        nbre_EntiteLigne = new_table_appui_capa_appui.featureCount()

        # Si le nombre d'étité est 0, on supprimer la couche, sinon on l'a ajoute dans le projet
        if nbre_EntiteLigne == 0:
            QgsProject.instance().removeMapLayer(new_table_appui_capa_appui)

        else:
            # Liste des appuis et leurs capa sont dans QGIS mais pas dans Annexe C6
            for feat_res in new_table_appui_capa_appui.getFeatures():
                pt_orig = feat_res['appui_start']
                capa = int(feat_res['cab_capa'])
                pt_dest = feat_res['appui_end']
                self.listeAppuiCapaAppuiAbsent.append([pt_orig, capa, pt_dest])

            # Ajout de la carte error_appui_capa_appui dans QGIS
            QgsProject.instance().addMapLayer(new_table_appui_capa_appui, False)

            # Appliquer le style à la couche error_apppui_capa_appui
            self.layerGroupError(new_table_appui_capa_appui, valeur)
            self.appliquerstyle("error_appui_capa_appui")

        # QgsProject.instance().removeMapLayer(infra_pot_copy)
        return cable_corresp, nbre_EntiteLigne

    def ajouterCoucherShp(self, coucheShp):
        """Fonction pour importer des couches géographiques dans QGIS"""
        # Le nom associé au fichier importé dans QGIS

        nom = os.path.basename(coucheShp)  # Extraction du nom du repertoire où est stocké le fichier
        self.coucheShp = QgsVectorLayer(f"{coucheShp}.shp", f"{nom}_copy", "ogr")
        crs = self.coucheShp.crs()
        crs.createFromId(2154)
        self.coucheShp.setCrs(crs)

        QgsProject.instance().addMapLayer(self.coucheShp, False)
        self.insertLayerInGroupGraceTHD(self.coucheShp)

        # return self.coucheShp

    def ajouterCoucherCsv(self, repertoireGTHD):
        """Fonction pour importer des couches géographiques dans QGIS"""
        # Le nom associé au fichier importé dans QGIS

        listeTableCsv = ["t_ptech.csv", "t_cable.csv", "t_sitetech.csv"]
        liste_import_data = []
        # Importation des données shapefile
        for table in listeTableCsv:
            # On parcours les dossiers et leurs sous dossiers
            for subdir, dirs, files in os.walk(repertoireGTHD + os.sep):

                # Chaque dossier est representé par une liste qui contient le nom de tous les fichiers qui s'y trouve
                for name in files:
                    # On veut récuperer uniquement des shape
                    if name.endswith('.csv'):

                        if table == name:
                            shape = os.path.splitext(name)[0]
                            tableCsv0 = f"file:{subdir}{os.sep}{name}?delimiter=;"
                            layer = QgsVectorLayer(tableCsv0, f"{shape}_copy", "delimitedtext")
                            QgsProject.instance().addMapLayer(layer, False)
                            self.insertLayerInGroupGraceTHD(layer)
                            liste_import_data.append(f"{name}")
                            break

        liste_absent = [data for data in liste_import_data if data not in listeTableCsv]
        return liste_absent

    def listeInfoCablesLines(self, nomTableCable, nomTablet_CableLines):
        """Fonction qui permet récuperer l'ensemble des informations liées aux cables"""

        t_cableline = QgsProject.instance().mapLayersByName(nomTablet_CableLines)[0]
        t_cable = QgsProject.instance().mapLayersByName(nomTableCable)[0]  # le lovage associé à chaque cable

        requete = QgsExpression("cb_etiquet NOT LIKE '%XXX%'")
        request = QgsFeatureRequest(requete)
        clause = QgsFeatureRequest.OrderByClause('cb_etiquet', ascending=True)
        orderby = QgsFeatureRequest.OrderBy([clause])
        request.setOrderBy(orderby)
        dictionnaireCableLine = {}

        # On récupère toutes les valeurs dans t_cable
        for feat_cable in t_cable.getFeatures(request):
            cb_code = feat_cable["cb_code"]
            # cb_nd1 = feat_cable["cb_nd1"]  # Le noeud origine du câble
            # cb_nd2 = feat_cable["cb_nd2"]  # Le noeud destination du câble
            # t_cable[gid] = [cb_code, cb_etiquet, cb_nd1, cb_nd2, cb_capafo]

            ########################## T_CABLELINE ##############################################
            # Récupérer la longueur associée à chaque cable
            for feat_cableline in t_cableline.getFeatures():
                if cb_code == feat_cableline["cl_cb_code"]:

                    # if feat_cableline.geometry() and feat_cableline.geometry() != NULL:
                    cb_etiquet = feat_cable["cb_etiquet"]
                    cb_typelog = feat_cable["cb_typelog"]
                    cb_capafo = int(feat_cable["cb_capafo"])
                    cl_cb_code = feat_cableline["cl_cb_code"]
                    # cb_longueur = math.ceil(feat_cableline.geometry().length())
                    cb_geom = feat_cableline.geometry()
                    dictionnaireCableLine[cb_code] = [cb_etiquet, cl_cb_code, cb_typelog, cb_capafo, cb_geom]
                    break

        return dictionnaireCableLine

    def listePointTechniquesPoteaux(self, nomTablet_ptech, nomTablet_noeud, nomTablett_sitetech):
        ################################ POINTS TECHNIQUES ###############################################
        dicoPointsTechniquesCompletes = {}
        # Trier par ordre croissants les contenus par rapport au champs pt_etiquet
        request = QgsFeatureRequest()
        clause = QgsFeatureRequest.OrderByClause('pt_etiquet', ascending=True)
        orderby = QgsFeatureRequest.OrderBy([clause])
        request.setOrderBy(orderby)

        # Liste complète : [sro, nro, collecte, nd_code, pt_etiquet, pt_code, pt_typephy, pt_prop, pt_avct, pt_ad_code,
        # pt_nature, pt_gest, pt_geom, pt_geometry]
        # colonne_t_ptech = ["pt_etiquet", "pt_nd_code", "pt_typephy"]
        colonne_t_ptech = ["pt_etiquet", "pt_typephy"]

        t_ptech = QgsProject.instance().mapLayersByName(nomTablet_ptech)[0]
        t_noeud = QgsProject.instance().mapLayersByName(nomTablet_noeud)[0]

        for feat_ptech in t_ptech.getFeatures(request):
            pt_nd_code = feat_ptech["pt_nd_code"]

            for feat_noeud in t_noeud.getFeatures():
                if feat_ptech["pt_nd_code"] == feat_noeud["nd_code"]:  # feat_ptech["feat_ptech"] == 'A' and

                    listeValeurs_T_ptech = []

                    for colonne_t_pt in colonne_t_ptech:
                        if feat_ptech[colonne_t_pt] == NULL:
                            listeValeurs_T_ptech.append(str())

                        else:
                            listeValeurs_T_ptech.append(feat_ptech[colonne_t_pt])

                    # pt_geom = feat_noeud.geometry()
                    # listeValeurs_T_ptech.append(pt_geom)
                    dicoPointsTechniquesCompletes[pt_nd_code] = listeValeurs_T_ptech
                    break

        ############################# T_SITETECH #############################################
        t_sitetech_origine = QgsProject.instance().mapLayersByName(nomTablett_sitetech)[0]

        # On récupère toutes les valeurs dans t_sitetech
        # colonne_sitech = ["st_nom", "st_nd_code", "st_typephy", "st_prop", "st_avct", "st_ad_code", "st_codeext"]
        colonne_sitech = ["st_nom", "st_typephy"]
        for gid, feat_sitetech in enumerate(t_sitetech_origine.getFeatures()):
            listeValeursSiteTechniques = []
            pt_nd_code = feat_sitetech["st_nd_code"]

            for feat_noeud in t_noeud.getFeatures():
                if feat_sitetech["st_nd_code"] == feat_noeud["nd_code"]:

                    # Récupération des champs qui nous intéresse.
                    for colonne_site_tech in colonne_sitech:
                        if feat_sitetech[colonne_site_tech] == NULL:
                            listeValeursSiteTechniques.append(str())
                        else:
                            listeValeursSiteTechniques.append(feat_sitetech[colonne_site_tech])

                    dicoPointsTechniquesCompletes[pt_nd_code] = listeValeursSiteTechniques

        return dicoPointsTechniquesCompletes
