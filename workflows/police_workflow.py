# -*- coding: utf-8 -*-
"""
Workflow pour l'analyse Police C6.
Orchestre la vérification, l'importation et l'analyse des données Police C6.
Note: Ce workflow s'exécute principalement sur le thread principal car il manipule intensivement des objets QGIS non thread-safe.
"""

from qgis.PyQt.QtCore import QObject, pyqtSignal
from qgis.core import Qgis, QgsMessageLog, QgsProject, QgsVectorLayer
from ..PoliceC6 import PoliceC6
from ..qgis_utils import get_layer_safe
import os
import time

class PoliceWorkflow(QObject):
    """
    Contrôleur pour le flux d'analyse Police C6.
    """
    
    # Signaux
    progress_changed = pyqtSignal(int)
    message_received = pyqtSignal(str, str)  # message, couleur
    analysis_finished = pyqtSignal(dict)  # resultats
    error_occurred = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.police_logic = PoliceC6()
        self.current_task = None # Pas de tache async pour l'instant car main thread

    def reset_logic(self):
        """Réinitialise la logique métier"""
        self.police_logic._reset_state()

    def check_layers_exist(self):
        """Vérifie la présence des couches nécessaires"""
        return self.police_logic.verificationnsDonnees()

    def import_gracethd_data(self, repertoire_gthd, sqlite_path=None):
        """
        Importe les données GraceTHD (SQLite ou Shapefiles/CSV).
        Retourne (success, message_erreur)
        """
        if sqlite_path and os.path.exists(sqlite_path):
            self.police_logic.removeGroup('GRACETHD')
            
            table_map = {
                't_cheminement': 't_cheminement_copy',
                't_cableline': 't_cableline_copy',
                't_noeud': 't_noeud_copy',
                't_ptech': 't_ptech_copy',
                't_cable': 't_cable_copy',
                't_sitetech': 't_sitetech_copy',
            }

            ok_all = True
            for src, dest in table_map.items():
                for old in QgsProject.instance().mapLayersByName(dest):
                    QgsProject.instance().removeMapLayer(old.id())

                uri = f"{sqlite_path}|layername={src}"
                lyr = QgsVectorLayer(uri, dest, 'ogr')
                if not lyr.isValid():
                    # On continue pour essayer d'importer les autres, mais on notera l'échec global
                    # Ou alors on considère que c'est optionnel ?
                    # Dans le code original, ok_all devenait False mais continuait.
                    # Cependant, le code original retournait ok_all.
                    ok_all = False
                    continue

                QgsProject.instance().addMapLayer(lyr, False)
                self.police_logic.insertLayerInGroupGraceTHD(lyr)

            if ok_all:
                return True, ""
            
            # Si SQLite échoue (partiellement ou totalement), on essaie le dossier ?
            # Le code original retournait True si ok_all est True, sinon
            # il continuait vers removeGroup('GRACETHD') puis essayait SHP/CSV si appelant continuait?
            # En fait dans Plc6ImportationDonneesDansQgis original:
            # if sqlite_ok:
            #    ...
            #    ok_sqlite = self._plc6_import_gracethd_sqlite(sqlite_path)
            #    if ok_sqlite:
            #        return True
            #    self.police.removeGroup('GRACETHD')
            
            # Donc si pas ok_all, on nettoie et on continue vers SHP
            self.police_logic.removeGroup('GRACETHD')

        listeCoucheSHP = ["t_cheminement", "t_cableline", "t_noeud"]
        manque = []
        
        # Données Géographiques
        for shp in listeCoucheSHP:
            try:
                full_path = os.path.join(repertoire_gthd, shp) # + .shp géré par ajouterCoucherShp ou détecté
                # ajouterCoucherShp attend un chemin sans extension ou avec ? 
                # Dans PoleAerien c'était f"{repertoireGTHD}/{shp}" et ajouterCoucherShp ajoute .shp
                self.police_logic.ajouterCoucherShp(full_path)
            except Exception:
                manque.append(shp)
            
            if not self.police_logic.coucheShp.isValid():
                manque.append(shp)

        # Données CSV
        liste_absent = self.police_logic.ajouterCoucherCsv(repertoire_gthd)
        
        if manque or liste_absent:
            missing = list(manque)
            if liste_absent:
                missing.extend(liste_absent)
            
            self.police_logic.removeGroup("GRACETHD")
            return False, f"Données manquantes: {missing}"
            
        return True, ""

    def apply_style(self, style_name):
        """Applique un style QGIS via la logique métier"""
        self.police_logic.appliquerstyle(style_name)

    def run_analysis(self, params):
        """
        Exécute l'analyse complète (Séquentiel Main Thread).
        
        Args:
            params (dict): Paramètres d'analyse
                - fname: Chemin fichier Excel
                - bpe: Nom couche BPE
                - filterValeur: Valeur filtre (nom étude)
                - attaches: Nom couche Attaches
                - table_etude: Nom couche Etudes
                - colonne_etude: Colonne Etudes
                - zone_layer_name: Nom couche Zone (optionnel)
        """
        try:
            self.progress_changed.emit(5)
            self.message_received.emit("************** PRESENCE DES APPUIS ********************", "grey")
            
            fname = params['fname']
            table = params['table_etude']
            colonne = params['colonne_etude']
            filterValeur = params['filterValeur']
            bpe = params['bpe']
            attaches = params['attaches']
            zone_layer_name = params.get('zone_layer_name')

            # 1. Lecture fichiers et analyse préliminaire
            liste_cable_appui_OD, infNumPoteauAbsent = self.police_logic.lireFichiers(
                fname, table, colonne, filterValeur, bpe, attaches, zone_layer_name
            )
            
            self.progress_changed.emit(15)
            self.police_logic.removeGroup(f"ERROR_{filterValeur}")
            self.progress_changed.emit(20)

            # Rapport Correspondances
            if self.police_logic.nb_appui_corresp >= 1:
                self.police_logic.potInfNumPresent.sort()
                chaine = ", ".join(self.police_logic.potInfNumPresent)
                msg = f"{self.police_logic.nb_appui_corresp} correspondance(s) trouvé(s) :\n{chaine}"
                self.message_received.emit(msg, "green")
            else:
                self.message_received.emit("Aucune correspondance trouvée", "black")

            # Rapport C6 -> QGIS
            self.message_received.emit("*** Annexe C6 --> Données QGIS ***", "grey")
            self.progress_changed.emit(25)
            
            if self.police_logic.nb_appui_absentPot >= 1:
                self.police_logic.absence.sort()
                chaine = ", ".join(self.police_logic.absence)
                msg = f"{self.police_logic.nb_appui_absentPot} Les appuis suivants ont été trouvés dans Annexe C6 mais pas dans QGIS (infra_pt_pot) :\n{chaine}"
                self.message_received.emit(msg, "orange") # Default color for info/warning
            else:
                self.message_received.emit("Tous les appuis dans C6 existent également dans infra_pt_pot", "green")

            # Rapport QGIS -> C6
            self.message_received.emit("*** Données QGIS  --> Annexe C6 ***", "grey")
            
            if self.police_logic.nb_appui_absent > 0:
                poteaux = "infra_pt_pot"
                msg_base = "appui n'existe pas" if self.police_logic.nb_appui_absent == 1 else "appuis n'existent pas"
                chaineInf_num = ", ".join(infNumPoteauAbsent)
                
                # Création couche erreur
                condition = tuple(self.police_logic.infNumPotAbsent) if len(self.police_logic.infNumPotAbsent) > 1 else f"({self.police_logic.infNumPotAbsent[0]})"
                
                # Note: createNewLayer est dans SecondFile, pas PoliceC6. 
                # Il faudra passer une callback ou injecter SecondFile.
                # Pour l'instant, on suppose que PoliceWorkflow a accès à SecondFile ou on émet un signal pour créer la couche.
                # Le plus propre est d'ajouter createNewLayer à PoliceC6 ou de le passer en paramètre.
                # Pour simplifier le refactoring, on va émettre un signal spécifique pour la création de couche erreur
                # ou déléguer au 'result' final.
                
                msg = f"{self.police_logic.nb_appui_absent} {msg_base} dans le fichier Annexe C6 \n{chaineInf_num}"
                self.message_received.emit(msg, "orange")
                self.message_received.emit(f"Voir la couche : error_{poteaux}", "orange")
                
                # Signal spécial pour demander à l'UI de créer la couche erreur (car nécessite SecondFile)
                # Ou mieux, on retourne les infos nécessaires dans analysis_finished
            else:
                self.message_received.emit(f"Tous les appuis dans infra_pt_pot sont dans Annexe C6 {filterValeur}", "green")

            self.progress_changed.emit(30)
            self.message_received.emit("************** EXTREMITES CABLES - APPUIS *************", "grey")

            # 2. Analyse Câbles
            cable_corresp, nbre_EntiteLigne = self.police_logic.analyseAppuiCableAppui(
                liste_cable_appui_OD, table, colonne, filterValeur, zone_layer_name
            )
            
            self.progress_changed.emit(60)

            if cable_corresp > 0:
                self.message_received.emit(f"{cable_corresp} correspondance(s) trouvé(s)", "green")
                self.message_received.emit("Ligne \t  Appuis-O \t  Capa \t  Appuis-D", "green")
                for item in self.police_logic.listeCableAppuitrouve:
                    self.message_received.emit(f"{item[0]}\t  {item[1]}\t  {item[2]}\t  {item[3]}", "green")
            else:
                self.message_received.emit("Aucune correspondance", "black")

            self.progress_changed.emit(65)
            self.message_received.emit("*** Annexe C6 --> Données QGIS ***", "grey")

            total_cables = len(liste_cable_appui_OD)
            if total_cables > 0:
                 self.message_received.emit(f"{total_cables} liaisons (appui-capa-appui) sont présents dans Annexe C6 mais pas dans QGIS", "black")
                 self.message_received.emit("Ligne \t  Appuis-O \t  Capa \t  Appuis-D", "black")
                 for item in liste_cable_appui_OD:
                     self.message_received.emit(f"{item[0]}\t  {item[1]}\t  {item[2]}\t  {item[3]}", "black")
            else:
                msg = "Toutes les liaisons cables-appuis dans Annexe C6 sont dans QGIS" if cable_corresp else "Aucun liaison trouvé dans Annexe C6"
                color = "green" if cable_corresp else "black"
                self.message_received.emit(msg, color)

            self.progress_changed.emit(70)
            self.message_received.emit("*** Données QGIS  --> Annexe C6 ***", "green")
            self.progress_changed.emit(75)

            if nbre_EntiteLigne:
                self.message_received.emit(f"{nbre_EntiteLigne} éléments n'ont pas trouvé de correspondance avec Annexe C6", "black")
                self.message_received.emit("Appuis-O \t\tCapa \tAppuis-D", "black")
                for pt_orig, capa, pt_dest in self.police_logic.listeAppuiCapaAppuiAbsent:
                    self.message_received.emit(f"{pt_orig}\t  {capa}\t  {pt_dest}", "black")
                self.message_received.emit("Voir la couche 'error_appui_capa_appui'", "orange")
            else:
                msg = "Toutes les données QGIS sont dans Annexe C6" if cable_corresp else "Aucun liaison dans la zone d'étude"
                color = "green" if cable_corresp else "black"
                self.message_received.emit(msg, color)

            self.progress_changed.emit(80)
            self.message_received.emit("************** PRESENCE EBP (boites) ******************", "grey")

            # 3. Analyse EBP
            if self.police_logic.presence_liste_appui_ebp:
                if self.police_logic.nb_pbo_corresp > 0:
                    self.message_received.emit(f"{self.police_logic.nb_pbo_corresp} correspondance(s) trouvé(s)", "green")
                    self.message_received.emit("Appuis \tType de boîtes (C7)", "green")
                    for item in self.police_logic.bpo_corresp:
                        self.message_received.emit(f"{item[1]} \t{item[2]}", "green")
                else:
                    self.message_received.emit("Aucune correspondance", "black")
            else:
                self.message_received.emit("Annexe C7 ne contient pas d'appui avec EBP", "black")

            self.progress_changed.emit(90)
            self.message_received.emit("*** Annexe C6 --> Données QGIS ***", "grey")

            total_ebp_absent = len(self.police_logic.liste_appui_ebp)
            if total_ebp_absent > 0:
                msg = f"{total_ebp_absent} appui(s) avec boîte(s) sont absents de BD :"
                self.message_received.emit(msg, "black")
                self.message_received.emit("Ligne \tAppui \tBoite", "black")
                for item in self.police_logic.liste_appui_ebp:
                    self.message_received.emit(f"{item[0]}\t{item[1]} \t{item[2]}", "black")
            
            if self.police_logic.presence_liste_appui_ebp and total_ebp_absent == 0:
                self.message_received.emit("Tous les appuis avec EBP dans Annexe C6 sont dans QGIS", "green")
            
            if not self.police_logic.presence_liste_appui_ebp:
                self.message_received.emit("Aucun", "orange")

            self.message_received.emit("*** Données QGIS  --> Annexe C6 ***", "grey")
            
            # EBP non appui
            if self.police_logic.ebp_non_appui:
                # Signal pour création couche erreur
                # TODO: Gérer la création de couche via le result
                pass
                
            self.progress_changed.emit(95)

            # EBP appui inconnu
            if self.police_logic.ebp_appui_inconnu:
                msg = f"{len(self.police_logic.ebp_appui_inconnu)} appui(s) avec EBP sont absents d'Annexe C6 :"
                self.message_received.emit(msg, "black")
                self.message_received.emit("Appuis \t \tBoîtes", "black")
                for item in self.police_logic.ebp_appui_inconnu:
                    self.message_received.emit(f"{item[1]} \t{item[2]}", "black")
                self.message_received.emit("Voir la couche 'error_infra_pot_EBP'", "orange")

            # Finalisation
            self.progress_changed.emit(100)
            
            # Construction objet résultat complet pour actions UI (création couches, styles, etc.)
            result = {
                'success': True,
                'filterValeur': filterValeur,
                'infNumPotAbsent': list(self.police_logic.infNumPotAbsent),
                'nb_appui_absent': self.police_logic.nb_appui_absent,
                'ebp_non_appui': list(self.police_logic.ebp_non_appui),
                'fied_id_Ebp': self.police_logic.fied_id_Ebp,
                'ebp_appui_inconnu': list(self.police_logic.ebp_appui_inconnu)
            }
            self.analysis_finished.emit(result)

        except Exception as e:
            QgsMessageLog.logMessage(f"Erreur PoliceWorkflow: {e}", "PoleAerien", Qgis.Critical)
            self.error_occurred.emit(str(e))
