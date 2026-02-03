#!/usr/bin/python
# -*- coding: utf-8 -*-

import time
from qgis.PyQt.QtCore import QObject, pyqtSignal
from qgis.PyQt.QtWidgets import QApplication
from qgis.core import (
    Qgis, QgsFeatureRequest, QgsExpression,
    QgsFeature, QgsGeometry, QgsPointXY, QgsProject,
    QgsTask, QgsApplication, QgsSpatialIndex, NULL,
    QgsDataSourceUri, QgsMessageLog, QgsVectorDataProvider, QgsWkbTypes
)
from qgis.PyQt.QtSql import QSqlDatabase, QSqlQuery
import pandas as pd
import os
import re
import random
from .qgis_utils import validate_same_crs, get_layer_safe
from .dataclasses_results import (
    ExcelValidationResult, PoteauxPolygoneResult, 
    EtudesValidationResult, ImplantationValidationResult
)

# REQ-MAJ-003: Colonnes obligatoires pour validation structure Excel
COLONNES_FT_REQUISES = ['Nom Etudes', 'N° appui', 'Action', 'inf_mat_replace', 
                        'Etiquette jaune', 'Zone privée', 'Transition aérosout']
COLONNES_BT_REQUISES = ['Nom Etudes', 'N° appui', 'Action', 'typ_po_mod', 
                        'Zone privée', 'Portée molle']

# REQ-C6BD-003/004: Actions autorisées
ACTIONS_FT_AUTORISEES = ['Renforcement', 'Recalage', 'Remplacement']
ACTIONS_BT_AUTORISEES = ['Implantation']


class MajFtBtSignals(QObject):
    """Signaux pour communication UI non-bloquante"""
    progress = pyqtSignal(int)
    message = pyqtSignal(str, str)  # (message, couleur)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)


class MajUpdateTask(QgsTask):
    """
    CRITICAL-001 FIX: Tâche asynchrone pour MAJ BD après confirmation.
    Évite le freeze UI lors de l'écriture en base de données.
    """
    
    def __init__(self, layer_name, data_ft, data_bt):
        super().__init__("MAJ BD FT/BT", QgsTask.CanCancel)
        self.layer_name = layer_name
        self.data_ft = data_ft
        self.data_bt = data_bt
        self.signals = MajFtBtSignals()
        self.exception = None
        self.result = {'ft_updated': 0, 'bt_updated': 0}
    
    def run(self):
        """
        CRITICAL FIX: Ne fait PAS de modification QGIS ici (Worker Thread).
        Les modifications de couche DOIVENT être sur le Main Thread.
        Cette tâche sert uniquement à signaler que l'UI est prête pour la MAJ.
        """
        try:
            # Simuler une courte préparation
            self.signals.progress.emit(random.randint(58, 63))
            self.signals.message.emit("Préparation MAJ BD...", "grey")
            
            if self.isCanceled():
                return False
            
            # Stocker les compteurs pour le callback
            self.result['ft_count'] = len(self.data_ft) if self.data_ft is not None and not self.data_ft.empty else 0
            self.result['bt_count'] = len(self.data_bt) if self.data_bt is not None and not self.data_bt.empty else 0
            
            self.signals.progress.emit(random.randint(64, 68))
            return True
            
        except Exception as e:
            self.exception = str(e)
            QgsMessageLog.logMessage(f"[MajUpdateTask] Erreur: {e}", "MAJ_FT_BT", Qgis.Critical)
            return False
    
    def finished(self, success):
        """
        Callback sur Main Thread - C'est ici que la MAJ QGIS sera déclenchée.
        """
        if success:
            # Émettre les données pour traitement sur Main Thread
            self.signals.finished.emit({
                'layer_name': self.layer_name,
                'data_ft': self.data_ft,
                'data_bt': self.data_bt,
                'ft_count': self.result.get('ft_count', 0),
                'bt_count': self.result.get('bt_count', 0)
            })
        else:
            self.signals.error.emit(self.exception or "Annulé")


class MajFtBtTask(QgsTask):
    """Tache asynchrone pour MAJ FT/BT - traitement spatial en Worker Thread"""

    def __init__(self, params, raw_data=None):
        super().__init__("MAJ FT/BT", QgsTask.CanCancel)
        self.params = params
        self.raw_data = raw_data or {}
        self.signals = MajFtBtSignals()
        self.result = {}
        self.exception = None

    def run(self):
        """Execute en arriere-plan - traitement spatial pur Python (Worker Thread)"""
        try:
            t0_task = time.perf_counter()
            QgsMessageLog.logMessage("[PERF-MAJ] ======= DEBUT TASK ASYNC (Worker Thread) =======", "PoleAerien", Qgis.Info)
            
            maj = MajFtBt()
            self.signals.progress.emit(random.randint(3, 7))
            self.signals.message.emit("Lecture Excel...", "grey")

            # 1. Lecture Excel
            t1 = time.perf_counter()
            excel_ft, excel_bt = maj.LectureFichiersExcelsFtBtKo(
                self.params['fichier_excel']
            )
            t2 = time.perf_counter()
            QgsMessageLog.logMessage(f"[PERF-MAJ] Lecture Excel: {t2-t1:.3f}s", "PoleAerien", Qgis.Info)
            
            if self.isCanceled():
                return False

            self.signals.progress.emit(random.randint(12, 18))
            self.signals.message.emit("Analyse spatiale FT...", "grey")

            # 2. Traitement spatial PUR PYTHON (pas d'accès QGIS)
            t3 = time.perf_counter()
            bd_ft, bd_bt = self._process_spatial_pure_python()
            t4 = time.perf_counter()
            QgsMessageLog.logMessage(f"[PERF-MAJ] Traitement spatial (Worker): {t4-t3:.3f}s", "PoleAerien", Qgis.Info)
            
            if self.isCanceled():
                return False

            self.signals.progress.emit(random.randint(35, 42))
            self.signals.message.emit("Comparaison données...", "grey")

            # 3. Comparaison
            t5 = time.perf_counter()
            liste_ft, liste_bt = maj.comparerLesDonnees(
                excel_ft, excel_bt, bd_ft, bd_bt
            )
            t6 = time.perf_counter()
            QgsMessageLog.logMessage(f"[PERF-MAJ] Comparaison données: {t6-t5:.3f}s", "PoleAerien", Qgis.Info)

            self.result = {
                'liste_ft': liste_ft,
                'liste_bt': liste_bt
            }
            
            t_total = time.perf_counter() - t0_task
            QgsMessageLog.logMessage(f"[PERF-MAJ] === TOTAL TASK ASYNC (Worker): {t_total:.3f}s ===", "PoleAerien", Qgis.Warning)
            
            self.signals.progress.emit(random.randint(48, 52))
            return True

        except Exception as e:
            self.exception = str(e)
            QgsMessageLog.logMessage(f"[PERF-MAJ] ERREUR Task: {e}", "PoleAerien", Qgis.Critical)
            return False

    def _point_in_polygon(self, x, y, vertices):
        """Ray-casting algorithm pour test point-in-polygon (pur Python)."""
        n = len(vertices)
        if n < 3:
            return False
        inside = False
        j = n - 1
        for i in range(n):
            xi, yi = vertices[i]
            xj, yj = vertices[j]
            if ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / (yj - yi) + xi):
                inside = not inside
            j = i
        return inside

    def _point_in_bbox(self, x, y, bbox):
        """Test rapide si point dans bounding box."""
        xmin, ymin, xmax, ymax = bbox
        return xmin <= x <= xmax and ymin <= y <= ymax

    def _process_spatial_pure_python(self):
        """Traitement spatial 100% Python - aucun appel QGIS."""
        poteaux_ft = self.raw_data.get('poteaux_ft', [])
        poteaux_bt = self.raw_data.get('poteaux_bt', [])
        etudes_cap_ft = self.raw_data.get('etudes_cap_ft', [])
        etudes_comac = self.raw_data.get('etudes_comac', [])
        
        # === FT: Point-in-polygon ===
        t1 = time.perf_counter()
        data_ft = []
        for pot in poteaux_ft:
            x, y = pot['x'], pot['y']
            for etude in etudes_cap_ft:
                # Test bbox d'abord (rapide)
                if self._point_in_bbox(x, y, etude['bbox']):
                    # Test précis polygon
                    if self._point_in_polygon(x, y, etude['vertices']):
                        inf_num = pot['inf_num']
                        try:
                            num_court = str(int(str(inf_num)[:7]))
                        except (ValueError, TypeError):
                            num_court = str(inf_num).split("/")[0]
                        data_ft.append({
                            'gid': pot['gid'],
                            'N° appui': num_court,
                            'inf_num': inf_num,
                            'Nom Etudes': etude['nom_etudes']
                        })
                        break  # Un poteau = une seule étude
        t2 = time.perf_counter()
        QgsMessageLog.logMessage(f"[PERF-MAJ] FT spatial: {len(data_ft)} en {t2-t1:.3f}s", "PoleAerien", Qgis.Info)
        
        if self.isCanceled():
            return pd.DataFrame(), pd.DataFrame()
        
        self.signals.progress.emit(random.randint(22, 28))
        self.signals.message.emit("Analyse spatiale BT...", "grey")
        
        # === BT: Point-in-polygon ===
        t3 = time.perf_counter()
        data_bt = []
        for pot in poteaux_bt:
            x, y = pot['x'], pot['y']
            for etude in etudes_comac:
                if self._point_in_bbox(x, y, etude['bbox']):
                    if self._point_in_polygon(x, y, etude['vertices']):
                        inf_num = pot['inf_num']
                        num_court = str(inf_num).split("/")[0]
                        data_bt.append({
                            'gid': pot['gid'],
                            'N° appui': num_court,
                            'inf_num': inf_num,
                            'Nom Etudes': etude['nom_etudes']
                        })
                        break
        t4 = time.perf_counter()
        QgsMessageLog.logMessage(f"[PERF-MAJ] BT spatial: {len(data_bt)} en {t4-t3:.3f}s", "PoleAerien", Qgis.Info)
        
        # Créer DataFrames
        df_ft = pd.DataFrame(data_ft) if data_ft else pd.DataFrame(
            columns=['gid', 'N° appui', 'inf_num', 'Nom Etudes']
        )
        df_bt = pd.DataFrame(data_bt) if data_bt else pd.DataFrame(
            columns=['gid', 'N° appui', 'inf_num', 'Nom Etudes']
        )
        if not df_bt.empty:
            df_bt = df_bt.set_index('gid', drop=False)
        
        return df_ft, df_bt

    def finished(self, success):
        """Callback thread principal"""
        if success:
            self.signals.finished.emit(self.result)
        else:
            self.signals.error.emit(self.exception or "Annule")

    def cancel(self):
        self.signals.message.emit("Annulation...", "orange")
        super().cancel()


class MajFtBt:
    """MAJ FT/BT avec index spatial et threading"""

    def __init__(self):
        pass

    # =========================================================================
    # REQ-MAJ-002: Validation nom fichier
    # =========================================================================
    def valider_nom_fichier(self, fichier_excel: str) -> bool:
        """
        REQ-MAJ-002: Vérifie que le fichier Excel commence par 'FT-BT'.
        
        QA-EXPERT: Validation entrée utilisateur.
        
        Args:
            fichier_excel: Chemin complet du fichier
            
        Returns:
            bool: True si nom valide
            
        Raises:
            ValueError: Si nom non conforme
        """
        if not fichier_excel:
            raise ValueError("[MAJ_BD] Chemin fichier Excel vide")
        
        nom_fichier = os.path.basename(fichier_excel)
        
        if not re.match(r'^FT-BT.*\.xlsx$', nom_fichier, re.IGNORECASE):
            raise ValueError(
                f"[MAJ_BD] Le fichier doit commencer par 'FT-BT' et avoir l'extension .xlsx.\n"
                f"Fichier fourni : {nom_fichier}"
            )
        
        return True

    # =========================================================================
    # REQ-MAJ-003: Validation structure Excel
    # =========================================================================
    def valider_structure_excel(self, df, colonnes_attendues: list, nom_onglet: str) -> dict:
        """
        REQ-MAJ-003: Valide structure DataFrame Excel.
        
        QA-EXPERT: Vérifie colonnes requises présentes.
        
        Args:
            df: DataFrame pandas
            colonnes_attendues: Liste colonnes obligatoires
            nom_onglet: 'FT' ou 'BT'
            
        Returns:
            dict: {'valide': bool, 'manquantes': [], 'en_trop': []}
        """
        if df is None or df.empty:
            return {
                'valide': False,
                'manquantes': colonnes_attendues,
                'en_trop': [],
                'onglet': nom_onglet
            }
        
        colonnes_presentes = set(df.columns)
        colonnes_requises = set(colonnes_attendues)
        
        manquantes = list(colonnes_requises - colonnes_presentes)
        en_trop = list(colonnes_presentes - colonnes_requises)
        
        return {
            'valide': len(manquantes) == 0,
            'manquantes': manquantes,
            'en_trop': en_trop,
            'onglet': nom_onglet
        }

    # =========================================================================
    # REQ-MAJ-001: Contrôle poteaux dans polygones
    # =========================================================================
    def verifier_poteaux_dans_polygones(self, table_poteau, table_etude_cap_ft, 
                                         table_etude_comac) -> PoteauxPolygoneResult:
        """
        REQ-MAJ-001: Vérifie que tous poteaux FT/BT sont dans polygones études.
        
        QA-EXPERT: Validation CRS, null checks, géométries.
        PERF-SPECIALIST: Index spatial O(n log n).
        
        Args:
            table_poteau: Nom couche infra_pt_pot
            table_etude_cap_ft: Nom couche etude_cap_ft
            table_etude_comac: Nom couche etude_comac
            
        Returns:
            PoteauxPolygoneResult avec listes poteaux hors polygones
        """
        result = PoteauxPolygoneResult()
        
        # QA: Récupération sécurisée des couches
        infra_pt_pot = get_layer_safe(table_poteau, "MAJ_BD")
        etude_cap_ft = get_layer_safe(table_etude_cap_ft, "MAJ_BD")
        etude_comac = get_layer_safe(table_etude_comac, "MAJ_BD")
        
        # QA: Validation CRS identiques
        validate_same_crs(infra_pt_pot, etude_cap_ft, "MAJ_BD")
        validate_same_crs(infra_pt_pot, etude_comac, "MAJ_BD")
        
        # PERF: Index spatial études CAP_FT
        idx_cap_ft = QgsSpatialIndex()
        etudes_cap_ft_dict = {}
        for feat in etude_cap_ft.getFeatures():
            if feat.hasGeometry() and not feat.geometry().isNull():
                idx_cap_ft.addFeature(feat)
                etudes_cap_ft_dict[feat.id()] = feat
        
        # PERF: Index spatial études COMAC
        idx_comac = QgsSpatialIndex()
        etudes_comac_dict = {}
        for feat in etude_comac.getFeatures():
            if feat.hasGeometry() and not feat.geometry().isNull():
                idx_comac.addFeature(feat)
                etudes_comac_dict[feat.id()] = feat
        
        # Vérif poteaux FT
        requete_ft = QgsExpression("inf_type LIKE 'POT-FT'")
        for feat_pot in infra_pt_pot.getFeatures(QgsFeatureRequest(requete_ft)):
            # QA: Null checks
            if not feat_pot.hasGeometry() or feat_pot.geometry().isNull():
                continue
            
            inf_num = feat_pot["inf_num"]
            if not inf_num or inf_num == NULL:
                continue
            
            dans_polygone = False
            bbox = feat_pot.geometry().boundingBox()
            candidates = idx_cap_ft.intersects(bbox)
            
            for fid in candidates:
                if fid in etudes_cap_ft_dict:
                    if etudes_cap_ft_dict[fid].geometry().contains(feat_pot.geometry()):
                        dans_polygone = True
                        break
            
            if not dans_polygone:
                pt = feat_pot.geometry().asPoint()
                result.ft_hors_polygone.append((str(inf_num), pt.x(), pt.y()))
        
        # Vérif poteaux BT
        requete_bt = QgsExpression("inf_type LIKE 'POT-BT'")
        for feat_pot in infra_pt_pot.getFeatures(QgsFeatureRequest(requete_bt)):
            # QA: Null checks
            if not feat_pot.hasGeometry() or feat_pot.geometry().isNull():
                continue
            
            inf_num = feat_pot["inf_num"]
            if not inf_num or inf_num == NULL:
                continue
            
            dans_polygone = False
            bbox = feat_pot.geometry().boundingBox()
            candidates = idx_comac.intersects(bbox)
            
            for fid in candidates:
                if fid in etudes_comac_dict:
                    if etudes_comac_dict[fid].geometry().contains(feat_pot.geometry()):
                        dans_polygone = True
                        break
            
            if not dans_polygone:
                pt = feat_pot.geometry().asPoint()
                result.bt_hors_polygone.append((str(inf_num), pt.x(), pt.y()))
        
        return result

    # =========================================================================
    # REQ-MAJ-004: Vérification études existent
    # =========================================================================
    def verifier_etudes_existent(self, df_excel, table_etude_cap_ft, table_etude_comac,
                                  colonne_nom_etude='Nom Etudes', 
                                  champ_qgis='nom_etudes') -> EtudesValidationResult:
        """
        REQ-MAJ-004: Vérifie que études Excel existent dans couches QGIS.
        
        QA-EXPERT: Validation null, case insensitive.
        
        Args:
            df_excel: DataFrame avec colonne études
            table_etude_cap_ft: Nom couche etude_cap_ft
            table_etude_comac: Nom couche etude_comac
            colonne_nom_etude: Nom colonne Excel
            champ_qgis: Nom champ QGIS
            
        Returns:
            EtudesValidationResult
        """
        result = EtudesValidationResult()
        
        if df_excel is None or df_excel.empty:
            return result
        
        # QA: Récupération sécurisée
        etude_cap_ft = get_layer_safe(table_etude_cap_ft, "MAJ_BD")
        etude_comac = get_layer_safe(table_etude_comac, "MAJ_BD")
        
        # Extraire noms études Excel (unique, upper)
        if colonne_nom_etude not in df_excel.columns:
            QgsMessageLog.logMessage(
                f"[MAJ_BD] Colonne '{colonne_nom_etude}' absente du DataFrame",
                "MAJ_FT_BT", Qgis.Warning
            )
            return result
        
        etudes_excel = set(
            str(e).upper().strip() 
            for e in df_excel[colonne_nom_etude].dropna().unique()
            if e and str(e).strip()
        )
        
        # Extraire noms études QGIS CAP_FT
        etudes_cap_ft_qgis = set()
        idx_champ = etude_cap_ft.fields().indexFromName(champ_qgis)
        if idx_champ >= 0:
            for feat in etude_cap_ft.getFeatures():
                nom = feat[champ_qgis]
                if nom and nom != NULL:
                    etudes_cap_ft_qgis.add(str(nom).upper().strip())
        
        # Extraire noms études QGIS COMAC
        etudes_comac_qgis = set()
        idx_champ = etude_comac.fields().indexFromName(champ_qgis)
        if idx_champ >= 0:
            for feat in etude_comac.getFeatures():
                nom = feat[champ_qgis]
                if nom and nom != NULL:
                    etudes_comac_qgis.add(str(nom).upper().strip())
        
        result.etudes_absentes_cap_ft = list(etudes_excel - etudes_cap_ft_qgis)
        result.etudes_absentes_comac = list(etudes_excel - etudes_comac_qgis)
        
        return result

    # =========================================================================
    # REQ-MAJ-006: Vérification type poteau si implantation
    # =========================================================================
    def verifier_type_si_implantation(self, df_bt) -> ImplantationValidationResult:
        """
        REQ-MAJ-006: Vérifie que si Action='Implantation', typ_po_mod est renseigné.
        
        QA-EXPERT: Validation données métier.
        
        Args:
            df_bt: DataFrame onglet BT
            
        Returns:
            ImplantationValidationResult
        """
        result = ImplantationValidationResult()
        
        if df_bt is None or df_bt.empty:
            return result
        
        if 'Action' not in df_bt.columns or 'typ_po_mod' not in df_bt.columns:
            return result
        
        for idx, row in df_bt.iterrows():
            action = str(row.get('Action', '')).lower()
            typ_po_mod = row.get('typ_po_mod', '')
            
            if 'implantation' in action:
                if pd.isna(typ_po_mod) or str(typ_po_mod).strip() == '':
                    result.erreurs_implantation.append({
                        'ligne': idx + 2,  # +2 car header + index 0
                        'etude': row.get('Nom Etudes', ''),
                        'appui': row.get('N° appui', ''),
                        'erreur': 'Action=Implantation mais typ_po_mod vide'
                    })
        
        return result

    def LectureFichiersExcelsFtBtKo(self, fichier_Excel):
        """Fonction pour parcourir les fichiers Excel pour renseigner la référence des appuis."""

        try:
            with pd.ExcelFile(fichier_Excel) as xls:
                # Lire avec header existant (ligne 0)
                df_ft = pd.read_excel(xls, "FT", header=0, index_col=None)
                df_bt = pd.read_excel(xls, "BT", header=0, index_col=None)
            
            # Normaliser noms colonnes (espaces multiples, trim)
            df_ft.columns = df_ft.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
            df_bt.columns = df_bt.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
            
            # Vérifier que colonnes requises existent
            colonnes_ft_requises = ["Nom Etudes", "N° appui", "Action", "inf_mat_replace"]
            colonnes_bt_requises = ["Nom Etudes", "N° appui", "Action", "typ_po_mod", "Portée molle"]
            
            # Vérif colonnes FT
            manquantes_ft = [col for col in colonnes_ft_requises if col not in df_ft.columns]
            if manquantes_ft:
                raise ValueError(f"Colonnes manquantes onglet FT: {manquantes_ft}")
            
            # Vérif colonnes BT
            manquantes_bt = [col for col in colonnes_bt_requises if col not in df_bt.columns]
            if manquantes_bt:
                raise ValueError(f"Colonnes manquantes onglet BT: {manquantes_bt}")

            ######################################### FT #######################################
            # On donne la liste des colonnes que l'on souhaite garder
            df_ft = df_ft.loc[:, ["Nom Etudes", "N° appui", "Action", "inf_mat_replace", 
                                   "Etiquette jaune", "Zone privée", "Transition aérosout"]]

            # REQ-MAJ-007: Gestion FT Implantation
            def get_etat_ft(action):
                action_lower = str(action).lower()
                if "implantation" in action_lower:
                    return "FT KO"
                elif "recalage" in action_lower:
                    return "A RECALER"
                elif "remplacement" in action_lower:
                    return "A REMPLACER"
                elif "renforcement" in action_lower:
                    return "A RENFORCER"
                return ""
            
            def get_action_ft(action):
                action_lower = str(action).lower()
                if "implantation" in action_lower:
                    return "IMPLANTATION"
                elif "recalage" in action_lower:
                    return "RECALAGE"
                elif "remplacement" in action_lower:
                    return "REMPLACEMENT"
                elif "renforcement" in action_lower:
                    return "RENFORCEMENT"
                return ""

            df_ft['etat'] = df_ft['Action'].apply(get_etat_ft)
            df_ft['action'] = df_ft['Action'].apply(get_action_ft)
            # REQ-MAJ-007: Si implantation, inf_type=POT-AC, inf_propri=RAUV
            df_ft['inf_type'] = df_ft['Action'].apply(lambda x: "POT-AC" if "implantation" in str(x).lower() else "")
            df_ft['inf_propri'] = df_ft['Action'].apply(lambda x: "RAUV" if "implantation" in str(x).lower() else "")
            
            # REQ-NOTE-010: Gestion étiquettes jaune/orange et zone privée
            def get_etiquette_jaune(row):
                val = str(row.get('Etiquette jaune', '')).strip().upper()
                return 'oui' if val == 'X' else None
            
            def get_etiquette_orange(action):
                action_lower = str(action).lower()
                return 'oui' if 'recalage' in action_lower else None
            
            def get_zone_privee(row):
                val = str(row.get('Zone privée', '')).strip().upper()
                return 'X' if val == 'X' else None
            
            df_ft['etiquette_jaune'] = df_ft.apply(get_etiquette_jaune, axis=1)
            df_ft['etiquette_orange'] = df_ft['Action'].apply(get_etiquette_orange)
            df_ft['zone_privee'] = df_ft.apply(get_zone_privee, axis=1)

            df_ft["Nom Etudes"] = df_ft["Nom Etudes"].ffill().str.upper()

            ######################################### BT #######################################
            # Inclure Zone privée si présente dans le fichier Excel
            cols_bt_excel = ["Nom Etudes", "N° appui", "Action", "typ_po_mod", "Portée molle"]
            if "Zone privée" in df_bt.columns:
                cols_bt_excel.append("Zone privée")
            df_bt = df_bt.loc[:, cols_bt_excel]
            
            # Traiter toutes les actions BT: implantation, recalage, remplacement, renforcement
            def get_etat_bt(action):
                action_lower = str(action).lower()
                if "implantation" in action_lower:
                    return "BT KO"
                elif "recalage" in action_lower:
                    return "A RECALER"
                elif "remplacement" in action_lower:
                    return "A REMPLACER"
                elif "renforcement" in action_lower:
                    return "A RENFORCER"
                return ""
            
            def get_action_bt(action):
                action_lower = str(action).lower()
                if "implantation" in action_lower:
                    return "IMPLANTATION"
                elif "recalage" in action_lower:
                    return "RECALAGE"
                elif "remplacement" in action_lower:
                    return "REMPLACEMENT"
                elif "renforcement" in action_lower:
                    return "RENFORCEMENT"
                return ""
            
            df_bt['etat'] = df_bt['Action'].apply(get_etat_bt)
            df_bt['inf_type'] = df_bt['Action'].apply(lambda x: "POT-AC" if "implantation" in str(x).lower() else "")
            df_bt['action'] = df_bt['Action'].apply(get_action_bt)
            df_bt['inf_propri'] = df_bt['Action'].apply(lambda x: "RAUV" if "implantation" in str(x).lower() else "")
            
            # REQ-NOTE-010: Gestion étiquette orange pour BT si recalage
            def get_etiquette_orange_bt(action):
                action_lower = str(action).lower()
                return 'oui' if 'recalage' in action_lower else None
            
            df_bt['etiquette_orange'] = df_bt['Action'].apply(get_etiquette_orange_bt)
            # df_ft.fillna(method='bfill', axis="Etude CAPFT", inplace=True)  # columns=["Etude CAPFT"],
            # df_ft[["Etude CAPFT"]] = df_ft[["Etude CAPFT"]].fillna(method='ffill', axis=0)
            # Remplacer les champs de la colonne vide par les valeurs pércédentes.
            df_bt["Nom Etudes"] = df_bt["Nom Etudes"].ffill().str.upper()

        except AttributeError as lettre:
            QgsMessageLog.logMessage(f"FICHIER : {fichier_Excel} - {lettre}", "MAJ_FT_BT", Qgis.Warning)
            df_ft = pd.DataFrame({})
            df_bt = pd.DataFrame({})

        except Exception as e:
            QgsMessageLog.logMessage(f"Erreur fichier {fichier_Excel}: {e}", "MAJ_FT_BT", Qgis.Critical)
            df_ft = pd.DataFrame({})
            df_bt = pd.DataFrame({})

        # Debug logs removed for production

        return df_ft, df_bt

    def liste_poteau_etudes(self, table_poteau, table_cap_ft, table_comac):
        """Extraction poteaux avec index spatial (performance O(n log n))"""
        t0_func = time.perf_counter()
        
        # Validation couches
        lyrs = QgsProject.instance().mapLayersByName
        infra_pot = lyrs(table_poteau)
        dcp_cap_ft = lyrs(table_cap_ft)
        dcp_comac = lyrs(table_comac)

        if not infra_pot or not infra_pot[0].isValid():
            raise ValueError(f"Couche invalide: {table_poteau}")
        if not dcp_cap_ft or not dcp_cap_ft[0].isValid():
            raise ValueError(f"Couche invalide: {table_cap_ft}")
        if not dcp_comac or not dcp_comac[0].isValid():
            raise ValueError(f"Couche invalide: {table_comac}")

        infra_pot = infra_pot[0]
        dcp_cap_ft = dcp_cap_ft[0]
        dcp_comac = dcp_comac[0]

        validate_same_crs(infra_pot, dcp_cap_ft, "MAJ_FT_BT")
        validate_same_crs(infra_pot, dcp_comac, "MAJ_FT_BT")
        
        # PERF METRICS: Compter features
        n_pot = infra_pot.featureCount()
        n_cap_ft = dcp_cap_ft.featureCount()
        n_comac = dcp_comac.featureCount()
        QgsMessageLog.logMessage(f"[PERF-MAJ] Features: infra_pot={n_pot}, cap_ft={n_cap_ft}, comac={n_comac}", "PoleAerien", Qgis.Info)

        # ===== INDEX SPATIAL FT =====
        t1 = time.perf_counter()
        req_ft = QgsFeatureRequest(QgsExpression("inf_type LIKE 'POT-FT'"))
        idx_ft = QgsSpatialIndex()
        poteaux_ft = {}
        n_ft = 0
        for feat in infra_pot.getFeatures(req_ft):
            if feat.hasGeometry():
                idx_ft.addFeature(feat)
                poteaux_ft[feat.id()] = feat
                n_ft += 1
        t2 = time.perf_counter()
        QgsMessageLog.logMessage(f"[PERF-MAJ] Index FT: {n_ft} poteaux en {t2-t1:.3f}s", "PoleAerien", Qgis.Info)
        
        # ===== EXTRACTION FT =====
        t3 = time.perf_counter()
        data_ft = []
        for feat_dcp in dcp_cap_ft.getFeatures():
            if not feat_dcp.hasGeometry():
                continue
            raw_etude = feat_dcp["nom_etudes"]
            etude = str(raw_etude).upper() if raw_etude and raw_etude != NULL else ""
            bbox = feat_dcp.geometry().boundingBox()
            candidates = idx_ft.intersects(bbox)

            for fid in candidates:
                pot = poteaux_ft[fid]
                if feat_dcp.geometry().contains(pot.geometry()):
                    inf_num = pot["inf_num"]
                    try:
                        num_court = str(int(inf_num[:7]))
                    except (ValueError, TypeError):
                        num_court = str(inf_num).split("/")[0]
                    data_ft.append({
                        'gid': pot["gid"],
                        'N° appui': num_court,
                        'inf_num': inf_num,
                        'Nom Etudes': etude
                    })
        t4 = time.perf_counter()
        QgsMessageLog.logMessage(f"[PERF-MAJ] Extraction FT: {len(data_ft)} resultats en {t4-t3:.3f}s", "PoleAerien", Qgis.Info)

        df_ft = pd.DataFrame(data_ft) if data_ft else pd.DataFrame(
            columns=['gid', 'N° appui', 'inf_num', 'Nom Etudes']
        )

        # ===== INDEX SPATIAL BT =====
        t5 = time.perf_counter()
        req_bt = QgsFeatureRequest(QgsExpression("inf_type LIKE 'POT-BT'"))
        idx_bt = QgsSpatialIndex()
        poteaux_bt = {}
        n_bt = 0
        for feat in infra_pot.getFeatures(req_bt):
            if feat.hasGeometry():
                idx_bt.addFeature(feat)
                poteaux_bt[feat.id()] = feat
                n_bt += 1
        t6 = time.perf_counter()
        QgsMessageLog.logMessage(f"[PERF-MAJ] Index BT: {n_bt} poteaux en {t6-t5:.3f}s", "PoleAerien", Qgis.Info)

        # ===== EXTRACTION BT =====
        t7 = time.perf_counter()
        data_bt = []
        for feat_dcp in dcp_comac.getFeatures():
            if not feat_dcp.hasGeometry():
                continue
            raw_etude = feat_dcp["nom_etudes"]
            etude = str(raw_etude).upper() if raw_etude and raw_etude != NULL else ""
            bbox = feat_dcp.geometry().boundingBox()
            candidates = idx_bt.intersects(bbox)

            for fid in candidates:
                pot = poteaux_bt[fid]
                if feat_dcp.geometry().contains(pot.geometry()):
                    inf_num = pot["inf_num"]
                    num_court = str(inf_num).split("/")[0]
                    data_bt.append({
                        'gid': pot["gid"],
                        'N° appui': num_court,
                        'inf_num': inf_num,
                        'Nom Etudes': etude
                    })
        t8 = time.perf_counter()
        QgsMessageLog.logMessage(f"[PERF-MAJ] Extraction BT: {len(data_bt)} resultats en {t8-t7:.3f}s", "PoleAerien", Qgis.Info)

        df_bt = pd.DataFrame(data_bt) if data_bt else pd.DataFrame(
            columns=['gid', 'N° appui', 'inf_num', 'Nom Etudes']
        )
        if not df_bt.empty:
            df_bt = df_bt.set_index('gid', drop=False)

        t_total = time.perf_counter() - t0_func
        QgsMessageLog.logMessage(f"[PERF-MAJ] === TOTAL liste_poteau_etudes: {t_total:.3f}s ===", "PoleAerien", Qgis.Warning)
        
        return df_ft, df_bt

    def comparerLesDonnees(self, excel_df_ft, excel_df_bt, bd_df_ft, bd_df_bt):
        """Compare les données Dataframe du fichier Excel par rapport à la base de données"""
        ######################################### FT #######################################
        liste_valeur_introuvbl_ft = pd.DataFrame({})
        liste_valeur_trouve_ft = pd.DataFrame({})

        excel_df_ft["N° appui"] = excel_df_ft["N° appui"].astype(str)
        bd_df_ft["N° appui"] = bd_df_ft["N° appui"].astype(str)

        # CRITICAL: Exclure lignes sans cles valides pour eviter produit cartesien
        excel_df_ft_valid = excel_df_ft[
            (excel_df_ft["Nom Etudes"].str.strip() != "") &
            (excel_df_ft["N° appui"].str.strip() != "") &
            (excel_df_ft["N° appui"].str.lower() != "nan")
        ]
        bd_df_ft_valid = bd_df_ft[
            (bd_df_ft["Nom Etudes"].str.strip() != "") &
            (bd_df_ft["N° appui"].str.strip() != "") &
            (bd_df_ft["N° appui"].str.lower() != "nan")
        ]

        df_ft = pd.merge(excel_df_ft_valid, bd_df_ft_valid, how="left", on=["N° appui", "Nom Etudes"], indicator=True)
        df_ft.fillna({"etat": "", "Action": "", "inf_mat_replace": ""}, inplace=True)

        tt_valeur_introuvable_ft = (df_ft['_merge'] == "left_only").sum()
        tt_valeur_trouve_ft = (df_ft['_merge'] == "both").sum()

        # Pas de correspondance existants : Présence dans Excel, mais absent de QGIS
        if tt_valeur_introuvable_ft > 0:
            liste_valeur_introuvbl_ft = df_ft.loc[:, ["Nom Etudes", "N° appui", "Action", "inf_mat_replace"]].loc[(df_ft["_merge"] == "left_only")]
            liste_valeur_introuvbl_ft.fillna("")

        # Nbre de correspondance trouve
        if tt_valeur_trouve_ft > 0:
            # COLONNES COMPLETES: gid, inf_num, action, etat, inf_mat_replace + etiquette_jaune, zone_privee, transition_aerosout
            cols_ft = ["gid", "inf_num", "action", "etat", "inf_mat_replace"]
            # Ajouter colonnes optionnelles si présentes dans Excel
            if "Etiquette jaune" in df_ft.columns:
                cols_ft.append("Etiquette jaune")
            if "Zone privée" in df_ft.columns:
                cols_ft.append("Zone privée")
            if "Transition aérosout" in df_ft.columns:
                cols_ft.append("Transition aérosout")
            
            liste_valeur_trouve_ft = df_ft.loc[df_ft["_merge"] == "both", cols_ft].copy()
            
            # CONVERSION X -> 'oui' pour les champs booléens
            if "Etiquette jaune" in liste_valeur_trouve_ft.columns:
                liste_valeur_trouve_ft["etiquette_jaune"] = liste_valeur_trouve_ft["Etiquette jaune"].apply(
                    lambda x: "oui" if str(x).strip().upper() == "X" else "")
            if "Zone privée" in liste_valeur_trouve_ft.columns:
                liste_valeur_trouve_ft["zone_privee"] = liste_valeur_trouve_ft["Zone privée"].apply(
                    lambda x: "X" if str(x).strip().upper() == "X" else "")
            if "Transition aérosout" in liste_valeur_trouve_ft.columns:
                liste_valeur_trouve_ft["transition_aerosout"] = liste_valeur_trouve_ft["Transition aérosout"].apply(
                    lambda x: "oui" if str(x).strip().upper() == "X" else "")
            
            # Etiquette orange si action = recalage
            liste_valeur_trouve_ft["etiquette_orange"] = liste_valeur_trouve_ft["action"].apply(
                lambda x: "oui" if str(x).strip().upper() == "RECALAGE" else "")
            
            liste_valeur_trouve_ft = liste_valeur_trouve_ft.set_index('gid')
        liste_ft = [tt_valeur_introuvable_ft, liste_valeur_introuvbl_ft, tt_valeur_trouve_ft, liste_valeur_trouve_ft]

        ######################################### BT #######################################
        liste_valeur_introuvbl_bt = pd.DataFrame({})
        liste_valeur_trouve_bt = pd.DataFrame({})

        excel_df_bt["N° appui"] = excel_df_bt["N° appui"].astype(str)
        bd_df_bt["N° appui"] = bd_df_bt["N° appui"].astype(str)

        # CRITICAL: Exclure lignes sans cles valides pour eviter produit cartesien
        excel_df_bt_valid = excel_df_bt[
            (excel_df_bt["Nom Etudes"].str.strip() != "") &
            (excel_df_bt["N° appui"].str.strip() != "") &
            (excel_df_bt["N° appui"].str.lower() != "nan")
        ]
        bd_df_bt_valid = bd_df_bt[
            (bd_df_bt["Nom Etudes"].str.strip() != "") &
            (bd_df_bt["N° appui"].str.strip() != "") &
            (bd_df_bt["N° appui"].str.lower() != "nan")
        ]

        df_bt = pd.merge(excel_df_bt_valid, bd_df_bt_valid, how="left", on=["N° appui", "Nom Etudes"], indicator=True)
        df_bt.fillna({"etat": "", "Action": "", "typ_po_mod": ""}, inplace=True)

        tt_valeur_introuvable_bt = (df_bt['_merge'] == "left_only").sum()
        tt_valeur_trouve_bt = (df_bt['_merge'] == "both").sum()

        # Pas de correspondance existants : Présence dans Excel, mais absent de QGIS
        if tt_valeur_introuvable_bt > 0:
            liste_valeur_introuvbl_bt = df_bt.loc[:, ["Nom Etudes", "N° appui", "Action", "typ_po_mod", "Portée molle"]].loc[
                (df_bt["_merge"] == "left_only")]

        # Nbre de correspondance trouve
        if tt_valeur_trouve_bt > 0:
            cols_bt = ["gid", "inf_num", "inf_propri", "inf_type", "action", "etat", "typ_po_mod", "Portée molle"]
            # Ajouter Zone privée si présente dans Excel
            if "Zone privée" in df_bt.columns:
                cols_bt.append("Zone privée")
            liste_valeur_trouve_bt = df_bt.loc[df_bt["_merge"] == "both", cols_bt].copy()
            # Conversion X -> 'X' pour zone_privee BT
            if "Zone privée" in liste_valeur_trouve_bt.columns:
                liste_valeur_trouve_bt["zone_privee"] = liste_valeur_trouve_bt["Zone privée"].apply(
                    lambda x: "X" if str(x).strip().upper() == "X" else "")
            liste_valeur_trouve_bt = liste_valeur_trouve_bt.set_index('gid')

        liste_bt = [tt_valeur_introuvable_bt, liste_valeur_introuvbl_bt, tt_valeur_trouve_bt, liste_valeur_trouve_bt]

        return liste_ft, liste_bt

    def miseAjourFinalDesDonnees(self, table_poteau, df):
        """MAJ FT dans infra_pt_pot via provider."""
        lyrs = QgsProject.instance().mapLayersByName(table_poteau)
        if not lyrs or not lyrs[0].isValid():
            raise ValueError(f"Couche invalide: {table_poteau}")
        infra_pt_pot = lyrs[0]

        if not infra_pt_pot.dataProvider().capabilities() & QgsVectorDataProvider.ChangeAttributeValues:
            raise ValueError("Provider ne supporte pas ChangeAttributeValues")

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

    # =========================================================================
    # REQ-MAJ-007: MAJ FT Implantation avec nouveau nommage et commentaire
    # =========================================================================
    def _prepare_update_ft(self, table_poteau, liste_valeur_trouve_ft):
        lyrs = QgsProject.instance().mapLayersByName(table_poteau)
        if not lyrs or not lyrs[0].isValid():
            raise ValueError(f"Couche invalide: {table_poteau}")
        infra_pt_pot = lyrs[0]

        fields = infra_pt_pot.fields()
        required = {"etat", "inf_type", "inf_propri", "inf_mat", "commentair", "dce", "inf_num"}
        missing = [f for f in required if fields.indexOf(f) == -1]
        if missing:
            raise ValueError(f"Champs requis manquants: {missing}")

        idx_etat = fields.indexOf("etat")
        idx_inf_type = fields.indexOf("inf_type")
        idx_inf_propri = fields.indexOf("inf_propri")
        idx_inf_mat = fields.indexOf("inf_mat")
        idx_dce = fields.indexOf("dce")
        idx_inf_num = fields.indexOf("inf_num")
        idx_commentaire = fields.indexOf("commentair")
        idx_etiquette_jaune = fields.indexOf("etiquette_jaune")
        idx_etiquette_orange = fields.indexOf("etiquette_orange")
        idx_transition_aerosout = fields.indexOf("transition_aerosout")

        gids_pot_ac = []

        t1 = time.perf_counter()
        gids_needed = set(int(g) for g in liste_valeur_trouve_ft.index)
        features_by_gid = {}
        for feat in infra_pt_pot.getFeatures():
            gid_val = feat["gid"]
            if gid_val in gids_needed:
                features_by_gid[gid_val] = feat
        t2 = time.perf_counter()
        QgsMessageLog.logMessage(
            f"[PERF-MAJ-BD] Index features FT: {len(features_by_gid)} en {t2-t1:.3f}s",
            "PoleAerien", Qgis.Info
        )

        if not infra_pt_pot.isEditable():
            infra_pt_pot.startEditing()

        return {
            "layer": infra_pt_pot,
            "features_by_gid": features_by_gid,
            "gids_pot_ac": gids_pot_ac,
            "idx_etat": idx_etat,
            "idx_inf_type": idx_inf_type,
            "idx_inf_propri": idx_inf_propri,
            "idx_inf_mat": idx_inf_mat,
            "idx_dce": idx_dce,
            "idx_inf_num": idx_inf_num,
            "idx_commentaire": idx_commentaire,
            "idx_etiquette_jaune": idx_etiquette_jaune,
            "idx_etiquette_orange": idx_etiquette_orange,
            "idx_transition_aerosout": idx_transition_aerosout,
        }

    def _apply_update_ft_row(self, state, gid, row):
        infra_pt_pot = state["layer"]
        featFT = state["features_by_gid"].get(int(gid))
        if featFT is None:
            return False

        fid = featFT.id()
        action = str(row.get("action", "")).upper()

        idx_etat = state["idx_etat"]
        idx_inf_type = state["idx_inf_type"]
        idx_inf_propri = state["idx_inf_propri"]
        idx_inf_mat = state["idx_inf_mat"]
        idx_dce = state["idx_dce"]
        idx_commentaire = state["idx_commentaire"]
        idx_etiquette_jaune = state["idx_etiquette_jaune"]
        idx_etiquette_orange = state["idx_etiquette_orange"]
        idx_transition_aerosout = state["idx_transition_aerosout"]

        if action == "IMPLANTATION":
            ancien_inf_num = featFT["inf_num"] or ""
            if idx_etat >= 0:
                infra_pt_pot.changeAttributeValue(fid, idx_etat, "FT KO")
            if idx_inf_mat >= 0:
                new_inf_mat = row.get("inf_mat_replace") or featFT["inf_mat"] or ""
                infra_pt_pot.changeAttributeValue(fid, idx_inf_mat, new_inf_mat)
            if idx_inf_propri >= 0:
                infra_pt_pot.changeAttributeValue(fid, idx_inf_propri, "RAUV")
            if idx_inf_type >= 0:
                infra_pt_pot.changeAttributeValue(fid, idx_inf_type, "POT-AC")
                state["gids_pot_ac"].append(int(gid))
            if idx_commentaire >= 0:
                nouveau_commentaire = f"POT FT (ancien nommage : {ancien_inf_num} est FT KO)"
                infra_pt_pot.changeAttributeValue(fid, idx_commentaire, nouveau_commentaire)
            if idx_dce >= 0:
                infra_pt_pot.changeAttributeValue(fid, idx_dce, 'O')
        else:
            if row.get("etat"):
                infra_pt_pot.changeAttributeValue(fid, idx_etat, row["etat"])
            if row.get("inf_mat_replace"):
                infra_pt_pot.changeAttributeValue(fid, idx_inf_mat, row["inf_mat_replace"])
            if idx_dce >= 0:
                infra_pt_pot.changeAttributeValue(fid, idx_dce, 'O')

        if idx_etiquette_jaune >= 0 and row.get("etiquette_jaune"):
            infra_pt_pot.changeAttributeValue(fid, idx_etiquette_jaune, row["etiquette_jaune"])
        if idx_etiquette_orange >= 0 and row.get("etiquette_orange"):
            infra_pt_pot.changeAttributeValue(fid, idx_etiquette_orange, row["etiquette_orange"])

        if row.get("zone_privee") == 'X' and idx_commentaire >= 0:
            commentaire_actuel = featFT["commentair"]
            commentaire_str = str(commentaire_actuel) if commentaire_actuel and commentaire_actuel != NULL else ''
            if '/PRIVE' not in commentaire_str.upper():
                nouveau_commentaire = f"{commentaire_str}/PRIVE" if commentaire_str.strip() else "/PRIVE"
                infra_pt_pot.changeAttributeValue(fid, idx_commentaire, nouveau_commentaire)
        
        # Transition aérosout: concaténer /AEROSOUTRANSI au commentaire
        if row.get("transition_aerosout") == 'oui' and idx_commentaire >= 0:
            commentaire_actuel = featFT["commentair"]
            commentaire_str = str(commentaire_actuel) if commentaire_actuel and commentaire_actuel != NULL else ''
            if '/AEROSOUTRANSI' not in commentaire_str.upper():
                nouveau_commentaire = f"{commentaire_str}/AEROSOUTRANSI" if commentaire_str.strip() else "/AEROSOUTRANSI"
                infra_pt_pot.changeAttributeValue(fid, idx_commentaire, nouveau_commentaire)

        if idx_transition_aerosout >= 0 and row.get("transition_aerosout"):
            infra_pt_pot.changeAttributeValue(fid, idx_transition_aerosout, row["transition_aerosout"])

        return True

    def _finalize_update_ft(self, state):
        infra_pt_pot = state["layer"]
        if not infra_pt_pot.commitChanges():
            errors = infra_pt_pot.commitErrors()
            err_detail = "; ".join(errors) if errors else "inconnu"
            raise RuntimeError(f"Commit échoué: {err_detail}")

        if state["gids_pot_ac"]:
            QgsMessageLog.logMessage(
                f"[MAJ-FT] Trigger PostgreSQL ({len(state['gids_pot_ac'])} implantations)...",
                "PoleAerien", Qgis.Info
            )
            self._execSqlTriggerBatch(infra_pt_pot, state["gids_pot_ac"])
            infra_pt_pot.triggerRepaint()

    def miseAjourFinalDesDonneesFT(self, table_poteau, liste_valeur_trouve_ft, progress_callback=None):
        """
        REQ-MAJ-007: MAJ FT dans infra_pt_pot avec gestion implantation.

        Args:
            progress_callback: callable(message, percent) pour feedback UI
        """
        t0_func = time.perf_counter()

        state = self._prepare_update_ft(table_poteau, liste_valeur_trouve_ft)
        infra_pt_pot = state["layer"]

        try:
            t3 = time.perf_counter()
            count = 0
            total = len(liste_valeur_trouve_ft)

            if progress_callback:
                progress_callback(f"MAJ FT: 0/{total} poteaux...", 60)

            for gid, row in liste_valeur_trouve_ft.iterrows():
                t_feat_start = time.perf_counter()
                count += 1
                action = str(row.get("action", "")).upper()

                if count % 5 == 0 or count == 1:
                    pct = 60 + int((count / total) * 25)  # 60% -> 85%
                    if progress_callback:
                        progress_callback(f"MAJ FT: {count}/{total} ({action})...", pct)
                    QApplication.processEvents()

                if not self._apply_update_ft_row(state, gid, row):
                    continue

                t_feat_end = time.perf_counter()
                if count <= 5 or (t_feat_end - t_feat_start) > 1.0:
                    QgsMessageLog.logMessage(
                        f"[PERF-MAJ-BD] Feature {count}/{total} (gid={gid}): {t_feat_end - t_feat_start:.3f}s",
                        "PoleAerien", Qgis.Info
                    )

            t4 = time.perf_counter()
            temps_maj = t4 - t3

            if progress_callback:
                progress_callback(f"Enregistrement FT ({count} poteaux)...", 86)
            QApplication.processEvents()

            t5 = time.perf_counter()
            self._finalize_update_ft(state)
            t6 = time.perf_counter()

            temps_total = time.perf_counter() - t0_func
            QgsMessageLog.logMessage("=" * 50, "PoleAerien", Qgis.Info)
            QgsMessageLog.logMessage("[MAJ-FT] TERMINE", "PoleAerien", Qgis.Info)
            QgsMessageLog.logMessage(f"[MAJ-FT] Poteaux traites: {count}", "PoleAerien", Qgis.Info)
            QgsMessageLog.logMessage(f"[MAJ-FT] Temps total: {temps_total:.1f}s", "PoleAerien", Qgis.Info)
            QgsMessageLog.logMessage(f"[MAJ-FT] MAJ attributs: {temps_maj:.1f}s", "PoleAerien", Qgis.Info)
            QgsMessageLog.logMessage(f"[MAJ-FT] Commit BD: {t6-t5:.1f}s", "PoleAerien", Qgis.Info)
            QgsMessageLog.logMessage("=" * 50, "PoleAerien", Qgis.Info)

        except Exception as e:
            infra_pt_pot.rollBack()
            msg = f"[MAJ_FT_BT] miseAjourFinalDesDonneesFT échoué: {e}"
            QgsMessageLog.logMessage(msg, "MAJ_FT_BT", Qgis.Critical)
            raise RuntimeError(msg) from e

# ... (rest of the code remains the same)
    def _execSqlTriggerBatch(self, layer, gids):
        """Execute SQL UPDATE batch pour declencher trigger PostgreSQL.
        
        PERF-01: Connexion unique pour tous les gids (batch).
        
        Args:
            layer: QgsVectorLayer PostgreSQL
            gids: list[int] - Liste des gid a traiter
            
        Returns:
            int: Nombre de gids traites avec succes
        """
        if not gids:
            return 0
        
        # QA-04: Verifier provider PostgreSQL
        if layer.dataProvider().name() != "postgres":
            QgsMessageLog.logMessage("Skip trigger: layer non PostgreSQL", "MAJ_FT_BT", Qgis.Info)
            return 0
        
        uri = QgsDataSourceUri(layer.dataProvider().dataSourceUri())
        
        # QA-04: Verifier schema/table non vides
        if not uri.schema() or not uri.table():
            QgsMessageLog.logMessage("Erreur: schema ou table vide", "MAJ_FT_BT", Qgis.Warning)
            return 0
        
        conn_name = "trigger_conn_batch"
        
        # PERF-01: Connexion unique
        db = QSqlDatabase.addDatabase("QPSQL", conn_name)
        db.setHostName(uri.host())
        db.setPort(int(uri.port()) if uri.port() else 5432)
        db.setDatabaseName(uri.database())
        db.setUserName(uri.username())
        db.setPassword(uri.password())
        
        if not db.open():
            QgsMessageLog.logMessage(f"Erreur connexion DB: {db.lastError().text()}", "MAJ_FT_BT", Qgis.Critical)
            QSqlDatabase.removeDatabase(conn_name)
            return 0
        
        # PERF-01: Batch UPDATE avec IN clause
        # SEC-01: Validation identifiants SQL pour éviter injection
        schema = uri.schema().replace('"', '""')
        table = uri.table().replace('"', '""')
        gids_str = ",".join(str(int(g)) for g in gids)
        
        # Sauvegarder inf_num dans nommage_fibees AVANT de le vider
        sql_save = f'''UPDATE "{schema}"."{table}" 
                  SET nommage_fibees = inf_num 
                  WHERE gid IN ({gids_str}) AND inf_type = 'POT-AC' AND inf_num IS NOT NULL'''
        
        query_save = QSqlQuery(db)
        if not query_save.exec_(sql_save):
            QgsMessageLog.logMessage(f"Erreur sauvegarde nommage_fibees: {query_save.lastError().text()}", "MAJ_FT_BT", Qgis.Warning)
        
        # Puis vider inf_num
        sql = f'''UPDATE "{schema}"."{table}" 
                  SET inf_num = NULL 
                  WHERE gid IN ({gids_str}) AND inf_type = 'POT-AC' '''
        
        query = QSqlQuery(db)
        success_count = 0
        
        if query.exec_(sql):
            success_count = len(gids)
        else:
            QgsMessageLog.logMessage(f"Erreur SQL batch: {query.lastError().text()}", "MAJ_FT_BT", Qgis.Critical)
        
        db.close()
        QSqlDatabase.removeDatabase(conn_name)
        return success_count

    def _prepare_update_bt(self, table_poteau, liste_valeur_trouve_bt):
        lyrs = QgsProject.instance().mapLayersByName(table_poteau)
        if not lyrs or not lyrs[0].isValid():
            raise ValueError(f"Couche invalide: {table_poteau}")
        infra_pt_pot = lyrs[0]

        fields = infra_pt_pot.fields()
        required = {"etat", "inf_type", "inf_propri", "noe_usage", "inf_mat", "dce", "inf_num"}
        missing = [f for f in required if fields.indexOf(f) == -1]
        if missing:
            raise ValueError(f"Champs requis manquants: {missing}")

        idx_etat = fields.indexOf("etat")
        idx_inf_type = fields.indexOf("inf_type")
        idx_inf_propri = fields.indexOf("inf_propri")
        idx_noe_usage = fields.indexOf("noe_usage")
        idx_inf_mat = fields.indexOf("inf_mat")
        idx_dce = fields.indexOf("dce")
        idx_commentaire = fields.indexOf("commentair")
        idx_etiquette_jaune = fields.indexOf("etiquette_jaune")
        idx_etiquette_orange = fields.indexOf("etiquette_orange")

        gids_pot_ac = []

        t1 = time.perf_counter()
        gids_needed = set(int(g) for g in liste_valeur_trouve_bt.index)
        features_by_gid = {}
        for feat in infra_pt_pot.getFeatures():
            gid_val = feat["gid"]
            if gid_val in gids_needed:
                features_by_gid[gid_val] = feat
        t2 = time.perf_counter()
        QgsMessageLog.logMessage(
            f"[PERF-MAJ-BD] Index features BT: {len(features_by_gid)} en {t2-t1:.3f}s",
            "PoleAerien", Qgis.Info
        )

        if not infra_pt_pot.isEditable():
            infra_pt_pot.startEditing()

        return {
            "layer": infra_pt_pot,
            "features_by_gid": features_by_gid,
            "gids_pot_ac": gids_pot_ac,
            "idx_etat": idx_etat,
            "idx_inf_type": idx_inf_type,
            "idx_inf_propri": idx_inf_propri,
            "idx_noe_usage": idx_noe_usage,
            "idx_inf_mat": idx_inf_mat,
            "idx_dce": idx_dce,
            "idx_commentaire": idx_commentaire,
            "idx_etiquette_jaune": idx_etiquette_jaune,
            "idx_etiquette_orange": idx_etiquette_orange,
        }

    def _apply_update_bt_row(self, state, gid, row):
        infra_pt_pot = state["layer"]
        featBT = state["features_by_gid"].get(int(gid))
        if featBT is None:
            return False

        fid = featBT.id()
        idx_etat = state["idx_etat"]
        idx_inf_type = state["idx_inf_type"]
        idx_inf_propri = state["idx_inf_propri"]
        idx_noe_usage = state["idx_noe_usage"]
        idx_inf_mat = state["idx_inf_mat"]
        idx_dce = state["idx_dce"]
        idx_etiquette_orange = state["idx_etiquette_orange"]

        if str(row["Portée molle"]).upper() != "X":
            if row["inf_type"]:
                infra_pt_pot.changeAttributeValue(fid, idx_inf_type, row["inf_type"])
                if row["inf_type"] == "POT-AC":
                    state["gids_pot_ac"].append(int(gid))
            if row["inf_propri"]:
                infra_pt_pot.changeAttributeValue(fid, idx_inf_propri, row["inf_propri"])
            infra_pt_pot.changeAttributeValue(fid, idx_noe_usage, "DI")
            if row["typ_po_mod"]:
                infra_pt_pot.changeAttributeValue(fid, idx_inf_mat, row["typ_po_mod"])
            infra_pt_pot.changeAttributeValue(fid, idx_etat, row["etat"])
            infra_pt_pot.changeAttributeValue(fid, idx_dce, 'O')

            if idx_etiquette_orange >= 0 and row.get("etiquette_orange"):
                infra_pt_pot.changeAttributeValue(fid, idx_etiquette_orange, row["etiquette_orange"])
            
            # Zone privée BT: concaténer /PRIVE au commentaire
            idx_commentaire = state["idx_commentaire"]
            if row.get("zone_privee") == 'X' and idx_commentaire >= 0:
                commentaire_actuel = featBT["commentair"]
                commentaire_str = str(commentaire_actuel) if commentaire_actuel and commentaire_actuel != NULL else ''
                if '/PRIVE' not in commentaire_str.upper():
                    nouveau_commentaire = f"{commentaire_str}/PRIVE" if commentaire_str.strip() else "/PRIVE"
                    infra_pt_pot.changeAttributeValue(fid, idx_commentaire, nouveau_commentaire)
            return True

        infra_pt_pot.changeAttributeValue(fid, idx_etat, 'PORTEE MOLLE')

        if not featBT.hasGeometry():
            QgsMessageLog.logMessage(
                f"miseAjourFinalDesDonneesBT: geometrie absente (gid {gid})",
                "MAJ_FT_BT", Qgis.Warning
            )
            return False

        geom = featBT.geometry()
        if geom.isNull() or geom.isEmpty():
            QgsMessageLog.logMessage(
                f"miseAjourFinalDesDonneesBT: geometrie vide (gid {gid})",
                "MAJ_FT_BT", Qgis.Warning
            )
            return False

        if geom.isMultipart():
            QgsMessageLog.logMessage(
                f"miseAjourFinalDesDonneesBT: geometrie multipart (gid {gid})",
                "MAJ_FT_BT", Qgis.Warning
            )
            return False

        if QgsWkbTypes.geometryType(geom.wkbType()) != QgsWkbTypes.PointGeometry:
            QgsMessageLog.logMessage(
                f"miseAjourFinalDesDonneesBT: geometrie non point (gid {gid})",
                "MAJ_FT_BT", Qgis.Warning
            )
            return False

        point_geom = geom.asPoint()
        point_decale = QgsPointXY(point_geom.x() + 1, point_geom.y())

        feat = QgsFeature(featBT)
        feat.setGeometry(QgsGeometry.fromPointXY(point_decale))
        if row["inf_type"]:
            feat.setAttribute("inf_type", row["inf_type"])
        if row["inf_propri"]:
            feat.setAttribute("inf_propri", row["inf_propri"])
        feat.setAttribute("noe_usage", "DI")
        if row["typ_po_mod"]:
            feat.setAttribute("inf_mat", row["typ_po_mod"])
        feat.setAttribute("etat", row["etat"])
        feat.setAttribute("dce", 'O')
        infra_pt_pot.addFeature(feat)
        return True

    def _finalize_update_bt(self, state):
        infra_pt_pot = state["layer"]
        if not infra_pt_pot.commitChanges():
            errors = infra_pt_pot.commitErrors()
            err_detail = "; ".join(errors) if errors else "inconnu"
            raise RuntimeError(f"Commit BT échoué: {err_detail}")

        if state["gids_pot_ac"]:
            QgsMessageLog.logMessage(
                f"[MAJ-BT] Trigger PostgreSQL ({len(state['gids_pot_ac'])} POT-AC)...",
                "PoleAerien", Qgis.Info
            )
            self._execSqlTriggerBatch(infra_pt_pot, state["gids_pot_ac"])
            infra_pt_pot.triggerRepaint()

    def miseAjourFinalDesDonneesBT(self, table_poteau, liste_valeur_trouve_bt, progress_callback=None):
        """MAJ BT dans infra_pt_pot. Preserve attributs existants.
        
        Args:
            progress_callback: callable(message, percent) pour feedback UI
        """
        t0_func = time.perf_counter()
        
        lyrs = QgsProject.instance().mapLayersByName(table_poteau)
        if not lyrs or not lyrs[0].isValid():
            raise ValueError(f"Couche invalide: {table_poteau}")
        infra_pt_pot = lyrs[0]

        # QA-07: Index des champs avec verification
        fields = infra_pt_pot.fields()
        required = {"etat", "inf_type", "inf_propri", "noe_usage", "inf_mat", "dce", "inf_num"}
        missing = [f for f in required if fields.indexOf(f) == -1]
        if missing:
            raise ValueError(f"Champs requis manquants: {missing}")
        
        idx_etat = fields.indexOf("etat")
        idx_inf_type = fields.indexOf("inf_type")
        idx_inf_propri = fields.indexOf("inf_propri")
        idx_noe_usage = fields.indexOf("noe_usage")
        idx_inf_mat = fields.indexOf("inf_mat")
        idx_dce = fields.indexOf("dce")
        idx_commentaire = fields.indexOf("commentair")
        idx_etiquette_jaune = fields.indexOf("etiquette_jaune")
        idx_etiquette_orange = fields.indexOf("etiquette_orange")
        
        # Liste des gid POT-AC pour trigger SQL
        gids_pot_ac = []
        
        # PERF-OPTIM: Construire un index gid -> feature UNE SEULE FOIS
        t1 = time.perf_counter()
        gids_needed = set(int(g) for g in liste_valeur_trouve_bt.index)
        features_by_gid = {}
        for feat in infra_pt_pot.getFeatures():
            gid_val = feat["gid"]
            if gid_val in gids_needed:
                features_by_gid[gid_val] = feat
        t2 = time.perf_counter()
        QgsMessageLog.logMessage(f"[PERF-MAJ-BD] Index features BT: {len(features_by_gid)} en {t2-t1:.3f}s", "PoleAerien", Qgis.Info)
        
        # PERF-02: Demarrer edition une seule fois
        if not infra_pt_pot.isEditable():
            infra_pt_pot.startEditing()

        t3 = time.perf_counter()
        count = 0
        total = len(liste_valeur_trouve_bt)
        
        # Message initial
        if progress_callback:
            progress_callback(f"MAJ BT: 0/{total} poteaux...", 88)
        
        for gid, row in liste_valeur_trouve_bt.iterrows():
            count += 1
            
            # Mise a jour UI tous les 5 features
            if count % 5 == 0 or count == 1:
                pct = 88 + int((count / total) * 8) if total > 0 else 88  # 88% -> 96%
                if progress_callback:
                    progress_callback(f"MAJ BT: {count}/{total}...", pct)
                QApplication.processEvents()
            
            # PERF-OPTIM: Lookup direct O(1) au lieu de requête O(n)
            featBT = features_by_gid.get(int(gid))
            if featBT is None:
                continue
            
            fid = featBT.id()

            if str(row["Portée molle"]).upper() != "X":
                # Cas normal: modifier uniquement les attributs nécessaires
                if row["inf_type"]:
                    infra_pt_pot.changeAttributeValue(fid, idx_inf_type, row["inf_type"])
                    if row["inf_type"] == "POT-AC":
                        gids_pot_ac.append(int(gid))
                if row["inf_propri"]:
                    infra_pt_pot.changeAttributeValue(fid, idx_inf_propri, row["inf_propri"])
                infra_pt_pot.changeAttributeValue(fid, idx_noe_usage, "DI")
                if row["typ_po_mod"]:
                    infra_pt_pot.changeAttributeValue(fid, idx_inf_mat, row["typ_po_mod"])
                infra_pt_pot.changeAttributeValue(fid, idx_etat, row["etat"])
                infra_pt_pot.changeAttributeValue(fid, idx_dce, 'O')
                
                if idx_etiquette_orange >= 0 and row.get("etiquette_orange"):
                    infra_pt_pot.changeAttributeValue(fid, idx_etiquette_orange, row["etiquette_orange"])
                
                # Zone privée BT: concaténer /PRIVE au commentaire
                if row.get("zone_privee") == 'X' and idx_commentaire >= 0:
                    commentaire_actuel = featBT["commentair"]
                    commentaire_str = str(commentaire_actuel) if commentaire_actuel and commentaire_actuel != NULL else ''
                    if '/PRIVE' not in commentaire_str.upper():
                        nouveau_commentaire = f"{commentaire_str}/PRIVE" if commentaire_str.strip() else "/PRIVE"
                        infra_pt_pot.changeAttributeValue(fid, idx_commentaire, nouveau_commentaire)
            else:
                # Cas portée molle: marquer l'existant et créer nouveau décalé
                infra_pt_pot.changeAttributeValue(fid, idx_etat, 'PORTEE MOLLE')
                
                if not featBT.hasGeometry():
                    QgsMessageLog.logMessage(
                        f"miseAjourFinalDesDonneesBT: geometrie absente (gid {gid})",
                        "MAJ_FT_BT", Qgis.Warning
                    )
                    continue

                geom = featBT.geometry()
                if geom.isNull() or geom.isEmpty():
                    QgsMessageLog.logMessage(
                        f"miseAjourFinalDesDonneesBT: geometrie vide (gid {gid})",
                        "MAJ_FT_BT", Qgis.Warning
                    )
                    continue

                if geom.isMultipart():
                    QgsMessageLog.logMessage(
                        f"miseAjourFinalDesDonneesBT: geometrie multipart (gid {gid})",
                        "MAJ_FT_BT", Qgis.Warning
                    )
                    continue

                if QgsWkbTypes.geometryType(geom.wkbType()) != QgsWkbTypes.PointGeometry:
                    QgsMessageLog.logMessage(
                        f"miseAjourFinalDesDonneesBT: geometrie non point (gid {gid})",
                        "MAJ_FT_BT", Qgis.Warning
                    )
                    continue
                
                # Créer nouveau point décalé avec copie de tous les attributs
                point_geom = geom.asPoint()
                point_decale = QgsPointXY(point_geom.x() + 1, point_geom.y())
                
                feat = QgsFeature(featBT)
                feat.setGeometry(QgsGeometry.fromPointXY(point_decale))
                if row["inf_type"]:
                    feat.setAttribute("inf_type", row["inf_type"])
                if row["inf_propri"]:
                    feat.setAttribute("inf_propri", row["inf_propri"])
                feat.setAttribute("noe_usage", "DI")
                if row["typ_po_mod"]:
                    feat.setAttribute("inf_mat", row["typ_po_mod"])
                feat.setAttribute("etat", row["etat"])
                feat.setAttribute("dce", 'O')
                infra_pt_pot.addFeature(feat)
        
        t4 = time.perf_counter()
        temps_maj = t4 - t3
        
        # Message progression: enregistrement
        if progress_callback:
            progress_callback(f"Enregistrement BT ({count} poteaux)...", 96)
        QApplication.processEvents()

        # PERF-02: Commit unique a la fin
        t5 = time.perf_counter()
        if not infra_pt_pot.commitChanges():
            errors = infra_pt_pot.commitErrors()
            err_detail = "; ".join(errors) if errors else "inconnu"
            msg = f"Commit BT échoué: {err_detail}"
            infra_pt_pot.rollBack()
            QgsMessageLog.logMessage(msg, "MAJ_FT_BT", Qgis.Critical)
            raise RuntimeError(msg)
        t6 = time.perf_counter()

        # PERF-01: Declencher trigger PostgreSQL via batch SQL
        if gids_pot_ac:
            QgsMessageLog.logMessage(f"[MAJ-BT] Trigger PostgreSQL ({len(gids_pot_ac)} POT-AC)...", "PoleAerien", Qgis.Info)
            self._execSqlTriggerBatch(infra_pt_pot, gids_pot_ac)
            infra_pt_pot.triggerRepaint()
        
        # RÉCAP FINAL BT
        temps_total = time.perf_counter() - t0_func
        QgsMessageLog.logMessage("=" * 50, "PoleAerien", Qgis.Info)
        QgsMessageLog.logMessage("[MAJ-BT] TERMINE", "PoleAerien", Qgis.Info)
        QgsMessageLog.logMessage(f"[MAJ-BT] Poteaux traites: {count}", "PoleAerien", Qgis.Info)
        QgsMessageLog.logMessage(f"[MAJ-BT] Temps total: {temps_total:.1f}s", "PoleAerien", Qgis.Info)
        QgsMessageLog.logMessage(f"[MAJ-BT] MAJ attributs: {temps_maj:.1f}s", "PoleAerien", Qgis.Info)
        QgsMessageLog.logMessage(f"[MAJ-BT] Commit BD: {t6-t5:.1f}s", "PoleAerien", Qgis.Info)
        QgsMessageLog.logMessage("=" * 50, "PoleAerien", Qgis.Info)
