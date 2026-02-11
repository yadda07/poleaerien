#!/usr/bin/python
# -*- coding: utf-8 -*-

from qgis.PyQt.QtCore import QObject, pyqtSignal
from qgis.core import (
    Qgis, QgsFeatureRequest, QgsExpression, QgsProject,
    QgsTask, QgsApplication, QgsSpatialIndex, NULL,
    QgsMessageLog
)
import pandas as pd
import os
import re
import random
import difflib
from .qgis_utils import validate_same_crs, get_layer_safe
from .core_utils import normalize_appui_num
from .dataclasses_results import (
    ExcelValidationResult, PoteauxPolygoneResult, 
    EtudesValidationResult, ImplantationValidationResult
)

def _get_etat(action, prefix):
    """Determine etat poteau selon action Excel. prefix='FT' ou 'BT'."""
    action_lower = str(action).lower()
    if "implantation" in action_lower:
        return f"{prefix} KO"
    elif "recalage" in action_lower:
        return "A RECALER"
    elif "remplacement" in action_lower:
        return "A REMPLACER"
    elif "renforcement" in action_lower:
        return "A RENFORCER"
    return ""


def _get_action(action):
    """Normalise action Excel en label standardise."""
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
            maj = MajFtBt()
            self.signals.progress.emit(random.randint(3, 7))
            self.signals.message.emit("Lecture Excel...", "grey")

            # 1. Lecture Excel
            excel_ft, excel_bt, colonnes_warnings = maj.LectureFichiersExcelsFtBtKo(
                self.params['fichier_excel']
            )
            
            # Afficher les warnings de colonnes mal écrites dans le log du plugin
            for warn in colonnes_warnings:
                self.signals.message.emit(f"ATTENTION : {warn}", "orange")
            
            if self.isCanceled():
                return False

            self.signals.progress.emit(random.randint(12, 18))
            self.signals.message.emit("Analyse spatiale FT...", "grey")

            # 2. Traitement spatial PUR PYTHON (pas d'accès QGIS)
            bd_ft, bd_bt = self._process_spatial_pure_python()
            
            if self.isCanceled():
                return False

            self.signals.progress.emit(random.randint(35, 42))
            self.signals.message.emit("Comparaison données...", "grey")

            # 3. Comparaison
            liste_ft, liste_bt = maj.comparerLesDonnees(
                excel_ft, excel_bt, bd_ft, bd_bt
            )

            self.result = {
                'liste_ft': liste_ft,
                'liste_bt': liste_bt
            }
            
            self.signals.progress.emit(random.randint(48, 52))
            return True

        except Exception as e:
            self.exception = str(e)
            QgsMessageLog.logMessage(f"[MAJ] ERREUR Task: {e}", "PoleAerien", Qgis.Critical)
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
        data_ft = []
        for pot in poteaux_ft:
            x, y = pot['x'], pot['y']
            for etude in etudes_cap_ft:
                # Test bbox d'abord (rapide)
                if self._point_in_bbox(x, y, etude['bbox']):
                    # Test précis polygon
                    if self._point_in_polygon(x, y, etude['vertices']):
                        inf_num = pot['inf_num']
                        num_court = normalize_appui_num(inf_num)
                        data_ft.append({
                            'gid': pot['gid'],
                            'N° appui': num_court,
                            'inf_num': inf_num,
                            'Nom Etudes': etude['nom_etudes']
                        })
                        break  # Un poteau = une seule étude
        
        # Diagnostic: combien de correspondances spatiales FT trouvees?
        QgsMessageLog.logMessage(
            f"[MAJ_BD] Spatial FT: {len(poteaux_ft)} poteaux, {len(etudes_cap_ft)} polygones "
            f"-> {len(data_ft)} correspondances",
            "PoleAerien", Qgis.Info
        )
        if data_ft:
            sample = data_ft[0]
            QgsMessageLog.logMessage(
                f"[MAJ_BD] Exemple BD FT: appui='{sample['N° appui']}', etude='{sample['Nom Etudes']}'",
                "PoleAerien", Qgis.Info
            )
        elif poteaux_ft and etudes_cap_ft:
            QgsMessageLog.logMessage(
                "[MAJ_BD] ATTENTION: poteaux et polygones existent mais aucune correspondance spatiale. "
                "Verifier que les poteaux sont geometriquement DANS les polygones etude_cap_ft.",
                "PoleAerien", Qgis.Warning
            )

        if self.isCanceled():
            return pd.DataFrame(), pd.DataFrame()
        
        self.signals.progress.emit(random.randint(22, 28))
        self.signals.message.emit("Analyse spatiale BT...", "grey")
        
        # === BT: Point-in-polygon ===
        data_bt = []
        for pot in poteaux_bt:
            x, y = pot['x'], pot['y']
            for etude in etudes_comac:
                if self._point_in_bbox(x, y, etude['bbox']):
                    if self._point_in_polygon(x, y, etude['vertices']):
                        inf_num = pot['inf_num']
                        num_court = normalize_appui_num(inf_num)
                        data_bt.append({
                            'gid': pot['gid'],
                            'N° appui': num_court,
                            'inf_num': inf_num,
                            'Nom Etudes': etude['nom_etudes']
                        })
                        break
        
        # Diagnostic: combien de correspondances spatiales BT trouvees?
        QgsMessageLog.logMessage(
            f"[MAJ_BD] Spatial BT: {len(poteaux_bt)} poteaux, {len(etudes_comac)} polygones "
            f"-> {len(data_bt)} correspondances",
            "PoleAerien", Qgis.Info
        )

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

    def _detect_and_fix_columns(self, df, colonnes_attendues, nom_onglet):
        """
        Détecte les colonnes mal écrites et les corrige automatiquement.
        
        Pour chaque colonne attendue absente, cherche la colonne existante
        la plus similaire. Si trouvée, renomme la colonne et collecte un warning.
        Si non trouvée, la colonne reste manquante (erreur bloquante).
        
        Args:
            df: DataFrame pandas à corriger (modifié in-place)
            colonnes_attendues: list des noms de colonnes requises
            nom_onglet: 'FT' ou 'BT' (pour les messages)
            
        Returns:
            tuple: (manquantes: list[str], warnings: list[str])
        """
        colonnes_presentes_list = list(df.columns)
        manquantes = []
        warnings = []
        
        for col_attendue in colonnes_attendues:
            if col_attendue in df.columns:
                continue
            
            # Chercher la colonne la plus similaire (seuil 0.8 = 80% de similarité)
            similaires = difflib.get_close_matches(
                col_attendue, colonnes_presentes_list, n=1, cutoff=0.8
            )
            
            if similaires:
                ancien_nom = similaires[0]
                df.rename(columns={ancien_nom: col_attendue}, inplace=True)
                colonnes_presentes_list = list(df.columns)
                warnings.append(
                    f"Onglet {nom_onglet} : colonne '{ancien_nom}' "
                    f"corrigée en '{col_attendue}'"
                )
            else:
                manquantes.append(col_attendue)
        
        return manquantes, warnings

    def LectureFichiersExcelsFtBtKo(self, fichier_Excel):
        """Fonction pour parcourir les fichiers Excel pour renseigner la référence des appuis."""

        colonnes_warnings = []
        try:
            with pd.ExcelFile(fichier_Excel) as xls:
                # Lire avec header existant (ligne 0)
                df_ft = pd.read_excel(xls, "FT", header=0, index_col=None)
                df_bt = pd.read_excel(xls, "BT", header=0, index_col=None)
            
            # Normaliser noms colonnes (espaces multiples, trim)
            df_ft.columns = df_ft.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
            df_bt.columns = df_bt.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
            
            # Vérifier et corriger colonnes FT (auto-renomme si typo détectée)
            colonnes_warnings = []
            manquantes_ft, warns_ft = self._detect_and_fix_columns(
                df_ft, COLONNES_FT_REQUISES, "FT"
            )
            colonnes_warnings.extend(warns_ft)
            if manquantes_ft:
                raise ValueError(
                    f"Onglet FT — colonnes manquantes (aucune correspondance trouvée) :\n"
                    f"  {manquantes_ft}\n"
                    f"Colonnes présentes dans le fichier :\n  {list(df_ft.columns)}"
                )
            
            # Vérifier et corriger colonnes BT (auto-renomme si typo détectée)
            manquantes_bt, warns_bt = self._detect_and_fix_columns(
                df_bt, COLONNES_BT_REQUISES, "BT"
            )
            colonnes_warnings.extend(warns_bt)
            if manquantes_bt:
                raise ValueError(
                    f"Onglet BT — colonnes manquantes (aucune correspondance trouvée) :\n"
                    f"  {manquantes_bt}\n"
                    f"Colonnes présentes dans le fichier :\n  {list(df_bt.columns)}"
                )

            ######################################### FT #######################################
            # On donne la liste des colonnes que l'on souhaite garder
            df_ft = df_ft.loc[:, ["Nom Etudes", "N° appui", "Action", "inf_mat_replace", 
                                   "Etiquette jaune", "Zone privée", "Transition aérosout"]]

            # REQ-MAJ-007: Gestion FT Implantation
            df_ft['etat'] = df_ft['Action'].apply(lambda a: _get_etat(a, "FT"))
            df_ft['action'] = df_ft['Action'].apply(_get_action)
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
            df_bt['etat'] = df_bt['Action'].apply(lambda a: _get_etat(a, "BT"))
            df_bt['inf_type'] = df_bt['Action'].apply(lambda x: "POT-AC" if "implantation" in str(x).lower() else "")
            df_bt['action'] = df_bt['Action'].apply(_get_action)
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

        except ValueError:
            raise

        except AttributeError as lettre:
            QgsMessageLog.logMessage(f"FICHIER : {fichier_Excel} - {lettre}", "MAJ_FT_BT", Qgis.Warning)
            df_ft = pd.DataFrame({})
            df_bt = pd.DataFrame({})

        except Exception as e:
            QgsMessageLog.logMessage(f"Erreur fichier {fichier_Excel}: {e}", "MAJ_FT_BT", Qgis.Critical)
            df_ft = pd.DataFrame({})
            df_bt = pd.DataFrame({})

        # Debug logs removed for production

        return df_ft, df_bt, colonnes_warnings


    def comparerLesDonnees(self, excel_df_ft, excel_df_bt, bd_df_ft, bd_df_bt):
        """Compare les données Dataframe du fichier Excel par rapport à la base de données"""
        ######################################### FT #######################################
        liste_valeur_introuvbl_ft = pd.DataFrame({})
        liste_valeur_trouve_ft = pd.DataFrame({})

        excel_df_ft["N° appui"] = excel_df_ft["N° appui"].astype(str).str.lstrip("'").apply(normalize_appui_num)
        bd_df_ft["N° appui"] = bd_df_ft["N° appui"].astype(str).apply(normalize_appui_num)
        excel_df_ft["Nom Etudes"] = excel_df_ft["Nom Etudes"].astype(str).str.strip()
        if not bd_df_ft.empty:
            bd_df_ft["Nom Etudes"] = bd_df_ft["Nom Etudes"].astype(str).str.strip()

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

        # Diagnostic: afficher tailles et echantillons avant merge
        QgsMessageLog.logMessage(
            f"[MAJ_BD] Pre-merge FT: Excel={len(excel_df_ft_valid)} lignes, BD={len(bd_df_ft_valid)} lignes",
            "PoleAerien", Qgis.Info
        )
        if not excel_df_ft_valid.empty:
            etudes_xls = sorted(excel_df_ft_valid["Nom Etudes"].unique())[:5]
            appuis_xls = sorted(excel_df_ft_valid["N° appui"].unique())[:5]
            QgsMessageLog.logMessage(
                f"[MAJ_BD] Excel FT etudes: {etudes_xls}", "PoleAerien", Qgis.Info)
            QgsMessageLog.logMessage(
                f"[MAJ_BD] Excel FT appuis: {appuis_xls}", "PoleAerien", Qgis.Info)
        if not bd_df_ft_valid.empty:
            etudes_bd = sorted(bd_df_ft_valid["Nom Etudes"].unique())[:5]
            appuis_bd = sorted(bd_df_ft_valid["N° appui"].unique())[:5]
            QgsMessageLog.logMessage(
                f"[MAJ_BD] BD FT etudes: {etudes_bd}", "PoleAerien", Qgis.Info)
            QgsMessageLog.logMessage(
                f"[MAJ_BD] BD FT appuis: {appuis_bd}", "PoleAerien", Qgis.Info)
        else:
            QgsMessageLog.logMessage(
                "[MAJ_BD] BD FT VIDE: aucun poteau FT dans polygones etude_cap_ft. "
                "Verifier couches et zone geographique.",
                "PoleAerien", Qgis.Warning
            )

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

        excel_df_bt["N° appui"] = excel_df_bt["N° appui"].astype(str).str.lstrip("'").apply(normalize_appui_num)
        bd_df_bt["N° appui"] = bd_df_bt["N° appui"].astype(str).apply(normalize_appui_num)
        excel_df_bt["Nom Etudes"] = excel_df_bt["Nom Etudes"].astype(str).str.strip()
        if not bd_df_bt.empty:
            bd_df_bt["Nom Etudes"] = bd_df_bt["Nom Etudes"].astype(str).str.strip()

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

        # Diagnostic: afficher tailles et echantillons avant merge BT
        QgsMessageLog.logMessage(
            f"[MAJ_BD] Pre-merge BT: Excel={len(excel_df_bt_valid)} lignes, BD={len(bd_df_bt_valid)} lignes",
            "PoleAerien", Qgis.Info
        )
        if bd_df_bt_valid.empty and not excel_df_bt_valid.empty:
            QgsMessageLog.logMessage(
                "[MAJ_BD] BD BT VIDE: aucun poteau BT dans polygones etude_comac. "
                "Verifier couches et zone geographique.",
                "PoleAerien", Qgis.Warning
            )

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

    @staticmethod
    def exporter_rapport_maj(liste_ft, liste_bt, export_dir):
        """Exporte un rapport Excel detaillant les modifications MAJ BD.

        Args:
            liste_ft: [nb_introuvables, df_introuvables, nb_trouves, df_trouves]
            liste_bt: idem pour BT
            export_dir: repertoire d'export

        Returns:
            str: chemin du fichier Excel genere
        """
        from datetime import datetime

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"rapport_maj_bd_{ts}.xlsx"
        filepath = os.path.join(export_dir, filename)

        tt_introuvable_ft = liste_ft[0] if len(liste_ft) > 0 else 0
        df_introuvable_ft = liste_ft[1] if len(liste_ft) > 1 else pd.DataFrame()
        tt_trouve_ft = liste_ft[2] if len(liste_ft) > 2 else 0
        df_trouve_ft = liste_ft[3] if len(liste_ft) > 3 else pd.DataFrame()

        tt_introuvable_bt = liste_bt[0] if len(liste_bt) > 0 else 0
        df_introuvable_bt = liste_bt[1] if len(liste_bt) > 1 else pd.DataFrame()
        tt_trouve_bt = liste_bt[2] if len(liste_bt) > 2 else 0
        df_trouve_bt = liste_bt[3] if len(liste_bt) > 3 else pd.DataFrame()

        resume = pd.DataFrame([
            {"Element": "FT identifies", "Nombre": tt_trouve_ft},
            {"Element": "FT introuvables (Excel sans QGIS)", "Nombre": tt_introuvable_ft},
            {"Element": "BT identifies", "Nombre": tt_trouve_bt},
            {"Element": "BT introuvables (Excel sans QGIS)", "Nombre": tt_introuvable_bt},
        ])

        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            resume.to_excel(writer, sheet_name='RESUME', index=False)

            if isinstance(df_trouve_ft, pd.DataFrame) and not df_trouve_ft.empty:
                df_trouve_ft.to_excel(writer, sheet_name='FT_MAJ')
            else:
                pd.DataFrame({"Info": ["Aucun FT identifie"]}).to_excel(
                    writer, sheet_name='FT_MAJ', index=False)

            if isinstance(df_trouve_bt, pd.DataFrame) and not df_trouve_bt.empty:
                df_trouve_bt.to_excel(writer, sheet_name='BT_MAJ')
            else:
                pd.DataFrame({"Info": ["Aucun BT identifie"]}).to_excel(
                    writer, sheet_name='BT_MAJ', index=False)

            if isinstance(df_introuvable_ft, pd.DataFrame) and not df_introuvable_ft.empty:
                df_introuvable_ft.to_excel(writer, sheet_name='FT_INTROUVABLES', index=False)

            if isinstance(df_introuvable_bt, pd.DataFrame) and not df_introuvable_bt.empty:
                df_introuvable_bt.to_excel(writer, sheet_name='BT_INTROUVABLES', index=False)

        return filepath
