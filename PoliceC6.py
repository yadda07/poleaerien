#!/usr/bin/python
# -*-coding:Utf-8-*-

# Second pour contenir quelques fonctions qui seront utilisées dans le fichier principale Police.py
from qgis.core import (
    Qgis,
    QgsProject,
    QgsExpression,
    QgsFeatureRequest,
    QgsVectorLayer,
    QgsFeature,
    QgsGeometry,
    QgsPointXY,
    QgsCoordinateReferenceSystem,
    QgsMessageLog,
    QgsSpatialIndex,
    NULL,
    QgsRectangle
)
from qgis.PyQt.QtWidgets import QApplication
from .qgis_utils import (
    remove_group,
    layer_group_error,
    insert_layer_in_group,
    get_layer_safe,
    validate_same_crs,
    normalize_appui_num,
)
from .dataclasses_results import (
    PoliceC6Result, CableCapaciteResult, BoitierValidationResult,
    EtudeC6Result, ParcourAutoC6Result
)
from typing import List, Optional

import os
import re
import warnings
import openpyxl
import pandas as pd

LYR_INFRA_PT_POT = "infra_pt_pot"
LYR_INFRA_PT_CHB = "infra_pt_chb"
LYR_T_CHEMINEMENT_COPY = "t_cheminement_copy"
LYR_BPE = "bpe"
LYR_ATTACHES = "attaches"
STYLE_DIR = os.path.join(os.path.dirname(__file__), "styles")

# Import comac_db_reader - relative import for plugin, absolute for standalone testing
try:
    from .comac_db_reader import get_cable_capacite
except ImportError as e:
    try:
        from comac_db_reader import get_cable_capacite
    except ImportError:
        raise ImportError(f"comac_db_reader module not found: {e}") from e


class PoliceC6:
    """ Classe qui fait tout le travail pour accéder à la base de données,
    Réaliser des requêtes dans la base de données,
    Exporter le résultat des requêtes vers le fichier Excel """

    def __init__(self):
        """Initialise les attributs de classe."""
        self.fied_id_Ebp = "gid"
        self._reset_state()
    
    def _reset_state(self):
        """Réinitialise l'état pour une nouvelle analyse.
        
        Appelée automatiquement dans __init__ et doit être appelée
        avant chaque nouvelle analyse si l'instance est réutilisée.
        """
        self.nb_appui_corresp = 0
        self.nb_pbo_corresp = 0
        self.bpo_corresp = []
        self.nb_appui_absent = 0
        self.nb_appui_absentPot = 0
        self.potInfNumPresent = []
        self.infNumPotAbsent = []
        self.ebp_non_appui = []
        self.absence = []
        self.idPotPresent = []
        self.idPotAbsent = []
        self.bpe_pot_cap_ft = []
        self.ebp_appui_inconnu = []
        self.liste_appui_ebp = []
        self.presence_liste_appui_ebp = False
        self._pbo_a_supprimer = set()
        self._ebp_a_supprimer = set()
        self.listeCableAppuitrouve = []
        self.listeAppuiCapaAppuiAbsent = []

    def _norm_inf_num(self, value):
        """Normalise inf_num."""
        return normalize_appui_num(value)

    # =========================================================================
    # Wrappers UI/Projet - délèguent à utils.py
    # =========================================================================
    def removeGroup(self, name):
        """Suppression d'un groupe de couches."""
        remove_group(name)

    def layerGroupError(self, couche, nom_etude):
        """Ajoute une couche dans un groupe d'erreur."""
        layer_group_error(couche, nom_etude)

    def insertLayerInGroupGraceTHD(self, couche):
        """Ajoute une couche dans le groupe GRACETHD."""
        insert_layer_in_group(couche, "GRACETHD")


    # =========================================================================
    # REQ-PLC6-003: Parcours automatique C6
    # =========================================================================
    def parcourir_etudes_auto(self, table_etude_cap_ft, colonne_nom_etude: str,
                               repertoire_c6: str = None,
                               colonne_chemin_c6: str = None) -> ParcourAutoC6Result:
        """
        REQ-PLC6-003: Parcourt automatiquement études CAP_FT et traite C6 associés.
        
        QA-EXPERT: Validation fichiers, gestion erreurs par étude.
        PERF-SPECIALIST: Traitement séquentiel avec feedback.
        
        Args:
            table_etude_cap_ft: Couche etude_cap_ft
            colonne_nom_etude: Nom champ contenant nom étude
            repertoire_c6: Répertoire base des fichiers C6
            colonne_chemin_c6: Nom champ contenant chemin C6 (optionnel)
            
        Returns:
            ParcourAutoC6Result avec résultats par étude
        """
        result = ParcourAutoC6Result()
        
        # QA: Récupération sécurisée couche
        try:
            etude_cap_ft = get_layer_safe(table_etude_cap_ft, "POLICE_C6")
        except ValueError as e:
            QgsMessageLog.logMessage(
                f"[POLICE_C6] parcourir_etudes_auto: {e}",
                "PoleAerien", Qgis.Critical
            )
            return result
        
        # QA: Vérifier champ nom étude existe
        idx_nom = etude_cap_ft.fields().indexFromName(colonne_nom_etude)
        if idx_nom == -1:
            QgsMessageLog.logMessage(
                f"[POLICE_C6] Champ '{colonne_nom_etude}' absent de etude_cap_ft",
                "PoleAerien", Qgis.Critical
            )
            return result
        
        # Optionnel: champ chemin C6
        idx_chemin = -1
        if colonne_chemin_c6:
            idx_chemin = etude_cap_ft.fields().indexFromName(colonne_chemin_c6)
        
        for feat in etude_cap_ft.getFeatures():
            nom_etude = feat[colonne_nom_etude]
            if not nom_etude or nom_etude == NULL:
                continue
            
            nom_etude_str = str(nom_etude).strip()
            if not nom_etude_str:
                continue
            
            etude_result = EtudeC6Result(etude=nom_etude_str)
            
            # Déterminer chemin C6
            chemin_c6 = None
            if idx_chemin >= 0 and feat[colonne_chemin_c6]:
                chemin_c6 = str(feat[colonne_chemin_c6])
            elif repertoire_c6:
                # Convention: repertoire/nom_etude/Annexe_C6.xlsx
                chemin_c6 = os.path.join(repertoire_c6, nom_etude_str, "Annexe_C6.xlsx")
                if not os.path.exists(chemin_c6):
                    # Essayer autre pattern
                    chemin_c6 = os.path.join(repertoire_c6, f"{nom_etude_str}.xlsx")
            
            etude_result.chemin_c6 = chemin_c6 or ""
            
            # QA: Vérifier fichier existe
            if not chemin_c6 or not os.path.exists(chemin_c6):
                etude_result.statut = "ERREUR"
                etude_result.erreur = f"Fichier C6 introuvable: {chemin_c6}"
                QgsMessageLog.logMessage(
                    f"[POLICE_C6] Fichier C6 introuvable pour étude {nom_etude_str}: {chemin_c6}",
                    "PoleAerien", Qgis.Warning
                )
                result.etudes_traitees.append(etude_result)
                continue
            
            # Traiter C6
            try:
                champs_xlsx, liste_cable = self.lectureFichierExcel(chemin_c6)
                etude_result.statut = "OK"
                # Stocker résumé dans resultat
                etude_result.resultat = PoliceC6Result(
                    nb_appui_corresp=len(champs_xlsx),
                    liste_cable_appui_trouve=liste_cable
                )
            except Exception as e:
                etude_result.statut = "ERREUR"
                etude_result.erreur = str(e)
                QgsMessageLog.logMessage(
                    f"[POLICE_C6] Erreur traitement {nom_etude_str}: {e}",
                    "PoleAerien", Qgis.Critical
                )
            
            result.etudes_traitees.append(etude_result)
            QApplication.processEvents()  # Keep UI responsive
        
        return result

    # =========================================================================
    # REQ-PLC6-006: Parsing câbles COMAC
    # =========================================================================
    def parse_cables_comac(self, colonne_ao: str) -> List[str]:
        """
        REQ-PLC6-006: Parse colonne AO COMAC pour extraire références câbles.
        
        Format: 'L1092-13-P-L1092-12-P-' → ['L1092-13-P', 'L1092-12-P']
        
        QA-EXPERT: Gestion null, formats variés.
        
        Args:
            colonne_ao: Chaîne format COMAC
            
        Returns:
            Liste codes câbles
        """
        if not colonne_ao or str(colonne_ao).strip() in ('', '-', 'nan', 'None'):
            return []
        
        colonne_ao = str(colonne_ao).strip()
        
        # Split par '-' et reconstituer codes (format: XXX-YY-Z)
        parts = colonne_ao.strip('-').split('-')
        cables = []
        
        # Chaque câble a 3 parties: code-num-type
        i = 0
        while i + 2 < len(parts):
            code = f"{parts[i]}-{parts[i+1]}-{parts[i+2]}"
            if code and code != '--':
                cables.append(code)
            i += 3
        
        # Gérer cas où dernier câble incomplet
        if i < len(parts) and len(parts) - i >= 2:
            code = f"{parts[i]}-{parts[i+1]}"
            if len(parts) > i + 2:
                code += f"-{parts[i+2]}"
            if code and code != '-':
                cables.append(code)
        
        return cables

    # =========================================================================
    # REQ-PLC6-006: Vérification capacité câbles
    # =========================================================================
    def verifier_capacite_cables(self, df_comac, col_ao: str = 'AO') -> CableCapaciteResult:
        """
        REQ-PLC6-006: Vérifie cohérence capacité câbles vs type ligne.
        
        QA-EXPERT: Validation null, gestion erreurs par câble.
        
        Args:
            df_comac: DataFrame avec colonnes AO
            col_ao: Nom colonne références câbles
            
        Returns:
            CableCapaciteResult
        """
        result = CableCapaciteResult()
        
        if df_comac is None or df_comac.empty:
            return result
        
        # Détecter colonne AO (case insensitive)
        col_ao_found = None
        col_appui = None
        for col in df_comac.columns:
            col_lower = str(col).lower().strip()
            if col_lower == col_ao.lower() or 'ao' in col_lower:
                col_ao_found = col
            if 'appui' in col_lower:
                col_appui = col
        
        if col_ao_found is None:
            return result
        
        for idx, row in df_comac.iterrows():
            colonne_ao = row.get(col_ao_found, '')
            num_appui = row.get(col_appui, '') if col_appui else f"ligne_{idx}"
            
            # Parser câbles
            cables = self.parse_cables_comac(colonne_ao)
            
            if not cables:
                continue
            
            result.cables_traites += len(cables)
            
            # Récupérer capacités
            for code_cable in cables:
                try:
                    cap = get_cable_capacite(code_cable)
                    if cap and cap > 0:
                        result.cables_valides += 1
                    else:
                        result.anomalies.append({
                            'appui': str(num_appui),
                            'cable': code_cable,
                            'capacite': 0,
                            'erreur': "Capacité nulle ou non trouvée"
                        })
                except Exception as e:
                    result.anomalies.append({
                        'appui': str(num_appui),
                        'cable': code_cable,
                        'erreur': f"Câble inconnu: {e}"
                    })
        
        return result

    # =========================================================================
    # REQ-PLC6-007: Vérification boîtiers
    # =========================================================================
    def verifier_boitiers(self, df_comac) -> BoitierValidationResult:
        """
        REQ-PLC6-007: Vérifie cohérence boîtiers.
        
        QA-EXPERT: Validation null, format boîtier.
        
        Args:
            df_comac: DataFrame avec colonne Boitier
            
        Returns:
            BoitierValidationResult
        """
        result = BoitierValidationResult()
        
        if df_comac is None or df_comac.empty:
            return result
        
        # Détecter colonne Boitier (case insensitive)
        col_boitier_found = None
        col_appui = None
        for col in df_comac.columns:
            col_lower = str(col).lower().strip()
            if 'boitier' in col_lower or 'boite' in col_lower:
                col_boitier_found = col
            if 'appui' in col_lower:
                col_appui = col
        
        if col_boitier_found is None:
            return result
        
        # Types boîtiers valides (PB, PBO, PEO, PBR, PEP)
        BOITIERS_VALIDES = ['PB', 'PBO', 'PEO', 'PBR', 'PEP', 'BPE', 'BPEO']
        
        for idx, row in df_comac.iterrows():
            boitier = row.get(col_boitier_found, '')
            num_appui = row.get(col_appui, '') if col_appui else f"ligne_{idx}"
            
            result.boitiers_traites += 1
            
            # QA: Vérifier boîtier renseigné
            if pd.isna(boitier) or str(boitier).strip() == '':
                result.anomalies.append({
                    'appui': str(num_appui),
                    'boitier': '',
                    'erreur': 'Boîtier non renseigné'
                })
                continue
            
            boitier_str = str(boitier).strip().upper()
            
            # Vérifier format valide
            if boitier_str not in BOITIERS_VALIDES:
                # Vérifier si commence par un type valide
                valide = False
                for bv in BOITIERS_VALIDES:
                    if boitier_str.startswith(bv):
                        valide = True
                        break
                
                if not valide:
                    result.anomalies.append({
                        'appui': str(num_appui),
                        'boitier': boitier_str,
                        'erreur': f"Type boîtier '{boitier_str}' invalide. Attendus: {BOITIERS_VALIDES}"
                    })
                    continue
            
            result.boitiers_valides += 1
        
        return result

    def lectureFichierExcel(self, fname):
        """Fonction qui permet de lire le contenu du fichier Excel et de renvoyer ses contenus (format .xlsx uniquement)."""

        # Chargement du fichier Excel
        try:
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
                document = openpyxl.load_workbook(fname, data_only=True)
        except Exception:
            return [], []

        try:
            feuille_1 = document.worksheets[3]
        except IndexError:
            return [], []

        champs_xlsx = []
        liste_cable_appui_OD = []
        self.liste_appui_ebp = []
        num_appui_o = ''

        for r_idx, row in enumerate(feuille_1.iter_rows(min_row=9), start=9):  # A8 = ligne 8 → index Excel 9
            num_appui_start = str(row[0].value).strip().split("/")[0]  if row[0].value else ''
            capa_cable = str(row[18].value).strip() if len(row) > 8 and row[18].value else ''
            num_appui_end = str(row[24].value).strip().split("/")[0]  if len(row) > 24 and row[24].value else ''
            pbo = str(row[35].value).strip() if len(row) > 35 and row[35].value else ''

            # Nettoyage numéro d'appui
            num_appui_start = re.sub(r'\.0$', '', num_appui_start)
            num_appui_start = re.sub(r'^0+', '', num_appui_start)

            num_appui_origine = num_appui_start if num_appui_start else num_appui_o

            if num_appui_start:
                champs_xlsx.append(num_appui_start)
            # # Si appui non vide
            # if num_appui_start and not num_appui_start.lower():
            #     if not re.match(r'E', num_appui_start) and not re.match(r'IMB', num_appui_start):
            #         champs_xlsx.append(num_appui_start)

            # Enregistre dernier appui non vide
            if num_appui_origine:
                num_appui_o = num_appui_origine

            # Câble entre appuis (hors FAC)
            if not re.match(r'IMB', num_appui_o) and not re.match(r'IMB', num_appui_end):
                if len(capa_cable.replace(' ', '')) >= 8:
                    # Récupération capacité depuis BD officielle comac_db
                    extrait_capa_cab = get_cable_capacite(capa_cable)
                    if extrait_capa_cab > 0:
                        liste_cable_appui_OD.append([r_idx, num_appui_o, extrait_capa_cab, num_appui_end, pbo])

            # Boîtes (PB, PEO)
            if pbo.upper() in ["PB", "PEO", "PBR", "PEP", "PBO"]:
                self.liste_appui_ebp.append([r_idx, num_appui_o, pbo])

        self.presence_liste_appui_ebp = bool(self.liste_appui_ebp)
        # print('champs_xlsx: ', champs_xlsx)
        # print('liste_appui_ebp: ', self.liste_appui_ebp)
        return champs_xlsx, liste_cable_appui_OD




    def lireFichiers(self, fname, table, colonne, valeur, t_bpe, t_attaches, zone_layer_name: Optional[str] = None):
        """Fonction pour parcourir les fichiers Excel pour renseigner la référence des appuis."""
        # Reset état pour nouvelle analyse
        self._reset_state()
        
        # CRIT-04: Validation fichier
        if not fname or not os.path.exists(fname):
            QgsMessageLog.logMessage(
                f"PoliceC6.lireFichiers: fichier manquant ({fname})",
                "POLICE_C6", Qgis.Critical
            )
            return [], []

        try:
            bpe = get_layer_safe(t_bpe, "Police_C6")
            attaches = get_layer_safe(t_attaches, "Police_C6")
        except ValueError as e:
            QgsMessageLog.logMessage(
                f"PoliceC6.lireFichiers: couche manquante ({e})",
                "POLICE_C6", Qgis.Critical
            )
            return [], []

        self.nb_appui_corresp = 0
        self.nb_pbo_corresp = 0
        self.bpo_corresp = []
        self.nb_appui_absent = 0
        self.nb_appui_absentPot = 0
        self.potInfNumPresent = []
        self.infNumPotAbsent = []
        infNumPoteauAbsent = []
        self.ebp_non_appui = []
        valeursmanquant = []
        self.absence = []

        champs_xlsx, liste_cable_appui_OD = self.lectureFichierExcel(fname)
        if not champs_xlsx and not liste_cable_appui_OD:
            QgsMessageLog.logMessage(
                "PoliceC6.lireFichiers: lecture Excel vide",
                "POLICE_C6", Qgis.Critical
            )
            return [], []

        colonne_safe = str(colonne).replace('"', '')
        valeur_safe = str(valeur).replace("'", "''")
        requete = QgsExpression(f'"{colonne_safe}" LIKE \'{valeur_safe}\'')

        self.idPotPresent = []
        self.idPotAbsent = []

        infra_pt_pot = ""
        infra_pt_chb = ""
        etude_cap_ft = ""
        bpe_pot = []
        t_cheminement_copy = ""

        self.bpe_pot_cap_ft = []
        self.ebp_appui_inconnu = []
        self._pbo_a_supprimer = set()
        self._ebp_a_supprimer = set()

        bufferDist = 0.5

        # Récupération sécurisée des couches via get_layer_safe
        try:
            infra_pt_pot = get_layer_safe(LYR_INFRA_PT_POT, "Police_C6")
            infra_pt_chb = get_layer_safe(LYR_INFRA_PT_CHB, "Police_C6")
            etude_cap_ft = get_layer_safe(table, "Police_C6")
            t_cheminement_copy = get_layer_safe(LYR_T_CHEMINEMENT_COPY, "Police_C6")
        except ValueError as e:
            QgsMessageLog.logMessage(
                f"PoliceC6.lireFichiers: {e}",
                "POLICE_C6", Qgis.Critical
            )
            return [], []

        zone_layer = None
        zone_index = None
        zone_cache = None
        if zone_layer_name:
            try:
                zone_layer = get_layer_safe(zone_layer_name, "Police_C6")
            except ValueError as e:
                QgsMessageLog.logMessage(
                    f"PoliceC6.lireFichiers: zone invalide ({e})",
                    "POLICE_C6", Qgis.Warning
                )
                zone_layer = None

        # Validation CRS - toutes les couches doivent avoir le même CRS
        try:
            validate_same_crs(infra_pt_pot, etude_cap_ft, "Police_C6")
            validate_same_crs(infra_pt_pot, bpe, "Police_C6")
            validate_same_crs(infra_pt_pot, attaches, "Police_C6")
            validate_same_crs(infra_pt_pot, infra_pt_chb, "Police_C6")
            validate_same_crs(infra_pt_pot, t_cheminement_copy, "Police_C6")
            if zone_layer is not None:
                validate_same_crs(infra_pt_pot, zone_layer, "Police_C6")
        except ValueError as e:
            QgsMessageLog.logMessage(str(e), "POLICE_C6", Qgis.Critical)
            return [], []

        if zone_layer is not None:
            zone_feats = [f for f in zone_layer.getFeatures() if f.hasGeometry()]
            zone_index = QgsSpatialIndex()
            for f in zone_feats:
                zone_index.addFeature(f)
            zone_cache = {f.id(): f for f in zone_feats}

        field_index = bpe.fields().indexFromName("gid")
        self.fied_id_Ebp = "id" if field_index == -1 else "gid"

        infra_pt_pot.selectByIds([])
        etude_cap_ft.selectByIds([])

        idx_inf_num = infra_pt_pot.dataProvider().fields().indexFromName('inf_num')

        etude_req = QgsFeatureRequest(requete)
        etude_feats = [feat for feat in etude_cap_ft.getFeatures(etude_req) if feat.hasGeometry()]
        if zone_index is not None:
            etude_filtered = []
            for feat in etude_feats:
                geom = feat.geometry()
                if geom is None:
                    continue
                ids = zone_index.intersects(geom.boundingBox())
                in_zone = False
                for fid in ids:
                    z = zone_cache.get(fid) if zone_cache else None
                    if z and z.hasGeometry() and geom.intersects(z.geometry()):
                        in_zone = True
                        break
                if in_zone:
                    etude_filtered.append(feat)
            etude_feats = etude_filtered
        etude_index = QgsSpatialIndex()
        for feat in etude_feats:
            etude_index.addFeature(feat)
        etude_cache = {feat.id(): feat for feat in etude_feats}

        # CRIT-001: Calculer bbox zone d'étude pour filtrer features
        etude_bbox = QgsRectangle()
        for feat in etude_feats:
            etude_bbox.combineExtentWith(feat.geometry().boundingBox())
        
        # Étendre bbox de 10% pour captures adjacents
        buffer = max(etude_bbox.width(), etude_bbox.height()) * 0.1
        etude_bbox.grow(buffer)
        
        # Build spatial indexes avec filtre spatial (PERF: évite chargement 50k+ features)
        from .qgis_utils import build_spatial_index
        req_spatial = QgsFeatureRequest().setFilterRect(etude_bbox)
        
        pot_index, pot_cache = build_spatial_index(infra_pt_pot, req_spatial)
        bpe_index, bpe_cache = build_spatial_index(bpe, req_spatial)
        chb_index, chb_cache = build_spatial_index(infra_pt_chb, req_spatial)
        att_index, att_cache = build_spatial_index(attaches, req_spatial)
        
        op_count = 0  # Operation counter for processEvents

        # Pour les appuis adjacents, hors de la zone d'étude
        for idx_chem, feat_t_chem in enumerate(t_cheminement_copy.getFeatures()):
            if idx_chem % 10 == 0:
                QApplication.processEvents()  # Keep UI responsive
            adjacent = False  # Par défaut pas d'adjacent de trouvé

            if not feat_t_chem.hasGeometry():
                continue

            tchem_geom = feat_t_chem.geometry()
            tchem_bbox = tchem_geom.boundingBox()
            etude_candidates = etude_index.intersects(tchem_bbox)

            for fid in etude_candidates:
                feat_cap_ft = etude_cache.get(fid)
                if not feat_cap_ft:
                    continue
                # Le cheminement doit intersecter la zone d'étude ...
                if tchem_geom.intersects(feat_cap_ft.geometry()):

                    if tchem_geom.contains(feat_cap_ft.geometry()):
                        pass

                    # ... mais il ne doit pas être contenu dans la zone d'étude
                    else:
                        geom = tchem_geom.asMultiPolyline()  # .asMultiPolyline  ou  asPolyline
                        start_point = QgsPointXY(geom[0][0])  # l'origine du cable
                        chem_geom_start_point = QgsGeometry.fromPointXY(start_point).buffer(bufferDist, 0)

                        # Si
                        if chem_geom_start_point.intersects(feat_cap_ft.geometry()):
                            pass

                        else:
                            # Use spatial index instead of full scan
                            pot_candidates = pot_index.intersects(chem_geom_start_point.boundingBox())
                            for pot_id in pot_candidates:
                                feat_pot_origine = pot_cache.get(pot_id)
                                if not feat_pot_origine or not feat_pot_origine.geometry():
                                    continue
                                
                                op_count += 1
                                if op_count % 100 == 0:
                                    QApplication.processEvents()

                                if feat_pot_origine.geometry().intersects(chem_geom_start_point):
                                    if feat_pot_origine[idx_inf_num] != NULL:
                                        chaine = self._norm_inf_num(feat_pot_origine[idx_inf_num])
                                        if not chaine:
                                            continue
                                        valeursmanquant.append(chaine)

                                        if chaine in champs_xlsx and chaine not in self.potInfNumPresent:
                                            self.nb_appui_corresp += 1
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
                                # Use spatial index instead of full scan
                                pot_dest_candidates = pot_index.intersects(chem_geom_end_point.boundingBox())
                                for pot_dest_id in pot_dest_candidates:
                                    feat_pot_destination = pot_cache.get(pot_dest_id)
                                    if not feat_pot_destination or not feat_pot_destination.geometry():
                                        continue
                                    
                                    op_count += 1
                                    if op_count % 100 == 0:
                                        QApplication.processEvents()

                                    if feat_pot_destination.geometry().intersects(chem_geom_end_point):
                                        if feat_pot_destination[idx_inf_num] != NULL and re.match('POT', feat_pot_destination[idx_inf_num]) is None:
                                            chaine = self._norm_inf_num(feat_pot_destination[idx_inf_num])
                                            if not chaine:
                                                continue

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
        for feat_cap_ft in etude_feats:
            QApplication.processEvents()  # Keep UI responsive
            cands = infra_pt_pot.getFeatures(QgsFeatureRequest().setFilterRect(feat_cap_ft.geometry().boundingBox()))

            for idx_pot, feat_pot in enumerate(cands):
                if idx_pot % 20 == 0:
                    QApplication.processEvents()
                # Si infra_pt_pot intersecte avec la table polygones
                if feat_pot.geometry().intersects(feat_cap_ft.geometry()):

                    # Je parcours ligne par ligne la colonne inf_num qui ne contient pas 'POT' et dont inf_num non vide
                    if feat_pot[idx_inf_num] != NULL and 'E' not in feat_pot[idx_inf_num]:

                        chaine = self._norm_inf_num(feat_pot[idx_inf_num])
                        if not chaine:
                            continue

                        # if len(chaine)>= 12:
                        # regexp = re.compile(r'\-FT\-')
                        # if regexp.search(chaine):
                            # position = chaine.find('FT-') + 3
                            # chaine = chaine[position:]
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
                        # Use spatial index: get BPE candidates near feat_pot
                        pot_bbox = feat_pot.geometry().boundingBox()
                        pot_bbox.grow(bufferDist * 2)  # Extend for buffer
                        bpe_candidates = bpe_index.intersects(pot_bbox)
                        
                        for bpe_id in bpe_candidates:
                            feat_bpe = bpe_cache.get(bpe_id)
                            if not feat_bpe or not feat_bpe.geometry():
                                continue
                            
                            op_count += 1
                            if op_count % 100 == 0:
                                QApplication.processEvents()

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
                                            chaine_norm = self._norm_inf_num(chaine)
                                            if valeur[1] == chaine_norm:
                                                self.nb_pbo_corresp += 1
                                                self.bpo_corresp.append(valeur)
                                                
                                                # Marquer pour suppression (évite modif pendant itération)
                                                self._pbo_a_supprimer.add(compte)

                                                # Normaliser chaine pour comparaison
                                                chaine_norm = self._norm_inf_num(chaine)
                                                
                                                # Marquer ebp_appui_inconnu pour suppression
                                                for idx, item in enumerate(self.ebp_appui_inconnu):
                                                    inf_num_norm = str(item[1]).split("/")[0]
                                                    if chaine_norm in inf_num_norm:
                                                        self._ebp_a_supprimer.add(idx)
                                                        break
                                                break

                                else:
                                    # Use spatial index for attaches
                                    bpe_bbox = feat_bpe.geometry().boundingBox()
                                    bpe_bbox.grow(bufferDist * 2)
                                    att_candidates = att_index.intersects(bpe_bbox)
                                    
                                    for att_id in att_candidates:
                                        feat_attaches = att_cache.get(att_id)
                                        if not feat_attaches or not feat_attaches.geometry():
                                            continue
                                        
                                        geom_attaches = feat_attaches.geometry().asPolyline()
                                        if not geom_attaches:
                                            continue
                                        
                                        att_start_point = QgsPointXY(geom_attaches[0])
                                        att_geom_start_point = QgsGeometry.fromPointXY(att_start_point).buffer(bufferDist, 0)
                                        att_end_point = QgsPointXY(geom_attaches[-1])
                                        att_geom_end_point = QgsGeometry.fromPointXY(att_end_point).buffer(bufferDist, 0)

                                        if feat_bpe.geometry().intersects(att_geom_start_point) or feat_bpe.geometry().intersects(att_geom_end_point):
                                            if feat_pot.geometry().intersects(att_geom_start_point) or feat_pot.geometry().intersects(att_geom_end_point):
                                                self.ebp_appui_inconnu.append([feat_pot.id(), feat_pot[idx_inf_num], feat_bpe['noe_type']])
                                                self.bpe_pot_cap_ft.append(feat_bpe.id())

                                                # Match with Excel data using set marking (no del during iteration)
                                                if self.liste_appui_ebp:
                                                    chaine_norm = self._norm_inf_num(chaine)
                                                    for compte, valeur in enumerate(self.liste_appui_ebp):
                                                        if valeur[1] == chaine_norm:
                                                            self.nb_pbo_corresp += 1
                                                            self.bpo_corresp.append(valeur)
                                                            self._pbo_a_supprimer.add(compte)
                                                            
                                                            for idx, item in enumerate(self.ebp_appui_inconnu):
                                                                if str(item[1]).split("/")[0].lstrip("0") == chaine_norm:
                                                                    self._ebp_a_supprimer.add(idx)
                                                                    break
                                                            break
                                                break

        # EBP qui n'intersecte pas d'appuis. EBP qui sont seuls, chose qui ne devrait pas exister.

        for gid_bpe in bpe_pot:
            if gid_bpe not in self.bpe_pot_cap_ft and gid_bpe not in self.ebp_non_appui:
                self.ebp_non_appui.append(gid_bpe)
        # print('bpe_pot: ', bpe_pot)
        # print('ebp_non_appui: ', self.ebp_non_appui)

        # Si EBP pas sur un appui, vérifie s'il n'intersecte pas une chambre
        if self.ebp_non_appui:
            bufferEbpChambre = 3
            ebp_a_retirer = set()  # Avoid modifying list during iteration
            
            condition2 = tuple(self.ebp_non_appui) if len(self.ebp_non_appui) > 1 else f"({self.ebp_non_appui[0]})"
            requeteEBP = QgsExpression(f"{self.fied_id_Ebp} IN {condition2}")

            for feat_bpe in bpe.getFeatures(QgsFeatureRequest(requeteEBP)):
                if not feat_bpe.geometry():
                    continue
                    
                op_count += 1
                if op_count % 50 == 0:
                    QApplication.processEvents()

                bpe_geom = feat_bpe.geometry()
                bpe_bbox = bpe_geom.boundingBox()
                bpe_bbox.grow(bufferEbpChambre * 2)
                etude_candidates = etude_index.intersects(bpe_bbox)

                for fid in etude_candidates:
                    feat_cap_ft = etude_cache.get(fid)
                    if not feat_cap_ft:
                        continue
                    # Use spatial index for chambres
                    chb_candidates = chb_index.intersects(bpe_bbox)
                    
                    for chb_id in chb_candidates:
                        feat_chb = chb_cache.get(chb_id)
                        if not feat_chb or not feat_chb.geometry():
                            continue
                        
                        if feat_chb.geometry().intersects(feat_cap_ft.geometry()):
                            buffer_feat_chb = feat_chb.geometry().buffer(bufferEbpChambre, 0)
                            if buffer_feat_chb.intersects(feat_bpe.geometry()):
                                ebp_a_retirer.add(feat_bpe.id())
                                break
            
            # Remove marked items after iteration
            self.ebp_non_appui = [x for x in self.ebp_non_appui if x not in ebp_a_retirer]

        # On vérifie si l'une des données existantes dans les fichiers Excels n'existent pas dans infra_pt_pot
        for absence in champs_xlsx:
            if absence not in valeursmanquant:
                self.nb_appui_absentPot += 1
                # Save numéros des appuis absents dans C3A
                # idée : abord, vérifier que les valeurs absence dans le fichier sont des entiers. Erreur connu "NoneType"
                self.absence.append(absence)

        infra_pt_pot.select(self.idPotPresent)

        # CRIT-003: Cleanup mémoire explicite (évite fuite 500MB)
        try:
            pot_cache.clear()
            bpe_cache.clear()
            chb_cache.clear()
            att_cache.clear()
            etude_cache.clear()
            if zone_cache:
                zone_cache.clear()
        except:
            pass
        
        return liste_cable_appui_OD, infNumPoteauAbsent

    def appliquerstyle(self, nomcouche):
        """Fonction qui permet d'appliquer un style au fichier erreur qui sera généré """
        # On applique un style au fichier C3A qui aurait été généré
        cheminstyle = ''
        style_map = {
            r'error_infra_pt_pot_': "infra_pt_pot.qml",
            r'error_infra_pt_pot_ebp': "infra_pt_pot_ebp.qml",
            r'error_appui_capa_appui': "t_cable.qml",
            r'error_bpe': "bpe.qml",
        }
        for pattern, filename in style_map.items():
            if re.match(pattern, nomcouche):
                cheminstyle = os.path.join(STYLE_DIR, filename)
                break

        verifstyle = True
        self.messagestyle = ''
        # On vérifie si le fichier contenant le style existe pour l'appliquer au fichier C3A

        if verifstyle:
            # Utiliser get_layer_safe au lieu de QgsProject.instance()
            try:
                couche = get_layer_safe(nomcouche, "Police_C6")
            except ValueError:
                # Couche non trouvée, skip
                return
            
            if couche:
                if not cheminstyle or not os.path.exists(cheminstyle):
                    verifstyle = False
                    self.messagestyle = u"Le style n'a pu être trouvé. A faire manuellement"
                    QgsMessageLog.logMessage(
                        "PoliceC6.appliquerstyle: style introuvable",
                        "POLICE_C6",
                        Qgis.Warning,
                    )
                    return
                try:
                    couche.loadNamedStyle(cheminstyle)
                    self.messagestyle = u"Le style a bien été appliquée au fichier C3A"
                except (OSError, RuntimeError) as e:
                    verifstyle = False
                    self.messagestyle = u"Le style n'a pu être appliqué. A faire manuellement"
                    QgsMessageLog.logMessage(
                        f"PoliceC6.appliquerstyle: {e}",
                        "POLICE_C6",
                        Qgis.Warning,
                    )

    def verificationnsDonnees(self):
        """Fonction qui vérifie que dans QGIS toutes les tables nécessaires existent déjà.
        Sinon, le programme ne s'éxecutera pas"""
        # Temps de démarrage
        liste_data_present = []  # Lister des données déjà présentes dans QGIS

        liste_import_data = [LYR_BPE, LYR_INFRA_PT_POT, LYR_INFRA_PT_CHB, LYR_ATTACHES]
        
        # Vérifier chaque couche attendue individuellement
        for layer_name in liste_import_data:
            try:
                get_layer_safe(layer_name, "Police_C6")
                liste_data_present.append(layer_name.lower())
            except ValueError:
                # Couche non trouvée, sera dans liste_absent
                pass

        # Test si parmi les données attendues, certaines sont absent.
        # Stocker des valeurs absentes dans QGIS
        liste_absent = [data for data in liste_import_data if data not in liste_data_present]

        return liste_absent

    def analyseAppuiCableAppui(self, liste_cable_appui_OD, t_etude_cap_ft, champs,  valeur, zone_layer_name: Optional[str] = None):
        """Fonction principale pour l'analyse des relations appuis - cables - appuis dans QGIS et dans Annexe C6 """
        cable_corresp = 0
        # self.non_cable_corresp = 0
        # self.id_non_cable_corresp = []  # Pour enregistrer les cables et leurs appuis qui n'ont pas trouvé de correspondance

        self.listeCableAppuitrouve = []
        self.listeAppuiCapaAppuiAbsent = []  # Liste des appuis et leurs capa qui sont absent d'Annexe C6

        try:
            t_chem = get_layer_safe(LYR_T_CHEMINEMENT_COPY, "Police_C6")
            etude_cap_ft = get_layer_safe(t_etude_cap_ft, "Police_C6")
        except ValueError as e:
            QgsMessageLog.logMessage(
                f"PoliceC6.analyseAppuiCableAppui: couche manquante ({e})",
                "POLICE_C6", Qgis.Critical
            )
            return 0, 0

        zone_layer = None
        zone_index = None
        zone_cache = None
        if zone_layer_name:
            try:
                zone_layer = get_layer_safe(zone_layer_name, "Police_C6")
            except ValueError as e:
                QgsMessageLog.logMessage(
                    f"PoliceC6.analyseAppuiCableAppui: zone invalide ({e})",
                    "POLICE_C6", Qgis.Warning
                )
                zone_layer = None

        try:
            if zone_layer is not None:
                validate_same_crs(t_chem, zone_layer, "Police_C6")
        except ValueError as e:
            QgsMessageLog.logMessage(str(e), "POLICE_C6", Qgis.Critical)
            return 0, 0

        if zone_layer is not None:
            zone_feats = [f for f in zone_layer.getFeatures() if f.hasGeometry()]
            zone_index = QgsSpatialIndex()
            for f in zone_feats:
                zone_index.addFeature(f)
            zone_cache = {f.id(): f for f in zone_feats}

        dicoPointsTechniquesCompletes = self.listePointTechniquesPoteaux("t_ptech_copy", "t_noeud_copy", "t_sitetech_copy")
        dicoCableLine = self.listeInfoCablesLines("t_cable_copy", "t_cableline_copy")

        if not dicoPointsTechniquesCompletes or not dicoCableLine:
            QgsMessageLog.logMessage(
                "PoliceC6.analyseAppuiCableAppui: données GraceTHD incomplètes",
                "POLICE_C6", Qgis.Critical
            )
            return 0, 0

        id_cable_chem_trouve = []  # Liste id des correspondances trouvés
        bufferDistExtremite = 0.5

        # CRIT-006: Fix injection SQL - utiliser double quotes pour noms colonnes
        requete = QgsExpression(f'"{champs}" = \'{valeur}\'')

        t_chem_feats = [feat for feat in t_chem.getFeatures() if feat.hasGeometry()]
        t_chem_index = QgsSpatialIndex()
        for feat in t_chem_feats:
            t_chem_index.addFeature(feat)
        t_chem_cache = {feat.id(): feat for feat in t_chem_feats}

        etude_req = QgsFeatureRequest(requete)
        etude_feats = [feat for feat in etude_cap_ft.getFeatures(etude_req) if feat.hasGeometry()]
        etude_index = QgsSpatialIndex()
        for feat in etude_feats:
            etude_index.addFeature(feat)
        etude_cache = {feat.id(): feat for feat in etude_feats}

        epsg = t_chem.crs().postgisSrid()  # SRID
        uri = (f"LineString?crs=epsg:{epsg}&field=id_cable_chem:string&field=appui_start:string&field=cab_capa:integer&"
               f"field=appui_end:string&field=cb_typelog:string&field=erreur:string&&index=yes")
        new_table_appui_capa_appui = QgsVectorLayer(uri, 'error_appui_capa_appui', 'memory')
        feat = QgsFeature()
        pr = new_table_appui_capa_appui.dataProvider()

        # self.barreP(35)

        # Parcourir d'abord la table des t_cable (GraceTHD) :
        # for feat_cb_line in t_cableline.getFeatures():
        for idx_cable, [_, cl_cb_code, cb_typelog, cb_capafo, feat_cb_line_geom] in enumerate(dicoCableLine.values()):
            if idx_cable % 5 == 0:
                QApplication.processEvents()  # Keep UI responsive

            if not feat_cb_line_geom:
                continue

            # Parcourir la table des etude_cap_ft en filtrant les données
            etude_ids = etude_index.intersects(feat_cb_line_geom.boundingBox())
            for fid in etude_ids:
                feat_etude_cap_ft = etude_cache.get(fid)
                if not feat_etude_cap_ft:
                    continue
                # On prend les géométries des t_cableline qui intersectent la zone d'étude "etude_cap_ft"
                if feat_cb_line_geom.intersects(feat_etude_cap_ft.geometry()):

                    # Créer un buffer de 2 mètres autour de la table t_cableline
                    buffer_t_cableline = feat_cb_line_geom.buffer(bufferDistExtremite, 0)

                    # Parcourir  la table des t_cheminement
                    chem_ids = t_chem_index.intersects(buffer_t_cableline.boundingBox())
                    for chem_id in chem_ids:
                        feat_t_chem = t_chem_cache.get(chem_id)
                        if not feat_t_chem:
                            continue

                        chem_geom = feat_t_chem.geometry()
                        # Le buffer de la géométrie doit contenir la table t_cheminement, pour être valide
                        if buffer_t_cableline.contains(chem_geom):

                            # Le t_cheminement doit intersecter la zone d'étude
                            # On prend les géométries des t_cheminement qui intersecte la zone d'étude
                            if chem_geom.intersects(feat_etude_cap_ft.geometry()):

                                inf_num_o, pt_typephy_o = dicoPointsTechniquesCompletes.get(
                                    feat_t_chem['cm_ndcode1'],
                                    ["", ""]
                                )
                                # print(f"inf_num_o : {inf_num_o} pt_nd_code : {pt_nd_code} pt_typephy : {pt_typephy} ")
                                # print(f"inf_num_o : {inf_num_o}")
                                inf_num_d, pt_typephy_d = dicoPointsTechniquesCompletes.get(
                                    feat_t_chem['cm_ndcode2'],
                                    ["", ""]
                                )

                                # Les chambres ne sont prise en compte dans les C6, dont on ne les comparent pas.
                                if pt_typephy_o != "C" and pt_typephy_d != "C":
                                    # Exclure les câbles entre immeubles (IMB)
                                    if "IMB" not in inf_num_o and "IMB" not in inf_num_d:
                                        id_cable_chem = f"{cl_cb_code} {feat_t_chem['cm_code']}"

                                        # MAJ de la nouvelle géométrie qui doit être créer
                                        feat.setAttributes([id_cable_chem, inf_num_o, cb_capafo, inf_num_d, cb_typelog, u"Introuvable dans C6"])
                                        feat.setGeometry(feat_t_chem.geometry())
                                        pr.addFeatures([feat])

        # self.barreP(40)

        # Liste des appuis et de leurs capa qui ont été trouvé dans la zone d'étude
        if new_table_appui_capa_appui.featureCount() > 0:
            existant = []

            for idx_c6, [ligne, origC6, capaC6, destC6, _] in enumerate(liste_cable_appui_OD):
                if idx_c6 % 5 == 0:
                    QApplication.processEvents()  # Keep UI responsive
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
                    
                    # Normaliser pour comparaison
                    origC6_norm = str(origC6).strip().lstrip("0")
                    destC6_norm = str(destC6).strip().lstrip("0")
                    pt_orig_norm = pt_orig.strip().lstrip("0")
                    pt_dest_norm = pt_dest.strip().lstrip("0")

                    # # Si le champs inf_num contient FT, on le prend que les valeurs situées après FT-
                    # regexp = re.compile(r'\-FT\-')
                    # if regexp.search(pt_orig):
                    #     position = pt_orig.find('FT-') + 3
                    #     pt_orig = pt_orig[position:]
                    #
                    # if regexp.search(pt_dest):
                    #     position = pt_dest.find('FT-') + 3
                    #     pt_dest = pt_dest[position:]

                    regexp_bt = re.compile(r'^E\d+')
                    # Les BT sont pris en charges uniquement si cela concerne uniquement une des deux extremités
                    # du câble.
                    # Si l'extremité d'origine contient 'BT'
                    if regexp_bt.match(str(origC6)) and not regexp_bt.match(str(destC6)):
                        #  on vérifie que les l'autre extrémité et capacité du cable correspondent aux données QGIS
                        if (capaC6 == capa and destC6_norm == pt_orig_norm) or (capaC6 == capa and destC6_norm == pt_dest_norm):

                            id_cable_chem_trouve.append(ide)

                            cable_corresp += 1
                            # Sauvegarde les valeurs qui correspondent dans les deux fichiers QGIS et (Annexe C6)
                            self.listeCableAppuitrouve.append([ligne, origC6, capaC6, destC6])
                            existant.append([ligne, origC6, capaC6, destC6])
                            break

                    # Si l'extremité de destination contient 'BT'
                    elif not regexp_bt.match(str(origC6)) and regexp_bt.match(str(destC6)):
                        #  on vérifie que les l'autre extrémité et capacité du cable correspondent aux données QGIS
                        if (capaC6 == capa and origC6_norm == pt_orig_norm) or (capaC6 == capa and origC6_norm == pt_dest_norm):

                            id_cable_chem_trouve.append(ide)

                            cable_corresp += 1
                            # Sauvegarde les valeurs qui correspondent dans les deux fichiers QGIS et (Annexe C6)
                            self.listeCableAppuitrouve.append([ligne, origC6, capaC6, destC6])
                            existant.append([ligne, origC6, capaC6, destC6])
                            break

                    # Si les trois vs du fichier Annexe C6 correspondant au cable
                    # et ses deux intersections d'appuis
                    elif ((origC6_norm == pt_orig_norm and capaC6 == capa and destC6_norm == pt_dest_norm) or
                          (destC6_norm == pt_orig_norm and capaC6 == capa and origC6_norm == pt_dest_norm)):
                        id_cable_chem_trouve.append(ide)

                        # myIndex.append(i)

                        cable_corresp += 1
                        # cm_code_corresp.append(contenu[3])
                        # Sauvegarde les valeurs qui correspondent dans les deux fichiers QGIS et (Annexe C6)
                        self.listeCableAppuitrouve.append([ligne, origC6, capaC6, destC6])
                        existant.append([ligne, origC6, capaC6, destC6])
                        break

            # for feat_resultat in new_table_appui_capa_appui.getFeatures():
            #     for i, v in enumerate(liste_cable_appui_OD):
            #         print("Qgis: feat_resultat['appui_start']: {}, feat_resultat['cab_capa']: {}, feat_resultat['appui_end']: {}".format(feat_resultat['appui_start'], feat_resultat['cab_capa'], feat_resultat['appui_end']))
            #         print("C6: feat_resultat['appui_start']: {}, feat_resultat['cab_capa']: {}, feat_resultat['appui_end']: {}".format(v[1], v[2], v[3]))
            #         if (
            #                 (feat_resultat['appui_start'] == v[1] and feat_resultat['cab_capa'] == v[2] and
            #                  feat_resultat['appui_end'] == v[3])
            #                 or
            #                 (feat_resultat['appui_end'] == v[1] and feat_resultat['cab_capa'] == v[2] and feat_resultat[
            #                     'appui_start'] == v[3])
            #         ):
            #             existant.append(liste_cable_appui_OD[i])
            #             print('correspondance')
            #         else:
            #             print('pas de correspondance')

            # Suppressions des correspondances trouvées dans Annexe C6
            for _, b in enumerate(existant):
                for i, v in enumerate(liste_cable_appui_OD):
                    if v[1] == b[1] and v[2] == b[2] and v[3] == b[3]:
                        del liste_cable_appui_OD[i]
                        break

            # Suppression des correspondances trouvées dans QGIS (batch)
            if id_cable_chem_trouve:
                ids_to_delete = [
                    feat.id()
                    for feat in new_table_appui_capa_appui.getFeatures()
                    if feat['id_cable_chem'] in id_cable_chem_trouve
                ]
                if ids_to_delete:
                    new_table_appui_capa_appui.dataProvider().deleteFeatures(ids_to_delete)

            # Nettoyage complémentaire (BT-BT et correspondances restantes)
            ids_to_delete = []
            for feat_resultat in new_table_appui_capa_appui.getFeatures():
                for i, v in enumerate(liste_cable_appui_OD):
                    q1 = feat_resultat['appui_start'].strip().lstrip("0")
                    q2 = str(feat_resultat['cab_capa']).strip()
                    q3 = feat_resultat['appui_end'].strip().lstrip("0")
                    c1 = str(v[1]).strip().lstrip("0")
                    c2 = str(v[2]).strip()
                    c3 = str(v[3]).strip().lstrip("0")
                    regexp_bt = re.compile(r'^E\d+')

                    if regexp_bt.match(q1) and regexp_bt.match(q3):
                        ids_to_delete.append(feat_resultat.id())
                        break

                    if (
                            (q1 == c1 and q2 == c2.strip() and q3 == c3)
                            or
                            (q1 == c3 and q2 == c2.strip() and q3 == c1)
                    ):
                        existant.append(liste_cable_appui_OD[i])
                        ids_to_delete.append(feat_resultat.id())
                        break

            if ids_to_delete:
                new_table_appui_capa_appui.dataProvider().deleteFeatures(ids_to_delete)

            for _, b in enumerate(existant):
                for i, v in enumerate(liste_cable_appui_OD):
                    if v[1] == b[1] and v[2] == b[2] and v[3] == b[3]:
                        del liste_cable_appui_OD[i]
                        break




        # self.barreP(50)

        # Enregistrer les id du t_cheminement qui sont dans la zone d'étude
        nbre_EntiteLigne = new_table_appui_capa_appui.featureCount()

        # Si le nombre d'étité est 0, on supprimer la couche, sinon on l'a ajoute dans le projet
        if nbre_EntiteLigne == 0:
            # Couche vide, pas besoin de l'ajouter au projet
            pass

        else:
            # Liste des appuis et leurs capa sont dans QGIS mais pas dans Annexe C6
            for feat_res in new_table_appui_capa_appui.getFeatures():
                pt_orig = feat_res['appui_start']
                capa = int(feat_res['cab_capa'])
                pt_dest = feat_res['appui_end']
                self.listeAppuiCapaAppuiAbsent.append([pt_orig, capa, pt_dest])

            # Ajout de la carte error_appui_capa_appui dans QGIS
            # NOTE: Cette couche doit être ajoutée par le workflow, pas par la logique métier
            # Pour l'instant, on stocke la référence pour que le workflow l'ajoute
            self._error_layer_to_add = new_table_appui_capa_appui

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
        crs = QgsCoordinateReferenceSystem("EPSG:2154")
        self.coucheShp.setCrs(crs)

        # NOTE: L'ajout au projet doit être fait par le workflow
        # On retourne la couche pour que le workflow l'ajoute
        return self.coucheShp

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
                            # NOTE: L'ajout au projet doit être fait par le workflow
                            # On stocke les couches créées pour que le workflow les ajoute
                            if not hasattr(self, '_layers_to_add'):
                                self._layers_to_add = []
                            self._layers_to_add.append(layer)
                            liste_import_data.append(f"{name}")
                            break

        liste_absent = [data for data in liste_import_data if data not in listeTableCsv]
        return liste_absent

    def listeInfoCablesLines(self, nomTableCable, nomTablet_CableLines):
        """Fonction qui permet de récupérer l'ensemble des informations liées aux câbles"""

        try:
            t_cableline = get_layer_safe(nomTablet_CableLines, "Police_C6")
            t_cable = get_layer_safe(nomTableCable, "Police_C6")
        except ValueError as e:
            QgsMessageLog.logMessage(
                f"PoliceC6.listeInfoCablesLines: couche manquante ({e})",
                "POLICE_C6", Qgis.Critical
            )
            return {}

        # Filtrer uniquement les cb_etiquet qui commencent par "TR"
        requete = QgsExpression("substr(\"cb_etiquet\", 1, 2) = 'TR%'")
        request = QgsFeatureRequest(requete)

        # Optionnel : tri sur le champ cb_etiquet
        clause = QgsFeatureRequest.OrderByClause('cb_etiquet', ascending=True)
        orderby = QgsFeatureRequest.OrderBy([clause])
        request.setOrderBy(orderby)

        dictionnaireCableLine = {}

        for feat_cable in t_cable.getFeatures():
            cb_code = feat_cable["cb_code"]
            if cb_code and cb_code != NULL:
                # print("cb_code :", cb_code)
                # cb_nd1 = feat_cable["cb_nd1"]  # Le noeud origine du câble
                # cb_nd2 = feat_cable["cb_nd2"]  # Le noeud destination du câble
                # t_cable[gid] = [cb_code, cb_etiquet, cb_nd1, cb_nd2, cb_capafo]

                ########################## T_CABLELINE ##############################################
                # Récupérer la longueur associée à chaque cable
                for feat_cableline in t_cableline.getFeatures():

                    # print('feat_cableline["cl_cb_code"]: ',feat_cableline["cl_cb_code"])
                    if cb_code == feat_cableline["cl_cb_code"]:
                        cb_etiquet = feat_cable["cb_codeext"]
                        cb_typelog = feat_cable["cb_typelog"]

                        raw_capafo = feat_cable["cb_capafo"]
                        if raw_capafo and raw_capafo != NULL:
                            capafo_int = int(raw_capafo)
                            if capafo_int == 36:
                                cb_capafo = 24
                            elif capafo_int == 72:
                                cb_capafo = 48
                            else:
                                cb_capafo = capafo_int
                        else:
                            cb_capafo = 0
                        cl_cb_code = feat_cableline["cl_cb_code"]
                        cb_geom = feat_cableline.geometry()
                        dictionnaireCableLine[cb_code] = [cb_etiquet, cl_cb_code, cb_typelog, cb_capafo, cb_geom]
                        break

        return dictionnaireCableLine

    def listePointTechniquesPoteaux(self, nomTablet_ptech, nomTablet_noeud, nomTablett_sitetech):
        ################################ POINTS TECHNIQUES ###############################################
        dicoPointsTechniquesCompletes = {}
        
        # Charger infra_pt_pot pour obtenir les vrais inf_num
        dico_infra_inf_num = {}  # noe_codext ou pot_codext -> inf_num
        try:
            infra = get_layer_safe("infra_pt_pot", "Police_C6")
            for feat_infra in infra.getFeatures():
                noe_codext = feat_infra["noe_codext"]
                pot_codext = feat_infra["pot_codext"]
                inf_num_raw = feat_infra["inf_num"]
                if inf_num_raw:
                    inf_num = str(inf_num_raw).split("/")[0] if "/" in str(inf_num_raw) else str(inf_num_raw)
                    if noe_codext:
                        dico_infra_inf_num[str(noe_codext)] = inf_num
                    if pot_codext:
                        dico_infra_inf_num[str(pot_codext)] = inf_num
        except ValueError:
            # infra_pt_pot non disponible, continuer sans mapping
            pass

        try:
            t_ptech = get_layer_safe(nomTablet_ptech, "Police_C6")
            t_noeud = get_layer_safe(nomTablet_noeud, "Police_C6")
        except ValueError as e:
            QgsMessageLog.logMessage(
                f"PoliceC6.listePointTechniquesPoteaux: couche manquante ({e})",
                "POLICE_C6", Qgis.Critical
            )
            return {}

        for feat_ptech in t_ptech.getFeatures():
            pt_nd_code = feat_ptech["pt_nd_code"]
            pt_typephy = feat_ptech["pt_typephy"] if feat_ptech["pt_typephy"] != NULL else ""

            for feat_noeud in t_noeud.getFeatures():
                if feat_ptech["pt_nd_code"] == feat_noeud["nd_code"]:
                    # Utiliser nd_codeext pour jointure avec infra_pt_pot
                    nd_codeext = feat_noeud["nd_codeext"] if feat_noeud["nd_codeext"] != NULL else ""
                    inf_num = dico_infra_inf_num.get(str(nd_codeext), "")
                    
                    dicoPointsTechniquesCompletes[pt_nd_code] = [inf_num, pt_typephy]
                    break

            # print('listeValeurs_T_ptech',listeValeurs_T_ptech)

        ############################# T_SITETECH #############################################
        try:
            t_sitetech_origine = get_layer_safe(nomTablett_sitetech, "Police_C6")
        except ValueError as e:
            QgsMessageLog.logMessage(
                f"PoliceC6.listePointTechniquesPoteaux: couche manquante ({e})",
                "POLICE_C6", Qgis.Critical
            )
            return dicoPointsTechniquesCompletes

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
        # print('dicoPointsTechniquesCompletes: ', dicoPointsTechniquesCompletes)
        return dicoPointsTechniquesCompletes
