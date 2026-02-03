#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
Fonctions utilitaires dépendantes de QGIS pour le plugin PoleAerien.
ATTENTION: Ne pas importer dans les workers threads sans précaution.
"""

from qgis.core import (
    QgsProject, QgsLayerTreeLayer, QgsMessageLog, Qgis, 
    QgsSpatialIndex, QgsFeatureRequest, QgsMapLayer, QgsWkbTypes, NULL
)
import re
from .core_utils import normalize_appui_num, temps_ecoule, find_default_layer_index, normalize_appui_num_bt


def remove_group(name):
    """Suppression d'un groupe de couches par son nom.
    
    Args:
        name: Nom du groupe à supprimer
    """
    root = QgsProject.instance().layerTreeRoot()
    group = root.findGroup(name)
    if group is not None:
        for child in group.children():
            if hasattr(child, 'layerId'):
                QgsProject.instance().removeMapLayer(child.layerId())
        root.removeChildNode(group)


def layer_group_error(couche, nom_etude):
    """Ajoute une couche dans un groupe d'erreur.
    
    Args:
        couche: La couche QGIS à ajouter
        nom_etude: Nom de l'étude pour nommer le groupe
    """
    root = QgsProject.instance().layerTreeRoot()
    group_name = f"ERROR_{nom_etude}"
    group = root.findGroup(group_name)
    
    if not group:
        group = root.addGroup(group_name)
    
    group.insertChildNode(1, QgsLayerTreeLayer(couche))


def insert_layer_in_group(couche, group_name):
    """Ajoute une couche dans un groupe spécifié.
    
    Args:
        couche: La couche QGIS à ajouter
        group_name: Nom du groupe
    """
    root = QgsProject.instance().layerTreeRoot()
    group = root.findGroup(group_name)
    if not group:
        group = root.addGroup(group_name)
    group.insertChildNode(1, QgsLayerTreeLayer(couche))


def get_layer_safe(layer_name, context=""):
    """Récupère une couche QGIS de manière sécurisée.
    
    Args:
        layer_name: Nom de la couche à récupérer
        context: Contexte pour message d'erreur (ex: "MAJ FT/BT")
        
    Returns:
        QgsVectorLayer valide
        
    Raises:
        ValueError: Si la couche n'existe pas, n'est pas chargée ou invalide
    """
    if not layer_name:
        raise ValueError(f"[{context}] Nom de couche vide ou None")
    
    layers = QgsProject.instance().mapLayersByName(layer_name)
    if not layers:
        raise ValueError(
            f"[{context}] Couche '{layer_name}' introuvable. "
            "Vérifiez qu'elle est chargée dans le projet."
        )
    
    layer = layers[0]
    if not layer.isValid():
        raise ValueError(
            f"[{context}] Couche '{layer_name}' invalide. "
            "Source de données inaccessible."
        )
    return layer


def get_layer_fields(layer_name, default_pattern=r'etudes*'):
    """Récupère la liste des champs d'une couche avec sélection par défaut.
    
    Args:
        layer_name: Nom de la couche
        default_pattern: Pattern regex pour sélectionner le champ par défaut
        
    Returns:
        Tuple (champs_list, index_defaut)
    """
    layers = QgsProject.instance().mapLayers().values()
    champs_list = ['']
    index_defaut = 0
    
    for layer in layers:
        if layer.name() == layer_name:
            for champs in layer.fields():
                champs_list.append(str(champs.name()))
    
    if len(champs_list) > 1:
        regexp = re.compile(default_pattern)
        for i, valeur in enumerate(champs_list):
            if regexp.search(valeur.lower()):
                index_defaut = i
                break
    
    return champs_list, index_defaut


def get_layers_by_geometry(geom_types):
    """Récupère la liste des couches vectorielles filtrées par type de géométrie.
    
    Args:
        geom_types: Tuple des types de géométrie QGIS (0=Point, 1=Line, 2=Polygon, etc.)
        
    Returns:
        Liste triée des noms de couches correspondant
    """
    layers = QgsProject.instance().mapLayers().values()
    layer_list = ['']
    warned = False
    
    for layer in layers:
        if layer.type() == QgsMapLayer.VectorLayer:
            try:
                if QgsWkbTypes.geometryType(layer.wkbType()) in geom_types:
                    layer_list.append(layer.name())
            except Exception as err:
                if not warned:
                    QgsMessageLog.logMessage(
                        f"[utils.get_layers_by_geometry] Type geom invalide: {err}",
                        "PoleAerien",
                        Qgis.Warning
                    )
                    warned = True
                continue
    
    layer_list.sort()
    return layer_list


def set_default_layer_for_combobox(combobox, pattern):
    """Définit la couche par défaut dans un QgsMapLayerComboBox selon un pattern.
    
    Args:
        combobox: QgsMapLayerComboBox à configurer
        pattern: Pattern regex pour matcher le nom de couche
    """
    regexp = re.compile(pattern, re.IGNORECASE)
    
    # Chercher parmi les couches vectorielles de type Point
    for layer in QgsProject.instance().mapLayers().values():
        if layer.type() == QgsMapLayer.VectorLayer:
            if regexp.search(layer.name()):
                combobox.setLayer(layer)
                return
    # Aucune correspondance - garde sélection actuelle


def validate_same_crs(ref_layer, other_layer, context=""):
    """Valide que deux couches partagent le même CRS.

    Args:
        ref_layer: Couche de référence
        other_layer: Couche à comparer
        context: Contexte pour message d'erreur

    Raises:
        ValueError: CRS différents
    """
    if ref_layer is None or other_layer is None:
        return
    if ref_layer.crs() != other_layer.crs():
        ref_crs = ref_layer.crs().authid() or ref_layer.crs().toWkt()
        other_crs = other_layer.crs().authid() or other_layer.crs().toWkt()
        raise ValueError(
            f"[{context}] CRS incoherent: {ref_layer.name()}={ref_crs} vs "
            f"{other_layer.name()}={other_crs}"
        )


def validate_crs_compatibility(layer, expected_crs="EPSG:2154", context=""):
    """Valide qu'une couche utilise le CRS attendu (EPSG:2154 par défaut).
    
    QGIS 3.28 REQUIREMENT: CRS MUST be explicit and validated.
    Le plugin assume EPSG:2154 (Lambert 93) pour tous les calculs géométriques.
    
    Args:
        layer: QgsVectorLayer à vérifier
        expected_crs: CRS attendu (défaut: EPSG:2154)
        context: Contexte pour message d'erreur
        
    Raises:
        ValueError: Si CRS incompatible ou invalide
    """
    if layer is None:
        raise ValueError(f"[{context}] Couche None fournie pour validation CRS")
    
    if not layer.isValid():
        raise ValueError(f"[{context}] Couche '{layer.name()}' invalide")
    
    layer_crs = layer.crs()
    if not layer_crs.isValid():
        raise ValueError(
            f"[{context}] CRS invalide pour couche '{layer.name()}'"
        )
    
    layer_crs_id = layer_crs.authid()
    if layer_crs_id != expected_crs:
        raise ValueError(
            f"[{context}] CRS incompatible pour '{layer.name()}':\n"
            f"  Attendu: {expected_crs}\n"
            f"  Reçu: {layer_crs_id}\n"
            f"  Veuillez reprojeter la couche en {expected_crs} avant l'analyse."
        )
    
    QgsMessageLog.logMessage(
        f"[{context}] CRS validé: {layer.name()} = {layer_crs_id}",
        "PoleAerien",
        Qgis.Info
    )


def build_spatial_index(layer, request=None):
    """Construit index spatial + cache features.
    
    Args:
        layer: QgsVectorLayer source
        request: QgsFeatureRequest optionnel pour filtrer
        
    Returns:
        tuple: (QgsSpatialIndex, dict{fid: QgsFeature})
    """
    idx = QgsSpatialIndex()
    cache = {}
    feats = layer.getFeatures(request) if request else layer.getFeatures()
    
    for feat in feats:
        if feat.hasGeometry():
            idx.addFeature(feat)
            cache[feat.id()] = feat
    
    return idx, cache


def make_ordered_request(field, expression=None, ascending=True):
    """Crée QgsFeatureRequest avec tri.
    
    Args:
        field: Champ pour le tri
        expression: Expression filtre optionnelle (str)
        ascending: Ordre croissant (défaut True)
        
    Returns:
        QgsFeatureRequest configuré
    """
    request = QgsFeatureRequest()
    if expression:
        request.setFilterExpression(expression)
    
    clause = QgsFeatureRequest.OrderByClause(field, ascending=ascending)
    request.setOrderBy(QgsFeatureRequest.OrderBy([clause]))
    
    return request


def detect_duplicates(layer, field, request=None):
    """Détecte doublons sur un champ.
    
    Args:
        layer: QgsVectorLayer source
        field: Nom du champ à vérifier
        request: QgsFeatureRequest optionnel
        
    Returns:
        list: Valeurs en doublon
    """
    doublons = []
    vus = set()
    feats = layer.getFeatures(request) if request else layer.getFeatures()
    
    for feat in feats:
        val = feat[field]
        if val in vus:
            if val not in doublons:
                doublons.append(val)
        else:
            vus.add(val)
    
    return doublons


def verifications_donnees_etude(
    table_poteau, table_etude, colonne_etude,
    pot_type_filter, context
):
    """Vérifie doublons études + poteaux hors étude.
    
    Args:
        table_poteau: Nom couche poteaux
        table_etude: Nom couche études
        colonne_etude: Champ nom étude
        pot_type_filter: 'POT-FT' ou 'POT-BT'
        context: Contexte erreur
    
    Returns:
        tuple: (doublons_etudes, poteaux_hors_etude)
        
    Raises:
        ValueError: Couche introuvable ou CRS incohérent
    """
    infra_pt_pot = get_layer_safe(table_poteau, context)
    etude = get_layer_safe(table_etude, context)
    validate_same_crs(infra_pt_pot, etude, context)
    
    # Request poteaux filtrés
    req_pot = make_ordered_request('inf_type', f"inf_type LIKE '{pot_type_filter}'")
    
    # Request études triées
    req_etude = make_ordered_request(colonne_etude)
    
    # Index spatial études
    idx_etudes, etudes_dict = build_spatial_index(etude, req_etude)
    
    # Poteaux hors étude
    hors_etude = []
    for feat_pot in infra_pt_pot.getFeatures(req_pot):
        raw_inf = feat_pot["inf_num"]
        if not raw_inf or raw_inf == NULL or not feat_pot.hasGeometry():
            continue
        
        bbox = feat_pot.geometry().boundingBox()
        dans_etude = False
        for fid in idx_etudes.intersects(bbox):
            if etudes_dict[fid].geometry().contains(feat_pot.geometry()):
                dans_etude = True
                break
        
        if not dans_etude:
            hors_etude.append(raw_inf)
    
    # Doublons études
    doublons = detect_duplicates(etude, colonne_etude, req_etude)
    
    return doublons, hors_etude


def liste_poteaux_par_etude(
    table_poteau, table_etude, colonne_etude,
    pot_type_filter, context
):
    """Liste poteaux par étude avec détection terrains privés.
    
    Args:
        table_poteau: Nom couche poteaux
        table_etude: Nom couche études
        colonne_etude: Champ nom étude
        pot_type_filter: 'POT-FT' ou 'POT-BT'
        context: Contexte erreur
    
    Returns:
        tuple: (dict_poteaux_par_etude, dict_poteaux_prives)
        
    Raises:
        ValueError: Couche introuvable ou CRS incohérent
    """
    from .security_rules import est_terrain_prive
    
    infra_pt_pot = get_layer_safe(table_poteau, context)
    etude = get_layer_safe(table_etude, context)
    validate_same_crs(infra_pt_pot, etude, context)
    
    req_pot = make_ordered_request('inf_type', f"inf_type LIKE '{pot_type_filter}'")
    req_etude = make_ordered_request(colonne_etude)
    
    # Index spatial poteaux
    idx_pot, poteaux_dict = build_spatial_index(infra_pt_pot, req_pot)
    
    # Index champ commentaire
    idx_commentaire = infra_pt_pot.fields().indexFromName('commentaire')
    
    dico_etude = {}
    dico_prives = {}
    
    for feat_etude in etude.getFeatures(req_etude):
        if not feat_etude.hasGeometry():
            continue
        
        liste_pot = []
        liste_priv = []
        nom_etude = feat_etude[colonne_etude]
        bbox = feat_etude.geometry().boundingBox()
        
        for fid in idx_pot.intersects(bbox):
            feat_pot = poteaux_dict[fid]
            if feat_etude.geometry().contains(feat_pot.geometry()):
                raw_inf = feat_pot["inf_num"]
                if raw_inf and raw_inf != NULL:
                    liste_pot.append(raw_inf)
                    if idx_commentaire >= 0 and est_terrain_prive(feat_pot[idx_commentaire]):
                        liste_priv.append(raw_inf)
        
        if liste_pot:
            dico_etude[nom_etude] = liste_pot
        if liste_priv:
            dico_prives[nom_etude] = liste_priv
    
    return dico_etude, dico_prives
