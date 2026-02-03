#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
Fonctions utilitaires pures (sans dépendance QGIS) pour le plugin PoleAerien.
Sécurisées pour l'utilisation dans les threads workers.
"""

import re

def normalize_appui_num(inf_num):
    """Normalise un numéro d'appui pour comparaison.
    
    Args:
        inf_num: Numéro d'appui brut (ex: "1016436/63041", "E000123", "123")
        
    Returns:
        Numéro normalisé (sans zéros de tête, partie avant le /)
    """
    try:
        if inf_num is None:
            return ""
        s = str(inf_num).strip()
        if not s:
            return ""
        
        # Extraire la partie avant le slash (format numéro/insee)
        if "/" in s:
            s = s.split("/")[0].strip()
        
        # Enlever le suffixe .0 (conversion float -> str)
        if s.endswith(".0"):
            s = s[:-2]
        
        # Préfixe E: conserver tel quel
        if s.startswith("E"):
            return s
        
        # Nettoyer et normaliser
        s_clean = s.replace(" ", "")
        
        # Si c'est un nombre pur, enlever les zéros de tête
        if s_clean.isdigit():
            return s_clean.lstrip("0") or "0"
        
        # Tenter d'extraire les 7 premiers chiffres si présents
        if len(s_clean) >= 7:
            prefix = s_clean[:7]
            if prefix.isdigit():
                return prefix.lstrip("0") or "0"
        
        return s
    except Exception:
        return ""


def normalize_appui_num_bt(inf_num, strip_bt_prefix=True, strip_e_prefix=False):
    """Normalise appui BT/FT unifié.
    
    Args:
        inf_num: Numéro brut
        strip_bt_prefix: Si True, enlève préfixes BT-/BT
        strip_e_prefix: Si True, enlève préfixe E (pour matching)
        
    Returns:
        str: Numéro normalisé
    """
    if inf_num is None:
        return ""
    s = str(inf_num).strip()
    if not s:
        return ""
    
    s = s.split("/")[0].strip()
    
    if s.endswith(".0"):
        s = s[:-2]
    
    if s.startswith("E"):
        if strip_e_prefix:
            s = s[1:]
        else:
            return s
    
    if strip_bt_prefix:
        if s.startswith("BT-"):
            s = s[3:]
        elif s.startswith("BT"):
            s = s[2:]
    
    s_clean = s.replace(" ", "")
    if s_clean.isdigit():
        return s_clean.lstrip("0") or "0"
    
    return s


def temps_ecoule(seconde):
    """Formate une durée en secondes en format lisible.
    
    Args:
        seconde: Nombre de secondes
        
    Returns:
        Chaîne formatée (ex: "2mn : 30sec")
    """
    seconds = seconde % (24 * 3600)
    hour = int(seconds // 3600)
    seconds %= 3600
    minutes = int(seconds // 60)
    seconds %= 60

    if hour > 0:
        return f"{hour}h: {int(minutes)}mn : {int(seconds)}sec"
    return f"{minutes}mn : {int(seconds)}sec"


def find_default_layer_index(layer_list, pattern):
    """Trouve l'index d'une couche par défaut selon un pattern regex.
    
    Args:
        layer_list: Liste des noms de couches
        pattern: Pattern regex à rechercher
        
    Returns:
        Index de la couche correspondante ou 0
    """
    regexp = re.compile(pattern, re.IGNORECASE)
    for i, valeur in enumerate(layer_list):
        if valeur and regexp.search(valeur):
            return i
    return 0
