#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
Fonctions utilitaires pures (sans dépendance QGIS) pour le plugin PoleAerien.
Sécurisées pour l'utilisation dans les threads workers.
"""

import re
import xml.etree.ElementTree as ET


# =============================================================================
# PARSING HELPERS (centralises - utilises par comac_db_reader, pcm_parser)
# =============================================================================

def safe_float(value, default: float = 0.0) -> float:
    """Parse float avec protection None/erreur et virgule française.
    
    Args:
        value: Valeur à parser (None, int, float, str)
        default: Valeur par défaut si parsing échoue
        
    Returns:
        float parsée ou default
    """
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    val = str(value).strip().replace('m', '')
    if not val:
        return default
    try:
        return float(val.replace(',', '.'))
    except (ValueError, AttributeError):
        return default


def safe_int(value, default: int = 0) -> int:
    """Parse int avec protection None/erreur.
    
    Args:
        value: Valeur à parser (None, int, float, str)
        default: Valeur par défaut si parsing échoue
        
    Returns:
        int parsé ou default
    """
    if value is None:
        return default
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    val = str(value).strip()
    if not val:
        return default
    try:
        return int(float(val.replace(',', '.')))
    except (ValueError, AttributeError):
        return default


def get_xml_text(element: ET.Element, tag: str, default: str = "") -> str:
    """Extrait le texte d'un sous-élément XML de manière sûre.
    
    Args:
        element: Élément XML parent
        tag: Nom du sous-élément à chercher
        default: Valeur par défaut si absent
        
    Returns:
        Texte du sous-élément ou default
    """
    child = element.find(tag)
    return child.text.strip() if child is not None and child.text else default


def get_xml_float(element: ET.Element, tag: str, default: float = 0.0) -> float:
    """Extrait un float d'un sous-élément XML.
    
    Args:
        element: Élément XML parent
        tag: Nom du sous-élément
        default: Valeur par défaut
        
    Returns:
        float parsée ou default
    """
    return safe_float(get_xml_text(element, tag), default)


def parse_bool(value: str) -> bool:
    """Parse booléen (0/1)."""
    return str(value).strip() == '1'


# =============================================================================
# EXPORT PATH HELPER
# =============================================================================

def build_export_path(chemin: str, default_name: str) -> str:
    """Construit le chemin complet pour un fichier d'export Excel.
    
    Args:
        chemin: Chemin fourni (répertoire ou fichier)
        default_name: Nom de fichier par défaut (ex: "analyse.xlsx")
        
    Returns:
        Chemin complet vers le fichier .xlsx
    """
    import os
    if os.path.isdir(chemin):
        return os.path.join(chemin, default_name)
    if not chemin.endswith('.xlsx'):
        return chemin + '.xlsx'
    return chemin


# =============================================================================
# NORMALISATION APPUIS
# =============================================================================

def normalize_appui_num(inf_num, strip_e_prefix=False, strip_bt_prefix=False):
    """Normalise un numéro d'appui pour comparaison.
    
    Args:
        inf_num: Numéro d'appui brut (ex: "1016436/63041", "E000123", "BT-123")
        strip_e_prefix: Si True, enlève préfixe E (pour matching existant)
        strip_bt_prefix: Si True, enlève préfixes BT-/BT
        
    Returns:
        str: Numéro normalisé (sans zéros de tête, partie avant le /)
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
        
        # Préfixe E: conserver ou enlever selon strip_e_prefix
        if s.startswith("E"):
            if strip_e_prefix:
                s = s[1:]
            else:
                return s
        
        # Préfixe BT: enlever si demandé
        if strip_bt_prefix:
            if s.startswith("BT-"):
                s = s[3:]
            elif s.startswith("BT"):
                s = s[2:]
        
        # Nettoyer et normaliser
        s_clean = s.replace(" ", "")
        
        # Si c'est un nombre pur, enlever les zéros de tête
        if s_clean.isdigit():
            return s_clean.lstrip("0") or "0"
        
        return s
    except Exception:
        return ""


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


# =============================================================================
# FILE FILTERING (shared by C6_vs_Bd, police_workflow, etc.)
# =============================================================================

_OUTPUT_PREFIXES = (
    'analyse_c6', 'policec6_export', 'police_c6_export',
    'rapport_', 'rapport_batch_', 'export_', '~$',
    'analyse_', 'ficheappui', 'gespot',
)
_OUTPUT_KEYWORDS = (
    'ficheappui', '_c7', 'annexe c7', 'annexe_c7',
    'comac', 'capft', 'cap_ft',
)


def is_plugin_output_file(filename):
    """Return True if the file is a plugin output / temp file, not a real C6 annexe.

    Args:
        filename: Filename (basename, not full path)

    Returns:
        True if the file should be excluded from C6 detection
    """
    name = filename.lower()
    if any(name.startswith(p) for p in _OUTPUT_PREFIXES):
        return True
    if any(kw in name for kw in _OUTPUT_KEYWORDS):
        return True
    return False


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
