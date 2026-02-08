# -*- coding: utf-8 -*-
"""
Parseur de fichiers COMAC (.pcm) pour extraction des données de sécurité.
Les fichiers .pcm sont au format XML ISO-8859-1.

Structure:
- Supports: poteaux avec attributs (hauteur, traverse, portée molle)
- LignesTCF: lignes télécom avec câbles FO et portées
- LignesBT: lignes BT avec conducteurs
"""

import os
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple

try:
    from .core_utils import safe_float, parse_bool, get_xml_text, get_xml_float
except ImportError:
    from core_utils import safe_float, parse_bool, get_xml_text, get_xml_float

try:
    from qgis.core import QgsMessageLog, Qgis
    HAS_QGIS = True
except ImportError:
    HAS_QGIS = False

# Import security_rules - relative import for plugin, absolute for standalone testing
try:
    from .security_rules import (
        get_capacite_fo_from_code,
        verifier_portee,
        verifier_distance_sol,
        PORTEES_MAX_ZVN,
    )
except ImportError as e:
    # Standalone mode (testing outside QGIS plugin context)
    try:
        from security_rules import (
            get_capacite_fo_from_code,
            verifier_portee,
            verifier_distance_sol,
            PORTEES_MAX_ZVN,
        )
    except ImportError:
        raise ImportError(f"security_rules module not found: {e}") from e


def _log_message(msg: str, tag: str = "PCM_PARSER", level: int = 0):
    """Log avec fallback print si hors QGIS"""
    if HAS_QGIS:
        qgis_level = Qgis.Info if level == 0 else (Qgis.Warning if level == 1 else Qgis.Critical)
        QgsMessageLog.logMessage(msg, tag, qgis_level)
    else:
        print(f"[{tag}] {msg}")


# =============================================================================
# DATACLASSES - Structures de données
# =============================================================================

@dataclass
class Support:
    """Poteau/support extrait du .pcm"""
    nom: str
    nature: str = ""           # BE, FT, BO
    hauteur: float = 0.0
    classe: str = ""           # A, D, FT, FTX, S, FACADE
    effort: float = 0.0
    traverse_existante: float = 0.0
    traverse_a_poser: float = 0.0
    portee_molle: bool = False
    non_calcule: bool = False
    illisible: bool = False
    x: float = 0.0
    y: float = 0.0
    etat: str = ""


@dataclass
class LigneTCF:
    """Ligne télécom (coaxial ou fibre)"""
    cable: str = ""            # L1092-13-P, 98-8-4
    capacite_fo: int = 0
    a_poser: bool = False
    supports: List[str] = field(default_factory=list)
    traverses: List[int] = field(default_factory=list)  # 1=existante, 2=à poser
    portees: List[float] = field(default_factory=list)
    portee_max: float = 0.0


@dataclass 
class LigneBT:
    """Ligne BT électrique"""
    conducteur: str = ""
    supports: List[str] = field(default_factory=list)
    portees: List[float] = field(default_factory=list)


@dataclass
class VerifSecurite:
    """Résultat vérification sécurité pour un segment"""
    support_depart: str
    support_arrivee: str
    portee: float
    capacite_fo: int
    portee_max: float
    valide: bool
    depassement: float = 0.0
    hauteur_traverse: float = 0.0
    hauteur_valide: bool = True
    cable: str = ""
    a_poser: bool = False


@dataclass
class EtudePCM:
    """Étude COMAC complète"""
    num_etude: str = ""
    version: str = ""
    commune: str = ""
    hypotheses: List[str] = field(default_factory=list)
    supports: Dict[str, Support] = field(default_factory=dict)
    lignes_tcf: List[LigneTCF] = field(default_factory=list)
    lignes_bt: List[LigneBT] = field(default_factory=list)
    verifications: List[VerifSecurite] = field(default_factory=list)
    erreurs_parse: List[str] = field(default_factory=list)


# =============================================================================
# FONCTIONS UTILITAIRES (safe_float, parse_bool, get_xml_text, get_xml_float importés de core_utils)
# =============================================================================

def _get_bool(element, tag: str) -> bool:
    """Extrait booléen d'un sous-élément"""
    return parse_bool(get_xml_text(element, tag))


# =============================================================================
# PARSING PRINCIPAL
# =============================================================================

def parse_pcm_file(filepath: str) -> Optional[EtudePCM]:
    """
    Parse un fichier .pcm et retourne une structure EtudePCM.
    
    Args:
        filepath: Chemin complet du fichier .pcm
    
    Returns:
        EtudePCM ou None si erreur
    """
    if not os.path.exists(filepath):
        _log_message(f"Fichier introuvable: {filepath}", "PCM_PARSER", 1)
        return None
    
    etude = EtudePCM()
    
    try:
        # Parse XML avec encodage ISO-8859-1
        tree = ET.parse(filepath)
        root = tree.getroot()
        
        # Métadonnées
        etude.num_etude = get_xml_text(root, 'NumEtude')
        etude.version = get_xml_text(root, 'Version')
        etude.commune = get_xml_text(root, 'Commune')
        
        # Hypothèses climatiques
        hypotheses_elem = root.find('Hypotheses')
        if hypotheses_elem is not None:
            for hyp in hypotheses_elem.findall('Hypothese'):
                if hyp.text:
                    etude.hypotheses.append(hyp.text.strip())
        
        # Supports
        _parse_supports(root, etude)
        
        # Lignes TCF (télécom/fibre)
        _parse_lignes_tcf(root, etude)
        
        # Lignes BT
        _parse_lignes_bt(root, etude)
        
    except ET.ParseError as e:
        etude.erreurs_parse.append(f"Erreur XML: {e}")
        _log_message(f"Erreur parse {filepath}: {e}", "PCM_PARSER", 2)
    except Exception as e:
        etude.erreurs_parse.append(f"Erreur: {e}")
        _log_message(f"Erreur {filepath}: {e}", "PCM_PARSER", 2)
    
    return etude


def _parse_supports(root: ET.Element, etude: EtudePCM):
    """Parse section <Supports>"""
    supports_elem = root.find('Supports')
    if supports_elem is None:
        return
    
    for supp_elem in supports_elem.findall('Support'):
        nom = get_xml_text(supp_elem, 'Nom')
        if not nom:
            continue
        
        support = Support(
            nom=nom,
            nature=get_xml_text(supp_elem, 'Nature'),
            hauteur=get_xml_float(supp_elem, 'Hauteur'),
            classe=get_xml_text(supp_elem, 'Classe'),
            effort=get_xml_float(supp_elem, 'Effort'),
            traverse_existante=get_xml_float(supp_elem, 'TraverseExistante1'),
            traverse_a_poser=get_xml_float(supp_elem, 'TraverseAPoser2'),
            portee_molle=_get_bool(supp_elem, 'PorteeMolle'),
            non_calcule=_get_bool(supp_elem, 'NonCalcule'),
            illisible=_get_bool(supp_elem, 'Illisible'),
            x=get_xml_float(supp_elem, 'X'),
            y=get_xml_float(supp_elem, 'Y'),
            etat=get_xml_text(supp_elem, 'Etat')
        )
        etude.supports[nom] = support


# DEBUG FLAG - set False en production
_DEBUG_COMAC_CAPA = False

def _parse_lignes_tcf(root: ET.Element, etude: EtudePCM):
    """Parse section <LignesTCF> (télécom coaxial et fibre)"""
    tcf_elem = root.find('LignesTCF')
    if tcf_elem is None:
        if _DEBUG_COMAC_CAPA:
            print(f"[PCM_CAPA] Etude {etude.num_etude}: <LignesTCF> non trouvé")
        return
    
    nb_lignes = len(tcf_elem.findall('LigneTCF'))
    if _DEBUG_COMAC_CAPA:
        print(f"[PCM_CAPA] Etude {etude.num_etude}: {nb_lignes} lignes TCF trouvées")
    
    for idx, ligne_elem in enumerate(tcf_elem.findall('LigneTCF')):
        cable = get_xml_text(ligne_elem, 'Cable')
        capacite = get_capacite_fo_from_code(cable, debug=_DEBUG_COMAC_CAPA)
        
        if _DEBUG_COMAC_CAPA:
            print(f"[PCM_CAPA] Ligne {idx+1}/{nb_lignes}: cable='{cable}' -> capacite_fo={capacite}")
        
        ligne = LigneTCF(
            cable=cable,
            capacite_fo=capacite,
            a_poser=_get_bool(ligne_elem, 'APoser')
        )
        
        # Supports et traverses
        supports_elem = ligne_elem.find('Supports')
        if supports_elem is not None:
            current_support = None
            for child in supports_elem:
                if child.tag == 'Support' and child.text:
                    current_support = child.text.strip()
                    ligne.supports.append(current_support)
                elif child.tag == 'Traverse' and child.text:
                    ligne.traverses.append(int(child.text.strip()))
        
        # Portées
        portees_elem = ligne_elem.find('Portees')
        if portees_elem is not None:
            for portee_elem in portees_elem.findall('Portee'):
                if portee_elem.text:
                    ligne.portees.append(safe_float(portee_elem.text))
        
        # Portée max selon capacité
        if capacite > 0:
            ligne.portee_max = PORTEES_MAX_ZVN.get(capacite, 0)
        
        etude.lignes_tcf.append(ligne)


def _parse_lignes_bt(root: ET.Element, etude: EtudePCM):
    """Parse section <LignesBT>"""
    bt_elem = root.find('LignesBT')
    if bt_elem is None:
        return
    
    for ligne_elem in bt_elem.findall('LigneBT'):
        ligne = LigneBT(
            conducteur=get_xml_text(ligne_elem, 'Conducteur')
        )
        
        # Supports
        supports_elem = ligne_elem.find('Supports')
        if supports_elem is not None:
            for child in supports_elem:
                if child.tag == 'Support' and child.text:
                    ligne.supports.append(child.text.strip())
        
        # Portées
        portees_elem = ligne_elem.find('Portees')
        if portees_elem is not None:
            for portee_elem in portees_elem.findall('Portee'):
                if portee_elem.text:
                    ligne.portees.append(safe_float(portee_elem.text))
        
        etude.lignes_bt.append(ligne)


# =============================================================================
# VERIFICATION SECURITE
# =============================================================================

def verifier_securite_etude(etude: EtudePCM, zone: str = 'ZVN') -> List[VerifSecurite]:
    """
    Vérifie les règles de sécurité pour toutes les lignes FO d'une étude.
    
    Args:
        etude: Étude PCM parsée
        zone: Zone climatique ('ZVN' ou 'ZVF')
    
    Returns:
        Liste des vérifications (une par segment de portée)
    """
    verifications = []
    
    for ligne in etude.lignes_tcf:
        # Ignorer câbles non-FO (coaxial 98-8-4)
        if ligne.capacite_fo == 0:
            continue
        
        # Vérifier chaque portée
        for i, portee in enumerate(ligne.portees):
            # Supports départ/arrivée
            supp_depart = ligne.supports[i] if i < len(ligne.supports) else ""
            supp_arrivee = ligne.supports[i + 1] if i + 1 < len(ligne.supports) else ""
            
            # Hauteur traverse (à poser si traverse=2, existante si traverse=1)
            hauteur_traverse = 0.0
            if supp_arrivee and supp_arrivee in etude.supports:
                supp = etude.supports[supp_arrivee]
                # Traverse index correspond au support d'arrivée
                trav_idx = i + 1 if i + 1 < len(ligne.traverses) else -1
                if trav_idx >= 0 and trav_idx < len(ligne.traverses):
                    if ligne.traverses[trav_idx] == 2:
                        hauteur_traverse = supp.traverse_a_poser
                    else:
                        hauteur_traverse = supp.traverse_existante
            
            # Vérification portée
            result_portee = verifier_portee(portee, ligne.capacite_fo, zone)
            
            # Vérification hauteur sol (>= 4m)
            result_hauteur = verifier_distance_sol(hauteur_traverse) if hauteur_traverse > 0 else {'valide': True}
            
            verif = VerifSecurite(
                support_depart=supp_depart,
                support_arrivee=supp_arrivee,
                portee=portee,
                capacite_fo=ligne.capacite_fo,
                portee_max=result_portee.get('portee_max', 0),
                valide=result_portee.get('valide', True) and result_hauteur.get('valide', True),
                depassement=result_portee.get('depassement', 0),
                hauteur_traverse=hauteur_traverse,
                hauteur_valide=result_hauteur.get('valide', True),
                cable=ligne.cable,
                a_poser=ligne.a_poser
            )
            verifications.append(verif)
    
    etude.verifications = verifications
    return verifications


# =============================================================================
# FONCTIONS DE HAUT NIVEAU
# =============================================================================

def parse_repertoire_pcm(repertoire: str, zone: str = 'ZVN') -> Tuple[Dict[str, EtudePCM], Dict[str, str]]:
    """
    Parse tous les fichiers .pcm d'un répertoire.
    
    Args:
        repertoire: Chemin du répertoire
        zone: Zone climatique
    
    Returns:
        Tuple (dict études par nom, dict erreurs par fichier)
    """
    etudes = {}
    erreurs = {}
    
    for subdir, _, files in os.walk(repertoire):
        for name in files:
            if name.lower().endswith('.pcm'):
                filepath = os.path.join(subdir, name)
                try:
                    etude = parse_pcm_file(filepath)
                    if etude:
                        # Vérification sécurité
                        verifier_securite_etude(etude, zone)
                        etudes[etude.num_etude or name] = etude
                except Exception as e:
                    erreurs[filepath] = str(e)
    
    return etudes, erreurs


def get_anomalies_securite(etudes: Dict[str, EtudePCM]) -> List[dict]:
    """
    Extrait toutes les anomalies de sécurité des études.
    
    Returns:
        Liste de dicts avec détails anomalies
    """
    anomalies = []
    
    for nom_etude, etude in etudes.items():
        for verif in etude.verifications:
            if not verif.valide:
                anomalie = {
                    'etude': nom_etude,
                    'support_depart': verif.support_depart,
                    'support_arrivee': verif.support_arrivee,
                    'cable': verif.cable,
                    'portee': verif.portee,
                    'portee_max': verif.portee_max,
                    'depassement': verif.depassement,
                    'capacite_fo': verif.capacite_fo,
                    'hauteur_traverse': verif.hauteur_traverse,
                    'portee_ko': verif.depassement > 0,
                    'hauteur_ko': not verif.hauteur_valide,
                    'a_poser': verif.a_poser
                }
                anomalies.append(anomalie)
    
    return anomalies


def get_supports_portee_molle(etudes: Dict[str, EtudePCM]) -> List[dict]:
    """
    Extrait tous les supports marqués portée molle par COMAC.
    
    Returns:
        Liste de dicts avec détails supports
    """
    supports_pm = []
    
    for nom_etude, etude in etudes.items():
        for nom_supp, supp in etude.supports.items():
            if supp.portee_molle:
                supports_pm.append({
                    'etude': nom_etude,
                    'support': nom_supp,
                    'nature': supp.nature,
                    'hauteur': supp.hauteur,
                    'classe': supp.classe,
                    'etat': supp.etat
                })
    
    return supports_pm
