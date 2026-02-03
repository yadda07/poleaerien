# -*- coding: utf-8 -*-
"""
Loader combiné Excel + PCM pour données COMAC.
Fusionne les deux sources pour obtenir les données complètes.

Architecture:
- PCM: portées par segment, supports liés, câble FO, traverse
- Excel: hauteur hors sol (seule source)
"""

import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from .security_rules import (
        get_capacite_fo_from_code,
        verifier_portee,
        PORTEES_MAX_ZVN,
        PORTEES_MAX_ZVF
    )
    from .comac_db_reader import (
        get_zone_vent_from_hypotheses,
        get_zone_vent_from_insee
    )
except ImportError:
    from security_rules import (
        get_capacite_fo_from_code,
        verifier_portee,
        PORTEES_MAX_ZVN,
        PORTEES_MAX_ZVF
    )
    from comac_db_reader import (
        get_zone_vent_from_hypotheses,
        get_zone_vent_from_insee
    )


# =============================================================================
# DATACLASSES - Structures fusionnées
# =============================================================================

@dataclass
class SupportMerged:
    """Support avec données Excel + PCM fusionnées"""
    nom: str
    # PCM
    hauteur_totale: float = 0.0
    traverse_existante: float = 0.0
    traverse_a_poser: float = 0.0
    portee_molle: bool = False
    nature: str = ""
    classe: str = ""
    x: float = 0.0
    y: float = 0.0
    etat: str = ""
    # Excel uniquement
    hauteur_hors_sol: float = 0.0
    conducteur: str = ""


@dataclass
class SegmentFO:
    """Segment de câble FO entre 2 supports"""
    support_depart: str
    support_arrivee: str
    portee: float
    capacite_fo: int
    cable: str
    traverse_type: int  # 1=existante, 2=à poser
    a_poser: bool = False
    # Vérification
    portee_max: float = 0.0
    depassement: float = 0.0
    valide: bool = True


@dataclass
class EtudeComacMerged:
    """Étude COMAC fusionnée Excel + PCM"""
    num_etude: str = ""
    commune: str = ""
    zone_climatique: str = "ZVN"
    hypotheses: List[str] = field(default_factory=list)
    # Données fusionnées
    supports: Dict[str, SupportMerged] = field(default_factory=dict)
    segments_fo: List[SegmentFO] = field(default_factory=list)
    # Stats
    nb_segments_invalides: int = 0
    portee_max_depassement: float = 0.0
    # Métadonnées
    source_pcm: str = ""
    source_excel: str = ""
    erreurs: List[str] = field(default_factory=list)


# =============================================================================
# HELPERS
# =============================================================================

def _safe_float(value, default: float = 0.0) -> float:
    """Parse float avec protection None/erreur"""
    if value is None:
        return default
    try:
        return float(str(value).replace(',', '.').replace('m', '').strip())
    except (ValueError, AttributeError):
        return default


def _safe_int(value, default: int = 0) -> int:
    """Parse int avec protection"""
    if value is None:
        return default
    try:
        return int(float(str(value).replace(',', '.')))
    except (ValueError, AttributeError):
        return default


def _get_text(elem: ET.Element, tag: str, default: str = "") -> str:
    """Get text from XML element safely"""
    child = elem.find(tag)
    return child.text.strip() if child is not None and child.text else default


def _get_float(elem: ET.Element, tag: str, default: float = 0.0) -> float:
    """Get float from XML element"""
    return _safe_float(_get_text(elem, tag), default)


def _get_bool(elem: ET.Element, tag: str) -> bool:
    """Get bool from XML element (0/1)"""
    return _get_text(elem, tag) == '1'


def _extract_num_etude(filepath: str) -> str:
    """Extrait num_etude depuis nom de fichier ou contenu"""
    basename = os.path.basename(filepath)
    # Pattern: NGE-XXXXX-XXXXX-PA-X
    match = re.search(r'(NGE-\w+-\d+-PA-\d+)', basename, re.IGNORECASE)
    if match:
        return match.group(1).upper()
    # Fallback: nom sans extension
    return os.path.splitext(basename)[0]


# =============================================================================
# PCM READER
# =============================================================================

def _parse_pcm(filepath: str, zone_climatique: str = 'ZVN') -> Optional[EtudeComacMerged]:
    """Parse fichier PCM et retourne EtudeComacMerged partielle"""
    if not os.path.exists(filepath):
        return None
    
    etude = EtudeComacMerged(
        source_pcm=filepath,
        zone_climatique=zone_climatique
    )
    
    try:
        with open(filepath, 'r', encoding='iso-8859-1') as f:
            content = f.read()
        root = ET.fromstring(content)
    except ET.ParseError as e:
        etude.erreurs.append(f"PCM ParseError: {e}")
        return etude
    except Exception as e:
        etude.erreurs.append(f"PCM Error: {e}")
        return etude
    
    # Métadonnées
    etude.num_etude = _get_text(root, 'NumEtude') or _extract_num_etude(filepath)
    etude.commune = _get_text(root, 'Commune')
    
    # Hypothèses
    hypo_elem = root.find('Hypotheses')
    if hypo_elem is not None:
        etude.hypotheses = [h.text for h in hypo_elem.findall('Hypothese') if h.text]
    
    # Détection automatique zone climatique depuis hypothèses
    if etude.hypotheses and zone_climatique == 'ZVN':
        detected_zone = get_zone_vent_from_hypotheses(etude.hypotheses)
        if detected_zone != zone_climatique:
            etude.zone_climatique = detected_zone
    
    # Fallback: détection depuis code INSEE si disponible
    insee = _get_text(root, 'Insee')
    if insee and etude.zone_climatique == 'ZVN':
        insee_zone = get_zone_vent_from_insee(insee)
        if insee_zone == 'ZVF':
            etude.zone_climatique = 'ZVF'
    
    # Supports
    supports_elem = root.find('Supports')
    if supports_elem is not None:
        for supp_elem in supports_elem.findall('Support'):
            nom = _get_text(supp_elem, 'Nom')
            if not nom:
                continue
            support = SupportMerged(
                nom=nom,
                hauteur_totale=_get_float(supp_elem, 'Hauteur'),
                traverse_existante=_get_float(supp_elem, 'TraverseExistante1'),
                traverse_a_poser=_get_float(supp_elem, 'TraverseAPoser1') or _get_float(supp_elem, 'TraverseAPoser2'),
                portee_molle=_get_bool(supp_elem, 'PorteeMolle'),
                nature=_get_text(supp_elem, 'Nature'),
                classe=_get_text(supp_elem, 'Classe'),
                x=_get_float(supp_elem, 'X'),
                y=_get_float(supp_elem, 'Y'),
                etat=_get_text(supp_elem, 'Etat')
            )
            etude.supports[nom] = support
    
    # Lignes TCF (câbles FO)
    portees_max = PORTEES_MAX_ZVN if zone_climatique == 'ZVN' else PORTEES_MAX_ZVF
    
    tcf_elem = root.find('LignesTCF')
    if tcf_elem is not None:
        for ligne_elem in tcf_elem.findall('LigneTCF'):
            cable = _get_text(ligne_elem, 'Cable')
            capacite_fo = get_capacite_fo_from_code(cable)
            a_poser = _get_bool(ligne_elem, 'APoser')
            
            # Parse supports et traverses
            supports_list = []
            traverses_list = []
            supp_section = ligne_elem.find('Supports')
            if supp_section is not None:
                for child in supp_section:
                    if child.tag == 'Support' and child.text:
                        supports_list.append(child.text.strip())
                    elif child.tag == 'Traverse' and child.text:
                        traverses_list.append(_safe_int(child.text, 1))
            
            # Parse portées
            portees_list = []
            portees_section = ligne_elem.find('Portees')
            if portees_section is not None:
                for p in portees_section.findall('Portee'):
                    if p.text:
                        portees_list.append(_safe_float(p.text))
            
            # Créer segments (support[i] -> support[i+1] avec portee[i])
            portee_max = portees_max.get(capacite_fo, 50.0) if capacite_fo > 0 else 50.0
            
            for i, portee in enumerate(portees_list):
                if i < len(supports_list) - 1:
                    depassement = max(0, portee - portee_max)
                    valide = portee <= portee_max
                    
                    segment = SegmentFO(
                        support_depart=supports_list[i],
                        support_arrivee=supports_list[i + 1],
                        portee=portee,
                        capacite_fo=capacite_fo,
                        cable=cable,
                        traverse_type=traverses_list[i] if i < len(traverses_list) else 1,
                        a_poser=a_poser,
                        portee_max=portee_max,
                        depassement=depassement,
                        valide=valide
                    )
                    etude.segments_fo.append(segment)
                    
                    if not valide:
                        etude.nb_segments_invalides += 1
                        if depassement > etude.portee_max_depassement:
                            etude.portee_max_depassement = depassement
    
    return etude


# =============================================================================
# EXCEL READER
# =============================================================================

# Colonnes Excel COMAC (0-indexed)
EXCEL_COL_NUM_POTEAU = 0        # Col A
EXCEL_COL_HAUTEUR_TOTALE = 4    # Col E
EXCEL_COL_HAUTEUR_HORS_SOL = 6  # Col G
EXCEL_COL_CLASSE = 7            # Col H
EXCEL_COL_CONDUCTEUR = 13       # Col N
EXCEL_COL_FO_TYPE_LIGNE = 40    # Col AO
EXCEL_COL_LONGUEUR = 46         # Col AU


def _parse_excel(filepath: str) -> Dict[str, dict]:
    """Parse Excel COMAC, retourne dict par nom de poteau"""
    if not HAS_OPENPYXL or not os.path.exists(filepath):
        return {}
    
    data_by_support = {}
    
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        
        for row in ws.iter_rows(min_row=4, max_col=50, values_only=True):
            num_pot = row[EXCEL_COL_NUM_POTEAU] if row else None
            if not num_pot:
                continue
            
            nom = str(num_pot).strip().replace("BT ", "BT-")
            
            data_by_support[nom] = {
                'hauteur_hors_sol': _safe_float(row[EXCEL_COL_HAUTEUR_HORS_SOL] if len(row) > EXCEL_COL_HAUTEUR_HORS_SOL else None),
                'hauteur_totale': _safe_float(row[EXCEL_COL_HAUTEUR_TOTALE] if len(row) > EXCEL_COL_HAUTEUR_TOTALE else None),
                'conducteur': str(row[EXCEL_COL_CONDUCTEUR] or '') if len(row) > EXCEL_COL_CONDUCTEUR else '',
                'type_ligne_fo': str(row[EXCEL_COL_FO_TYPE_LIGNE] or '') if len(row) > EXCEL_COL_FO_TYPE_LIGNE else '',
                'longueur': _safe_float(row[EXCEL_COL_LONGUEUR] if len(row) > EXCEL_COL_LONGUEUR else None)
            }
        
        wb.close()
    except Exception as e:
        print(f"[COMAC_LOADER] Excel error {filepath}: {e}")
    
    return data_by_support


# =============================================================================
# LOADER PRINCIPAL
# =============================================================================

def load_etude_comac(
    pcm_path: Optional[str] = None,
    excel_path: Optional[str] = None,
    zone_climatique: str = 'ZVN'
) -> Optional[EtudeComacMerged]:
    """
    Charge une étude COMAC depuis PCM et/ou Excel.
    
    Args:
        pcm_path: Chemin fichier .pcm (prioritaire)
        excel_path: Chemin fichier Excel COMAC
        zone_climatique: 'ZVN' ou 'ZVF'
    
    Returns:
        EtudeComacMerged ou None si aucune source
    """
    if not pcm_path and not excel_path:
        return None
    
    # 1. Parse PCM (source principale)
    etude = None
    if pcm_path and os.path.exists(pcm_path):
        etude = _parse_pcm(pcm_path, zone_climatique)
    
    # 2. Parse Excel
    excel_data = {}
    if excel_path and os.path.exists(excel_path):
        excel_data = _parse_excel(excel_path)
        if etude:
            etude.source_excel = excel_path
    
    # 3. Si pas de PCM, créer étude depuis Excel
    if etude is None and excel_data:
        etude = EtudeComacMerged(
            num_etude=_extract_num_etude(excel_path),
            source_excel=excel_path,
            zone_climatique=zone_climatique
        )
        for nom, data in excel_data.items():
            etude.supports[nom] = SupportMerged(
                nom=nom,
                hauteur_totale=data.get('hauteur_totale', 0.0),
                hauteur_hors_sol=data.get('hauteur_hors_sol', 0.0),
                conducteur=data.get('conducteur', '')
            )
    
    # 4. Fusionner données Excel dans étude PCM
    if etude and excel_data:
        for nom, support in etude.supports.items():
            if nom in excel_data:
                support.hauteur_hors_sol = excel_data[nom].get('hauteur_hors_sol', 0.0)
                if not support.conducteur:
                    support.conducteur = excel_data[nom].get('conducteur', '')
        
        # Supports Excel non présents dans PCM
        for nom, data in excel_data.items():
            if nom not in etude.supports:
                etude.supports[nom] = SupportMerged(
                    nom=nom,
                    hauteur_hors_sol=data.get('hauteur_hors_sol', 0.0),
                    conducteur=data.get('conducteur', '')
                )
    
    return etude


def load_repertoire_comac(
    repertoire: str,
    zone_climatique: str = 'ZVN'
) -> Tuple[Dict[str, EtudeComacMerged], Dict[str, str]]:
    """
    Charge toutes les études COMAC d'un répertoire.
    Associe automatiquement PCM et Excel par num_etude.
    
    Returns:
        (dict études par num_etude, dict erreurs par fichier)
    """
    etudes = {}
    erreurs = {}
    
    # Collecter fichiers
    pcm_files = {}
    excel_files = {}
    
    for root_dir, _, files in os.walk(repertoire):
        for fname in files:
            fpath = os.path.join(root_dir, fname)
            num = _extract_num_etude(fname)
            
            if fname.lower().endswith('.pcm'):
                pcm_files[num] = fpath
            elif fname.lower().endswith('.xlsx') and 'exportcomac' in fname.lower():
                if '~$' not in fname:  # Ignore temp files
                    excel_files[num] = fpath
    
    # Fusionner par num_etude
    all_nums = set(pcm_files.keys()) | set(excel_files.keys())
    
    for num in all_nums:
        pcm_path = pcm_files.get(num)
        excel_path = excel_files.get(num)
        
        try:
            etude = load_etude_comac(pcm_path, excel_path, zone_climatique)
            if etude:
                etudes[num] = etude
        except Exception as e:
            erreurs[num] = str(e)
    
    return etudes, erreurs
