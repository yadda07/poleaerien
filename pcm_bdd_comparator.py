# -*- coding: utf-8 -*-
"""
Comparaison PCM vs BDD PostgreSQL.

Compare les donnees des fichiers PCM (sous-traitant) avec la BDD
(infra_pt_pot, fddcpiax) pour detecter les ecarts.

Architecture:
- Fonctions pures, sans etat, thread-safe
- Entrees: dataclasses PCM (pcm_parser) + donnees BDD (dicts/listes)
- Sorties: dataclasses resultats
"""

import math
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple

try:
    from .core_utils import normalize_appui_num  # noqa: F401 - used by consumers
except ImportError:
    from core_utils import normalize_appui_num  # noqa: F401

try:
    from .pcm_parser import EtudePCM, Support, LigneTCF
except ImportError:
    from pcm_parser import EtudePCM, Support, LigneTCF


# =============================================================================
# CONSTANTES
# =============================================================================

NATURE_TO_INF_TYPE = {
    'BE': 'POT-BT',
    'BO': 'POT-BT',
    'FT': 'POT-FT',
}

COORD_TOLERANCE_M = 5.0
PORTEE_TOLERANCE_PCT = 15.0


# =============================================================================
# DATACLASSES RESULTATS
# =============================================================================

@dataclass
class SupportMatch:
    """Resultat matching d'un support PCM avec infra_pt_pot."""
    nom_pcm: str
    inf_num_bdd: str = ""
    noe_codext_bdd: str = ""
    matched: bool = False
    match_method: str = ""       # 'inf_num', 'noe_codext', 'spatial', ''
    ecart_coord_m: float = -1.0
    type_pcm: str = ""           # BE, FT, BO
    type_bdd: str = ""           # POT-BT, POT-FT, POT-AC
    type_coherent: bool = True
    etat_pcm: str = ""
    etat_bdd: str = ""
    coord_pcm: Tuple[float, float] = (0.0, 0.0)
    coord_bdd: Tuple[float, float] = (0.0, 0.0)
    non_calcule: bool = False
    facade: bool = False


@dataclass
class CableMatch:
    """Resultat matching d'un cable PCM avec fddcpiax."""
    cable_pcm: str
    capacite_pcm: int = 0
    capacite_bdd: int = 0
    capacite_coherente: bool = True
    nb_segments_pcm: int = 0
    nb_segments_bdd: int = 0
    a_poser: bool = False
    portees_pcm: List[float] = field(default_factory=list)
    portees_bdd: List[float] = field(default_factory=list)
    ecarts_portees: List[dict] = field(default_factory=list)
    matched: bool = False


@dataclass
class EtudeComparison:
    """Resultat comparaison d'une etude PCM vs BDD."""
    num_etude: str
    commune: str = ""
    insee: str = ""
    supports_ok: List[SupportMatch] = field(default_factory=list)
    supports_absents_bdd: List[SupportMatch] = field(default_factory=list)
    supports_hors_perimetre: List[SupportMatch] = field(default_factory=list)
    supports_type_ko: List[SupportMatch] = field(default_factory=list)
    supports_coord_ko: List[SupportMatch] = field(default_factory=list)
    cables_ok: List[CableMatch] = field(default_factory=list)
    cables_absents_bdd: List[CableMatch] = field(default_factory=list)
    cables_capacite_ko: List[CableMatch] = field(default_factory=list)
    cables_portee_ko: List[CableMatch] = field(default_factory=list)
    erreurs: List[str] = field(default_factory=list)


@dataclass
class PCMvsBDDResult:
    """Resultat agrege de toutes les comparaisons PCM vs BDD."""
    etudes: List[EtudeComparison] = field(default_factory=list)
    nb_etudes: int = 0
    nb_supports_total: int = 0
    nb_supports_ok: int = 0
    nb_supports_absent: int = 0
    nb_supports_type_ko: int = 0
    nb_supports_coord_ko: int = 0
    nb_cables_total: int = 0
    nb_cables_ok: int = 0
    nb_cables_absent: int = 0
    nb_cables_capacite_ko: int = 0
    nb_cables_portee_ko: int = 0


# =============================================================================
# NORMALISATION NOMS SUPPORTS PCM
# =============================================================================

_PCM_PREFIXES_NC = ('NCBT', 'NCHTA', 'NC')
_PCM_PREFIXES_SKIP = ('FA', 'POTELET', 'AEOP')


def _normalize_pcm_support_name(nom_pcm: str) -> str:
    """Normalise un nom de support PCM pour matching avec noe_codext BDD.

    Exemples:
        BT0030     -> BT0030 (inchange)
        NCBT0029   -> BT0029 (strip NC, garde BT)
        NC106500   -> 106500
        NCHTA0069  -> HTA0069
        NC62230    -> 62230
        FT688752   -> FT688752 (inchange)
        E000058    -> E000058 (inchange)

    Returns:
        Nom normalise, en conservant le prefixe significatif.
    """
    if not nom_pcm:
        return ""
    nom = nom_pcm.strip()
    if nom.startswith('NCBT'):
        return 'BT' + nom[4:]
    if nom.startswith('NCHTA'):
        return 'HTA' + nom[5:]
    if nom.startswith('NC') and len(nom) > 2:
        return nom[2:]
    return nom


def _is_matching_candidate(nom_pcm: str) -> bool:
    """Verifie si un support PCM doit etre matche avec la BDD.

    Les facades, potelets et AEOP ne sont pas dans infra_pt_pot.
    """
    upper = nom_pcm.upper()
    return not any(upper.startswith(p) for p in _PCM_PREFIXES_SKIP)


# =============================================================================
# CONSTRUCTION INDEX BDD
# =============================================================================

def build_bdd_index(poteaux_bdd: List[dict]) -> dict:
    """Construit un index multi-cle pour le matching rapide.

    Args:
        poteaux_bdd: Liste de dicts avec au minimum
            {inf_num, noe_codext, inf_type, etat, x, y}

    Returns:
        {
            'by_inf_num': {num_normalise: [pot_dict, ...]},
            'by_codext': {codext_normalise: [pot_dict, ...]},
            'all_inf_nums': set(num_normalise),
        }
    """
    if not poteaux_bdd:
        return {'by_inf_num': {}, 'by_codext': {}, 'all_inf_nums': set()}

    by_inf_num = {}
    by_codext = {}
    all_inf_nums = set()

    for pot in poteaux_bdd:
        inf_num = pot.get('inf_num', '') or ''
        codext = pot.get('noe_codext', '') or ''

        if inf_num:
            norm = _normalize_bdd_inf_num(inf_num)
            by_inf_num.setdefault(norm, []).append(pot)
            all_inf_nums.add(norm)

        if codext:
            norm_c = _normalize_bdd_codext(codext)
            by_codext.setdefault(norm_c, []).append(pot)

    return {
        'by_inf_num': by_inf_num,
        'by_codext': by_codext,
        'all_inf_nums': all_inf_nums,
    }


def _normalize_bdd_inf_num(inf_num: str) -> str:
    """Normalise inf_num BDD pour matching.

    Exemples: E000001/03112 -> E000001, 0093824/03278 -> 93824
    """
    if '/' in inf_num:
        inf_num = inf_num.split('/')[0]
    inf_num = inf_num.strip()
    if inf_num.startswith('E'):
        return inf_num
    if inf_num.isdigit():
        return inf_num.lstrip('0') or '0'
    return inf_num


def _normalize_bdd_codext(codext: str) -> str:
    """Normalise noe_codext BDD pour matching.

    Exemples: BT0393/03112 -> BT0393, FT0305/03294 -> FT0305
    """
    if '/' in codext:
        codext = codext.split('/')[0]
    return codext.strip()


# =============================================================================
# COMPARAISON SUPPORTS
# =============================================================================

def comparer_supports(
    etude: EtudePCM,
    index_bdd: dict,
    coord_tolerance: float = COORD_TOLERANCE_M,
) -> Tuple[List[SupportMatch], List[SupportMatch], List[SupportMatch], List[SupportMatch]]:
    """Compare les supports d'une etude PCM avec la BDD.

    Args:
        etude: Etude PCM parsee
        index_bdd: Index retourne par build_bdd_index()
        coord_tolerance: Tolerance coordonnees en metres

    Returns:
        (ok, absents_bdd, type_ko, coord_ko)
    """
    if not etude.supports:
        return [], [], [], []
    if not index_bdd:
        index_bdd = {'by_inf_num': {}, 'by_codext': {}, 'all_inf_nums': set()}

    ok = []
    absents_bdd = []
    type_ko = []
    coord_ko = []

    by_inf_num = index_bdd.get('by_inf_num', {})
    by_codext = index_bdd.get('by_codext', {})

    for nom_pcm, support in etude.supports.items():
        if not _is_matching_candidate(nom_pcm):
            continue

        match = _match_single_support(
            nom_pcm, support, by_inf_num, by_codext, coord_tolerance
        )

        if not match.matched:
            absents_bdd.append(match)
            continue

        if not match.type_coherent:
            type_ko.append(match)
        elif match.ecart_coord_m > coord_tolerance:
            coord_ko.append(match)
        else:
            ok.append(match)

    return ok, absents_bdd, type_ko, coord_ko


def _match_single_support(
    nom_pcm: str,
    support: Support,
    by_inf_num: dict,
    by_codext: dict,
    coord_tolerance: float,
) -> SupportMatch:
    """Tente de matcher un support PCM avec la BDD.

    Strategie:
    1. Match par noe_codext (nom terrain BT0030 -> BT0030/commune)
    2. Match par inf_num (E000xxx -> E000xxx)
    3. Match par inf_num numerique (NC106500 -> 106500 -> 0106500)
    """
    result = SupportMatch(
        nom_pcm=nom_pcm,
        type_pcm=support.nature,
        etat_pcm=support.etat,
        coord_pcm=(support.x, support.y),
        non_calcule=support.non_calcule,
        facade=support.facade,
    )

    pcm_norm = _normalize_pcm_support_name(nom_pcm)

    # Strategie 1: match par noe_codext (unique)
    candidates = by_codext.get(pcm_norm, [])
    if len(candidates) == 1:
        _fill_match(result, candidates[0], 'noe_codext')
        return result

    # Strategie 1b: multi-candidats codext -- disambiguer par spatial si coords PCM
    if len(candidates) > 1 and support.x > 100000 and support.y > 6000000:
        nearest = _pick_nearest(support.x, support.y, candidates)
        if nearest is not None:
            _fill_match(result, nearest, 'noe_codext')
            return result

    # Strategie 2: match par inf_num (E-prefix)
    if nom_pcm.startswith('E'):
        candidates = by_inf_num.get(nom_pcm, [])
        if len(candidates) == 1:
            _fill_match(result, candidates[0], 'inf_num')
            return result

    # Strategie 3: match numerique (NC106500 -> 106500)
    if pcm_norm.isdigit():
        norm_stripped = pcm_norm.lstrip('0') or '0'
        candidates = by_inf_num.get(norm_stripped, [])
        if len(candidates) == 1:
            _fill_match(result, candidates[0], 'inf_num')
            return result

    # Strategie 4: match spatial global (dernier recours)
    if support.x > 100000 and support.y > 6000000:
        spatial_match = _find_nearest_bdd_pot(
            support.x, support.y, by_codext, coord_tolerance
        )
        if spatial_match is not None:
            _fill_match(result, spatial_match, 'spatial')
            return result

    return result


def _fill_match(
    result: SupportMatch,
    pot_bdd: dict,
    method: str,
) -> None:
    """Remplit un SupportMatch avec les donnees BDD."""
    if not pot_bdd:
        return
    result.matched = True
    result.match_method = method
    result.inf_num_bdd = pot_bdd.get('inf_num', '')
    result.noe_codext_bdd = pot_bdd.get('noe_codext', '')
    result.type_bdd = pot_bdd.get('inf_type', '')
    result.etat_bdd = pot_bdd.get('etat', '') or ''
    bdd_x = pot_bdd.get('x', 0.0) or 0.0
    bdd_y = pot_bdd.get('y', 0.0) or 0.0
    result.coord_bdd = (bdd_x, bdd_y)

    pcm_x, pcm_y = result.coord_pcm
    if pcm_x > 100000 and pcm_y > 6000000 and bdd_x > 100000 and bdd_y > 6000000:
        result.ecart_coord_m = math.hypot(pcm_x - bdd_x, pcm_y - bdd_y)

    expected_type = NATURE_TO_INF_TYPE.get(result.type_pcm, '')
    if expected_type and result.type_bdd:
        result.type_coherent = result.type_bdd == expected_type


def _pick_nearest(
    x: float, y: float,
    candidates: List[dict],
) -> Optional[dict]:
    """Parmi une liste de candidats BDD, retourne le plus proche spatialement.

    Utilise pour disambiguer quand plusieurs poteaux BDD matchent le meme codext.
    """
    best_pot = None
    best_dist = float('inf')
    for pot in candidates:
        px = pot.get('x', 0.0) or 0.0
        py = pot.get('y', 0.0) or 0.0
        if px < 100000 or py < 6000000:
            continue
        dist = math.hypot(x - px, y - py)
        if dist < best_dist:
            best_dist = dist
            best_pot = pot
    return best_pot


def _find_nearest_bdd_pot(
    x: float, y: float,
    by_codext: dict,
    tolerance: float,
) -> Optional[dict]:
    """Recherche spatiale brute dans by_codext. O(n) acceptable car < 1000 poteaux par SRO."""
    best_pot = None
    best_dist = tolerance + 1.0
    for pots in by_codext.values():
        for pot in pots:
            px = pot.get('x', 0.0) or 0.0
            py = pot.get('y', 0.0) or 0.0
            if px < 100000 or py < 6000000:
                continue
            dist = math.hypot(x - px, y - py)
            if dist <= tolerance and dist < best_dist:
                best_dist = dist
                best_pot = pot
    return best_pot


# =============================================================================
# COMPARAISON CABLES
# =============================================================================

def comparer_cables(
    etude: EtudePCM,
    cables_bdd: List[dict],
    appuis_bdd: dict,
    tolerance_pct: float = PORTEE_TOLERANCE_PCT,
) -> Tuple[List[CableMatch], List[CableMatch], List[CableMatch], List[CableMatch]]:
    """Compare les cables FO d'une etude PCM avec fddcpiax.

    Args:
        etude: Etude PCM parsee
        cables_bdd: Segments depuis fddcpiax (filtre CDI, posemode 1/2)
            chaque dict: {gid, length, cab_capa, geom_start_x, geom_start_y, geom_end_x, geom_end_y}
        appuis_bdd: Index retourne par build_bdd_index() (pour resoudre les endpoints)
        tolerance_pct: Tolerance ecart portee en pourcentage

    Returns:
        (ok, absents_bdd, capacite_ko, portee_ko)
    """
    if not cables_bdd or not etude.lignes_tcf:
        return [], [], [], []

    ok = []
    absents_bdd = []
    capacite_ko = []
    portee_ko = []

    cables_par_appui_bdd = _grouper_cables_par_appui(cables_bdd, appuis_bdd or {})

    for ligne in etude.lignes_tcf:
        if not ligne.a_poser:
            continue
        if ligne.capacite_fo == 0:
            continue

        match = _match_cable(ligne, cables_par_appui_bdd, tolerance_pct)

        if not match.matched:
            absents_bdd.append(match)
        elif not match.capacite_coherente:
            capacite_ko.append(match)
        elif match.ecarts_portees:
            portee_ko.append(match)
        else:
            ok.append(match)

    return ok, absents_bdd, capacite_ko, portee_ko


def _grouper_cables_par_appui(
    cables_bdd: List[dict],
    appuis_bdd: dict,
) -> Dict[str, List[dict]]:
    """Groupe les segments BDD par appui le plus proche de chaque endpoint.

    Returns:
        {nom_appui_normalise: [segment_dict, ...]}
    """
    result = {}
    by_codext = appuis_bdd.get('by_codext', {})

    for seg in cables_bdd:
        start_x = seg.get('geom_start_x', 0.0)
        start_y = seg.get('geom_start_y', 0.0)
        end_x = seg.get('geom_end_x', 0.0)
        end_y = seg.get('geom_end_y', 0.0)

        for px, py in [(start_x, start_y), (end_x, end_y)]:
            if px < 100000 or py < 6000000:
                continue
            nearest = _find_nearest_bdd_pot(px, py, by_codext, 2.0)
            if nearest:
                codext = _normalize_bdd_codext(nearest.get('noe_codext', ''))
                if codext:
                    result.setdefault(codext, []).append(seg)

    return result


def _match_cable(
    ligne: LigneTCF,
    cables_par_appui: Dict[str, List[dict]],
    tolerance_pct: float,
) -> CableMatch:
    """Tente de matcher un cable PCM avec les segments BDD via les supports."""
    match = CableMatch(
        cable_pcm=ligne.cable,
        capacite_pcm=ligne.capacite_fo,
        a_poser=ligne.a_poser,
        nb_segments_pcm=len(ligne.portees),
        portees_pcm=list(ligne.portees),
    )

    if len(ligne.supports) < 2:
        return match

    # Chercher des segments BDD qui touchent les appuis PCM
    segments_trouves = set()
    for support_name in ligne.supports:
        pcm_norm = _normalize_pcm_support_name(support_name)
        segs = cables_par_appui.get(pcm_norm, [])
        for seg in segs:
            segments_trouves.add(seg.get('gid', 0))

    if not segments_trouves:
        return match

    match.matched = True

    # Verifier capacite
    capas_bdd = set()
    longueurs_bdd = []
    for seg in _collect_segments(cables_par_appui, ligne.supports):
        capas_bdd.add(seg.get('cab_capa', 0))
        longueurs_bdd.append(seg.get('length', 0.0))

    match.nb_segments_bdd = len(longueurs_bdd)
    match.portees_bdd = longueurs_bdd

    if capas_bdd:
        match.capacite_bdd = max(capas_bdd)
        match.capacite_coherente = ligne.capacite_fo in capas_bdd

    # Comparer portees segment par segment
    match.ecarts_portees = _comparer_portees_liste(
        ligne.portees, longueurs_bdd, tolerance_pct
    )

    return match


def _collect_segments(
    cables_par_appui: Dict[str, List[dict]],
    supports_pcm: List[str],
) -> List[dict]:
    """Collecte les segments BDD uniques touchant les appuis PCM."""
    seen_gids = set()
    result = []
    for support_name in supports_pcm:
        pcm_norm = _normalize_pcm_support_name(support_name)
        for seg in cables_par_appui.get(pcm_norm, []):
            gid = seg.get('gid', 0)
            if gid not in seen_gids:
                seen_gids.add(gid)
                result.append(seg)
    return result


def _comparer_portees_liste(
    portees_pcm: List[float],
    portees_bdd: List[float],
    tolerance_pct: float,
) -> List[dict]:
    """Compare deux listes de portees, retourne les ecarts hors tolerance.

    Matching par valeur la plus proche (greedy).
    """
    ecarts = []
    bdd_used = set()

    for pcm_val in portees_pcm:
        if pcm_val <= 0:
            continue
        best_idx = -1
        best_diff = float('inf')
        for j, bdd_val in enumerate(portees_bdd):
            if j in bdd_used or bdd_val <= 0:
                continue
            diff = abs(pcm_val - bdd_val)
            if diff < best_diff:
                best_diff = diff
                best_idx = j

        if best_idx >= 0:
            bdd_val = portees_bdd[best_idx]
            bdd_used.add(best_idx)
            pct = (best_diff / pcm_val) * 100 if pcm_val > 0 else 0
            if pct > tolerance_pct:
                ecarts.append({
                    'portee_pcm': pcm_val,
                    'portee_bdd': bdd_val,
                    'ecart_m': round(best_diff, 1),
                    'ecart_pct': round(pct, 1),
                })
        else:
            ecarts.append({
                'portee_pcm': pcm_val,
                'portee_bdd': 0.0,
                'ecart_m': pcm_val,
                'ecart_pct': 100.0,
            })

    return ecarts


# =============================================================================
# ORCHESTRATEUR
# =============================================================================

def comparer_etude_pcm_vs_bdd(
    etude: EtudePCM,
    poteaux_bdd: List[dict],
    cables_bdd: List[dict],
    coord_tolerance: float = COORD_TOLERANCE_M,
    portee_tolerance_pct: float = PORTEE_TOLERANCE_PCT,
) -> EtudeComparison:
    """Compare une etude PCM complete avec la BDD.

    Args:
        etude: Etude PCM parsee
        poteaux_bdd: Liste de dicts depuis infra_pt_pot
            {inf_num, noe_codext, inf_type, etat, x, y}
        cables_bdd: Segments depuis fddcpiax (filtre CDI, posemode 1/2)
            {gid, length, cab_capa, geom_start_x, geom_start_y, geom_end_x, geom_end_y}
        coord_tolerance: Tolerance coordonnees en metres
        portee_tolerance_pct: Tolerance portee en pourcentage

    Returns:
        EtudeComparison avec tous les resultats
    """
    index_bdd = build_bdd_index(poteaux_bdd)
    return _comparer_etude_with_index(
        etude, index_bdd, cables_bdd, coord_tolerance, portee_tolerance_pct
    )


def _comparer_etude_with_index(
    etude: EtudePCM,
    index_bdd: dict,
    cables_bdd: List[dict],
    coord_tolerance: float = COORD_TOLERANCE_M,
    portee_tolerance_pct: float = PORTEE_TOLERANCE_PCT,
) -> EtudeComparison:
    """Compare une etude PCM avec un index BDD pre-construit."""
    comp = EtudeComparison(
        num_etude=etude.num_etude,
        commune=etude.commune,
        insee=etude.insee,
    )

    # Supports
    s_ok, s_absent, s_type_ko, s_coord_ko = comparer_supports(
        etude, index_bdd, coord_tolerance
    )
    comp.supports_ok = s_ok
    comp.supports_absents_bdd = s_absent
    comp.supports_type_ko = s_type_ko
    comp.supports_coord_ko = s_coord_ko

    # Cables
    c_ok, c_absent, c_capa_ko, c_portee_ko = comparer_cables(
        etude, cables_bdd, index_bdd, portee_tolerance_pct
    )
    comp.cables_ok = c_ok
    comp.cables_absents_bdd = c_absent
    comp.cables_capacite_ko = c_capa_ko
    comp.cables_portee_ko = c_portee_ko

    return comp


def comparer_batch_pcm_vs_bdd(
    etudes: Dict[str, EtudePCM],
    poteaux_bdd: List[dict],
    cables_bdd: List[dict],
    coord_tolerance: float = COORD_TOLERANCE_M,
    portee_tolerance_pct: float = PORTEE_TOLERANCE_PCT,
) -> PCMvsBDDResult:
    """Compare toutes les etudes PCM d'un batch avec la BDD.

    Args:
        etudes: {nom_etude: EtudePCM}
        poteaux_bdd: Liste de dicts depuis infra_pt_pot
        cables_bdd: Segments depuis fddcpiax

    Returns:
        PCMvsBDDResult avec totaux et detail par etude
    """
    if not etudes:
        return PCMvsBDDResult()

    result = PCMvsBDDResult(nb_etudes=len(etudes))

    # Index BDD construit UNE SEULE FOIS pour toutes les etudes
    index_bdd = build_bdd_index(poteaux_bdd)

    for _nom, etude in etudes.items():
        comp = _comparer_etude_with_index(
            etude, index_bdd, cables_bdd,
            coord_tolerance, portee_tolerance_pct,
        )
        result.etudes.append(comp)

        result.nb_supports_total += (
            len(comp.supports_ok) + len(comp.supports_absents_bdd)
            + len(comp.supports_type_ko) + len(comp.supports_coord_ko)
        )
        result.nb_supports_ok += len(comp.supports_ok)
        result.nb_supports_absent += len(comp.supports_absents_bdd)
        result.nb_supports_type_ko += len(comp.supports_type_ko)
        result.nb_supports_coord_ko += len(comp.supports_coord_ko)

        result.nb_cables_total += (
            len(comp.cables_ok) + len(comp.cables_absents_bdd)
            + len(comp.cables_capacite_ko) + len(comp.cables_portee_ko)
        )
        result.nb_cables_ok += len(comp.cables_ok)
        result.nb_cables_absent += len(comp.cables_absents_bdd)
        result.nb_cables_capacite_ko += len(comp.cables_capacite_ko)
        result.nb_cables_portee_ko += len(comp.cables_portee_ko)

    return result


# =============================================================================
# PHASE 4: VALIDATION DONNEES MECANIQUES PCM vs TABLES COMAC
# =============================================================================

@dataclass
class ValidationMecanique:
    """Resultat validation donnees mecaniques d'une etude PCM."""
    hypotheses_inconnues: List[str] = field(default_factory=list)
    armements_inconnus: List[dict] = field(default_factory=list)
    cables_inconnus: List[str] = field(default_factory=list)
    supports_catalogue_ko: List[dict] = field(default_factory=list)
    zone_climatique: str = ""


def valider_hypotheses(
    etude: EtudePCM,
    hypotheses_bdd: List[str],
) -> List[str]:
    """Verifie que les hypotheses PCM existent dans comac.hypothese.

    Args:
        etude: Etude PCM parsee
        hypotheses_bdd: Liste des noms valides depuis comac.hypothese

    Returns:
        Liste des hypotheses PCM inconnues dans le referentiel
    """
    noms_valides = {h.upper().strip() for h in hypotheses_bdd}
    return [h for h in etude.hypotheses if h.upper().strip() not in noms_valides]


def valider_armements(
    etude: EtudePCM,
    armements_bdd: List[str],
) -> List[dict]:
    """Verifie que les armements BT PCM existent dans comac.armements.

    Args:
        etude: Etude PCM parsee
        armements_bdd: Liste des noms valides depuis comac.armements

    Returns:
        Liste de dicts {support, nom_armement, conducteur} pour armements inconnus
    """
    noms_valides = {a.upper().strip() for a in armements_bdd}
    inconnus = []
    for ligne in etude.lignes_bt:
        for arm in ligne.armements:
            if not arm.nom_armement:
                continue
            if arm.nom_armement.upper().strip() not in noms_valides:
                inconnus.append({
                    'support': arm.support,
                    'nom_armement': arm.nom_armement,
                    'conducteur': ligne.conducteur,
                })
    return inconnus


def valider_cables_referentiel(
    etude: EtudePCM,
    cables_bdd: List[str],
) -> List[str]:
    """Verifie que les references cables PCM existent dans comac.cables.

    Args:
        etude: Etude PCM parsee
        cables_bdd: Liste des noms valides depuis comac.cables

    Returns:
        Liste des references cables PCM inconnues
    """
    noms_valides = {c.upper().strip() for c in cables_bdd}
    inconnus = []
    seen = set()
    for ligne in etude.lignes_tcf:
        ref = ligne.cable.upper().strip()
        if ref and ref not in noms_valides and ref not in seen:
            inconnus.append(ligne.cable)
            seen.add(ref)
    for ligne in etude.lignes_bt:
        ref = ligne.conducteur.upper().strip()
        if ref and ref not in noms_valides and ref not in seen:
            inconnus.append(ligne.conducteur)
            seen.add(ref)
    return inconnus


def valider_supports_catalogue(
    etude: EtudePCM,
    supports_bdd: List[dict],
) -> List[dict]:
    """Verifie que les supports PCM correspondent au catalogue comac.supports.

    Cherche par combinaison (hauteur, classe, effort) -> nom catalogue.

    Args:
        etude: Etude PCM parsee
        supports_bdd: Liste de dicts depuis comac.supports
            {nom, nature, classe, effort_nominal, hauteur_totale}

    Returns:
        Liste de dicts {nom_pcm, hauteur, classe, effort, raison} pour KO
    """
    catalogue = set()
    for s in supports_bdd:
        key = (
            s.get('hauteur_totale', 0.0),
            (s.get('classe', '') or '').upper().strip(),
            s.get('effort_nominal', 0.0),
        )
        catalogue.add(key)

    ko = []
    for nom_pcm, support in etude.supports.items():
        if support.non_calcule or support.facade:
            continue
        key = (support.hauteur, support.classe.upper().strip(), support.effort)
        if key not in catalogue and support.hauteur > 0:
            ko.append({
                'nom_pcm': nom_pcm,
                'hauteur': support.hauteur,
                'classe': support.classe,
                'effort': support.effort,
                'nature': support.nature,
            })
    return ko


def determiner_zone_climatique(
    insee: str,
    communes_bdd: List[dict],
) -> str:
    """Determine la zone climatique depuis le code INSEE via comac.commune.

    Args:
        insee: Code INSEE de l'etude PCM
        communes_bdd: Liste de dicts {insee, zone1, zone2, zone3, zone4}

    Returns:
        'ZVN' ou 'ZVF' ou '' si inconnu
    """
    if not insee:
        return ""
    insee_clean = insee.strip().zfill(5)
    for commune in communes_bdd:
        if (commune.get('insee', '') or '').strip().zfill(5) == insee_clean:
            zone1 = commune.get('zone1', 1)
            if zone1 and zone1 >= 2:
                return 'ZVF'
            return 'ZVN'
    return ""


def valider_mecanique_etude(
    etude: EtudePCM,
    hypotheses_bdd: List[str],
    armements_bdd: List[str],
    cables_catalogue_bdd: List[str],
    supports_catalogue_bdd: List[dict],
    communes_bdd: List[dict],
) -> ValidationMecanique:
    """Validation mecanique complete d'une etude PCM.

    Args:
        etude: Etude PCM parsee
        hypotheses_bdd: Noms valides depuis comac.hypothese
        armements_bdd: Noms valides depuis comac.armements
        cables_catalogue_bdd: Noms valides depuis comac.cables
        supports_catalogue_bdd: Dicts depuis comac.supports
        communes_bdd: Dicts depuis comac.commune

    Returns:
        ValidationMecanique avec toutes les anomalies
    """
    return ValidationMecanique(
        hypotheses_inconnues=valider_hypotheses(etude, hypotheses_bdd),
        armements_inconnus=valider_armements(etude, armements_bdd),
        cables_inconnus=valider_cables_referentiel(etude, cables_catalogue_bdd),
        supports_catalogue_ko=valider_supports_catalogue(etude, supports_catalogue_bdd),
        zone_climatique=determiner_zone_climatique(etude.insee, communes_bdd),
    )
