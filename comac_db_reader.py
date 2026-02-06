# -*- coding: utf-8 -*-
"""
Lecteur de la base de données COMAC.
Source: PostgreSQL (schéma 'comac' dans la BDD Auvergne)

Fournit les références câbles, supports, zones de vent par commune,
et les capacités FO possibles par référence câble.

Tables: commune, cables, supports, hypothese, armements, fleche,
        pincefusible, nappetv, cable_capacites_possibles
"""

import os
import threading
from dataclasses import dataclass
from typing import Dict, List, Optional

try:
    import psycopg2
    _HAS_PSYCOPG2 = True
except ImportError:
    _HAS_PSYCOPG2 = False

# Configuration
PG_SCHEMA = "comac"


# =============================================================================
# DATACLASSES
# =============================================================================

@dataclass
class CableReference:
    """Référence câble officielle COMAC"""
    nom: str
    section: float = 0.0
    diametre: float = 0.0
    masse_lineique: float = 0.0
    charge_rupture: float = 0.0
    volt: str = ""  # TV, BT, FO
    tension_zvn_agglo: float = 0.0
    tension_zvn_ecart: float = 0.0
    tension_zvf_agglo: float = 0.0
    tension_zvf_ecart: float = 0.0
    description: str = ""
    capacite_fo: int = 0  # Calculé depuis description
    fournisseur: str = ""  # ACOME, PRYSMIAN, SILEC


@dataclass
class SupportReference:
    """Référence support officielle COMAC"""
    nom: str
    nature: str = ""  # BE, BO, ME, AC
    classe: str = ""  # A, B, C, D
    effort_nominal: float = 0.0
    hauteur_totale: float = 0.0
    nom_capft: str = ""
    nom_gespot: str = ""


@dataclass
class CommuneInfo:
    """Info commune avec zones de vent"""
    insee: str
    nom: str
    departement: str
    zone1: int = 1  # <1991
    zone2: int = 1  # <2001
    zone3: int = 1  # <2008
    zone4: int = 1  # >=2008


@dataclass
class HypotheseClimatique:
    """Hypothèse climatique COMAC"""
    nom: str
    volt: str = ""
    description: str = ""
    temperature: float = 0.0
    pression_vent: float = 0.0
    complementaire: bool = False


# =============================================================================
# CACHE GLOBAL (singleton pattern) - CRIT-02: Thread-safe
# =============================================================================

_cache_lock = threading.Lock()
_cache_cables: Dict[str, CableReference] = {}
_cache_supports: Dict[str, SupportReference] = {}
_cache_communes: Dict[str, CommuneInfo] = {}
_cache_hypotheses: Dict[str, HypotheseClimatique] = {}
_cache_capacites_possibles: Dict[str, List[int]] = {}  # reference -> [capacités]
_cache_loaded: bool = False
_cache_source: str = ""  # "postgresql" or "gpkg"


# =============================================================================
# HELPERS
# =============================================================================

def _safe_float(value, default: float = 0.0) -> float:
    """Parse float (gère None et str)"""
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    val = str(value).strip()
    if not val:
        return default
    try:
        return float(val.replace(',', '.'))
    except ValueError:
        return default


def _safe_int(value, default: int = 0) -> int:
    """Parse int (gère None et str)"""
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
    except ValueError:
        return default


def _extract_capacite_fo(description: str, nom: str) -> int:
    """Extrait capacité FO depuis description ou nom câble"""
    import re
    
    # Pattern "XX fo" dans description
    match = re.search(r'(\d+)\s*fo', description.lower())
    if match:
        return int(match.group(1))
    
    # Pattern L1092-XX dans nom
    if 'L1092-11' in nom:
        return 6
    elif 'L1092-12' in nom:
        return 12
    elif 'L1092-13' in nom:
        return 36  # 18-36 fo modulo 6
    elif 'L1092-14' in nom:
        return 72  # 42-72 fo modulo 6
    elif 'L1092-15' in nom:
        return 144
    elif 'L1047-1' in nom:
        return 36  # 12-36 fo modulo 12
    elif 'L1047-2' in nom:
        return 72  # 48-72 fo modulo 12
    elif 'L1048' in nom:
        return 144  # 84-144 fo modulo 12
    elif 'L1083' in nom:
        return 1
    
    return 0


def _extract_fournisseur(nom: str) -> str:
    """Extrait fournisseur depuis nom câble"""
    nom_upper = nom.upper()
    if '-P' in nom_upper or 'PRYSMIAN' in nom_upper:
        return 'PRYSMIAN'
    elif '-A' in nom_upper or 'ACOME' in nom_upper:
        return 'ACOME'
    elif '-S' in nom_upper or 'SILEC' in nom_upper:
        return 'SILEC'
    return ''


def _get_pg_connection():
    """Tente d'obtenir une connexion PostgreSQL via les credentials QGIS."""
    if not _HAS_PSYCOPG2:
        return None
    
    try:
        from qgis.core import QgsSettings
        settings = QgsSettings()
        settings.beginGroup("PostgreSQL/connections")
        connections = settings.childGroups()
        settings.endGroup()
        
        for name in connections:
            settings.beginGroup(f"PostgreSQL/connections/{name}")
            host = settings.value("host", "")
            database = settings.value("database", "")
            port = int(settings.value("port", 5432))
            user = settings.value("username", "")
            password = settings.value("password", "")
            settings.endGroup()
            
            if host == "10.241.228.107" or "auvergne" in database.lower():
                conn = psycopg2.connect(
                    host=host, port=port, database=database,
                    user=user, password=password
                )
                # Vérifier que le schéma comac existe
                cur = conn.cursor()
                cur.execute(
                    "SELECT 1 FROM information_schema.schemata WHERE schema_name = %s",
                    (PG_SCHEMA,)
                )
                if cur.fetchone():
                    return conn
                else:
                    conn.close()
                    return None
    except Exception as e:
        print(f"[COMAC_DB] PostgreSQL non disponible: {e}")
    
    return None


def _query_pg(sql: str, params: tuple = ()) -> List[dict]:
    """Exécute requête sur PostgreSQL (schéma comac)."""
    rows = []
    conn = _get_pg_connection()
    if not conn:
        return rows
    
    try:
        cursor = conn.cursor()
        cursor.execute(sql, params)
        columns = [desc[0] for desc in cursor.description]
        rows = [dict(zip(columns, row)) for row in cursor.fetchall()]
    except Exception as e:
        print(f"[COMAC_DB] Erreur requête PG: {e}")
    finally:
        conn.close()
    
    return rows


def _query(sql_pg: str, params: tuple = ()) -> List[dict]:
    """Exécute une requête sur PostgreSQL (schéma comac)."""
    return _query_pg(sql_pg, params)


# =============================================================================
# LOADERS
# =============================================================================

def _load_cables() -> Dict[str, CableReference]:
    """Charge table cables depuis PostgreSQL"""
    cables = {}
    
    for row in _query(f"SELECT * FROM {PG_SCHEMA}.cables"):
        nom = row.get('nom', '')
        if not nom:
            continue
        
        description = row.get('description', '') or ''
        cable = CableReference(
            nom=nom,
            section=_safe_float(row.get('section_reelle')),
            diametre=_safe_float(row.get('diametre')),
            masse_lineique=_safe_float(row.get('masse_lineique')),
            charge_rupture=_safe_float(row.get('charge_rupture')),
            volt=(row.get('volt') or '').strip(),
            tension_zvn_agglo=_safe_float(row.get('tension_agglo')),
            tension_zvn_ecart=_safe_float(row.get('tension_ecart')),
            tension_zvf_agglo=_safe_float(row.get('tension_agglo_zvf')),
            tension_zvf_ecart=_safe_float(row.get('tension_ecart_zvf')),
            description=description,
            capacite_fo=_extract_capacite_fo(description, nom),
            fournisseur=_extract_fournisseur(nom)
        )
        cables[nom] = cable
    
    return cables


def _load_supports() -> Dict[str, SupportReference]:
    """Charge table supports depuis PostgreSQL"""
    supports = {}
    
    for row in _query(f"SELECT * FROM {PG_SCHEMA}.supports"):
        nom = row.get('nom', '')
        if not nom:
            continue
        
        support = SupportReference(
            nom=nom,
            nature=(row.get('nature') or '').strip(),
            classe=(row.get('classe') or '').strip(),
            effort_nominal=_safe_float(row.get('effort_nominal')),
            hauteur_totale=_safe_float(row.get('hauteur_totale')),
            nom_capft=(row.get('nom_capft') or '').strip(),
            nom_gespot=(row.get('nom_gespot') or '').strip()
        )
        supports[nom] = support
    
    return supports


def _load_communes() -> Dict[str, CommuneInfo]:
    """Charge table commune depuis PostgreSQL"""
    communes = {}
    
    for row in _query(f"SELECT * FROM {PG_SCHEMA}.commune"):
        insee = row.get('insee', '')
        if not insee:
            continue
        
        commune = CommuneInfo(
            insee=insee,
            nom=(row.get('nom') or '').strip(),
            departement=(row.get('dep') or '').strip(),
            zone1=_safe_int(row.get('zone1'), 1),
            zone2=_safe_int(row.get('zone2'), 1),
            zone3=_safe_int(row.get('zone3'), 1),
            zone4=_safe_int(row.get('zone4'), 1)
        )
        communes[insee] = commune
    
    return communes


def _load_hypotheses() -> Dict[str, HypotheseClimatique]:
    """Charge table hypothese depuis PostgreSQL"""
    hypotheses = {}
    
    for row in _query(f"SELECT * FROM {PG_SCHEMA}.hypothese"):
        nom = row.get('nom', '')
        if not nom:
            continue
        
        hypo = HypotheseClimatique(
            nom=nom,
            volt=(row.get('volt') or '').strip(),
            description=(row.get('description') or '').strip(),
            temperature=_safe_float(row.get('temperature')),
            pression_vent=_safe_float(row.get('pression_vent')),
            complementaire=_safe_int(row.get('complementaire')) == 1
        )
        hypotheses[nom] = hypo
    
    return hypotheses


def _load_capacites_possibles() -> Dict[str, List[int]]:
    """Charge table cable_capacites_possibles depuis PostgreSQL."""
    capas = {}
    
    rows = _query_pg(
        f"SELECT reference, capacite_fo FROM {PG_SCHEMA}.cable_capacites_possibles ORDER BY reference, capacite_fo"
    )
    
    for row in rows:
        ref = row.get('reference', '')
        capa = row.get('capacite_fo', 0)
        if ref:
            capas.setdefault(ref, []).append(capa)
    
    return capas


def _ensure_loaded():
    """Charge les données si pas encore fait - CRIT-02: Thread-safe"""
    global _cache_cables, _cache_supports, _cache_communes, _cache_hypotheses
    global _cache_capacites_possibles, _cache_loaded, _cache_source
    
    if _cache_loaded:
        return
    
    with _cache_lock:
        if _cache_loaded:
            return
        
        # Vérifier connexion PostgreSQL
        pg_conn = _get_pg_connection()
        if pg_conn:
            pg_conn.close()
            _cache_source = "postgresql"
            print(f"[COMAC_DB] Source: PostgreSQL (schéma {PG_SCHEMA})")
        else:
            print("[COMAC_DB] ERREUR: PostgreSQL non disponible")
            _cache_loaded = True
            return
        
        try:
            _cache_cables = _load_cables()
            _cache_supports = _load_supports()
            _cache_communes = _load_communes()
            _cache_hypotheses = _load_hypotheses()
            _cache_capacites_possibles = _load_capacites_possibles()
            
            print(
                f"[COMAC_DB] Chargé: {len(_cache_cables)} câbles, "
                f"{len(_cache_supports)} supports, {len(_cache_communes)} communes, "
                f"{len(_cache_capacites_possibles)} refs multi-capa"
            )
        except Exception as e:
            print(f"[COMAC_DB] ERR: Chargement BD échoué: {e}")
        
        _cache_loaded = True


def reload_database():
    """Force rechargement de la base - CRIT-02: Thread-safe"""
    global _cache_loaded, _cache_source
    with _cache_lock:
        _cache_loaded = False
        _cache_source = ""
    _ensure_loaded()


def get_source() -> str:
    """Retourne la source actuelle ('postgresql' ou 'gpkg' ou '')."""
    _ensure_loaded()
    return _cache_source


# =============================================================================
# API PUBLIQUE
# =============================================================================

def get_cable(nom: str) -> Optional[CableReference]:
    """Récupère référence câble par nom exact"""
    _ensure_loaded()
    return _cache_cables.get(nom)


def get_cable_capacite(code_cable: str) -> int:
    """
    Récupère capacité FO depuis code câble.
    Compatible avec codes type L1092-13-P ou noms complets.
    
    Args:
        code_cable: Code câble (ex: "L1092-13-P")
    
    Returns:
        Capacité FO ou 0 si non trouvé
    """
    _ensure_loaded()
    
    if not code_cable:
        return 0
    
    code_clean = str(code_cable).strip().upper()
    
    # Recherche exacte
    cable = _cache_cables.get(code_cable)
    if cable:
        return cable.capacite_fo
    
    # Recherche partielle
    for nom, cable in _cache_cables.items():
        if code_clean in nom.upper():
            return cable.capacite_fo
    
    # Fallback: extraction depuis pattern
    return _extract_capacite_fo('', code_clean)


def get_cable_capacites_possibles(code_cable: str) -> List[int]:
    """
    Retourne les capacités FO possibles pour un code câble.
    Source: table cable_capacites_possibles (PostgreSQL) si disponible.
    
    Ex: L1092-13-P → [24, 36], L1092-14-P → [48, 72]
    
    Args:
        code_cable: Code câble (ex: "L1092-13-P")
    
    Returns:
        Liste de capacités possibles, ou [capacite_unique] si pas de multi-capa
    """
    _ensure_loaded()
    
    if not code_cable:
        return [0]
    
    code_clean = str(code_cable).strip().upper()
    
    # Recherche dans le cache PostgreSQL (table cable_capacites_possibles)
    if _cache_capacites_possibles:
        # Recherche exacte
        if code_cable in _cache_capacites_possibles:
            return _cache_capacites_possibles[code_cable]
        # Recherche partielle (L1092-13-P dans "L1092-13-P xxx")
        for ref, capas in _cache_capacites_possibles.items():
            if ref.upper() in code_clean or code_clean in ref.upper():
                return capas
    
    # Fallback: capacité unique depuis get_cable_capacite
    capa = get_cable_capacite(code_cable)
    return [capa] if capa > 0 else [0]


def get_support(nom: str) -> Optional[SupportReference]:
    """Récupère référence support par nom"""
    _ensure_loaded()
    return _cache_supports.get(nom)


def get_commune(insee: str) -> Optional[CommuneInfo]:
    """Récupère info commune par code INSEE"""
    _ensure_loaded()
    return _cache_communes.get(str(insee).zfill(5))


def get_zone_vent_from_insee(insee: str, periode: int = 4) -> str:
    """
    Détermine ZVN ou ZVF depuis code INSEE et période.
    
    Args:
        insee: Code INSEE commune
        periode: Période réglementaire (1=<1991, 2=<2001, 3=<2008, 4=>=2008)
    
    Returns:
        'ZVN' (zone=1) ou 'ZVF' (zone>=2)
    """
    commune = get_commune(insee)
    if not commune:
        return 'ZVN'  # Défaut: ZVN
    
    zone_map = {1: commune.zone1, 2: commune.zone2, 3: commune.zone3, 4: commune.zone4}
    zone = zone_map.get(periode, commune.zone4)
    
    return 'ZVF' if zone >= 2 else 'ZVN'


def get_zone_vent_from_hypotheses(hypotheses: List[str]) -> str:
    """
    Détermine ZVN ou ZVF depuis hypothèses climatiques PCM.
    
    Args:
        hypotheses: Liste hypothèses (ex: ['A1', 'B1', 'DP1'])
    
    Returns:
        'ZVN' si A1, 'ZVF' si A2/A3/A4
    """
    _ensure_loaded()
    
    for h in hypotheses:
        h_clean = h.strip().upper()
        
        # ZVF: A2 (480 Pa), A3 (760 Pa cyclone), A4 (1200 Pa cyclone écart)
        if h_clean in ('A2', 'A3', 'A4') or h_clean.startswith('A3-') or h_clean.startswith('A4-'):
            return 'ZVF'
        
        # Vérification via base
        hypo = _cache_hypotheses.get(h_clean)
        if hypo and hypo.pression_vent >= 480:
            return 'ZVF'
    
    return 'ZVN'


def get_hypothese(nom: str) -> Optional[HypotheseClimatique]:
    """Récupère hypothèse climatique par nom"""
    _ensure_loaded()
    return _cache_hypotheses.get(nom.strip())


def list_cables_fo() -> List[CableReference]:
    """Liste tous les câbles FO"""
    _ensure_loaded()
    return [c for c in _cache_cables.values() if c.volt == 'FO']


def list_cables_by_capacite(capacite: int) -> List[CableReference]:
    """Liste câbles FO par capacité"""
    _ensure_loaded()
    return [c for c in _cache_cables.values() if c.capacite_fo == capacite]


def get_all_cables() -> Dict[str, CableReference]:
    """Retourne tous les câbles (copie)"""
    _ensure_loaded()
    return dict(_cache_cables)


def get_all_supports() -> Dict[str, SupportReference]:
    """Retourne tous les supports (copie)"""
    _ensure_loaded()
    return dict(_cache_supports)


def get_all_communes() -> Dict[str, CommuneInfo]:
    """Retourne toutes les communes (copie)"""
    _ensure_loaded()
    return dict(_cache_communes)


# =============================================================================
# MAPPING CAPACITÉS POUR COMPATIBILITÉ PoliceC6
# =============================================================================

# Mapping codes câbles -> capacités (alimenté par BD officielle)
def get_codes_cable_capacites() -> Dict[str, int]:
    """
    Retourne mapping code câble -> capacité FO.
    Compatible avec format CODES_CABLE_PRYSMIAN de security_rules.py
    """
    _ensure_loaded()
    
    mapping = {}
    for nom, cable in _cache_cables.items():
        if cable.capacite_fo > 0:
            mapping[nom] = cable.capacite_fo
    
    return mapping
