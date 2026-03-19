# -*- coding: utf-8 -*-
"""
Lecture et normalisation des fichiers CSV GESPOT.

Responsabilité unique : parser les CSV GESPOT et appliquer les règles
métier BR-01 à BR-06 (+ BR-03bis/ter, BR-04bis) pour produire des
enregistrements comparables aux champs C6.

Zéro dépendance QGIS — thread-safe.
"""

import csv
import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple

from .core_utils import normalize_appui_num

_GESPOT_HEADER_MARKERS = {'NUM', 'CARAC1', 'CENTRE'}


# ===========================================================================
#  CONSTANTES — indices de colonnes CSV GESPOT (0-bases)
# ===========================================================================

_COL_NUM = 1
_COL_VOIE = 6
_COL_NUM_VOIE = 7
_COL_CENTRE = 8
_COL_TYPE = 12
_COL_STRATEGIQUE = 13
_COL_PRES_ELECT = 14
_COL_SUPPORT_PC = 15
_COL_INACC = 17
_COL_RECALAGE = 19
_COL_CARAC = [21, 22, 23, 24, 25]
_COL_ENV = [26, 27, 28, 29, 30]
_COL_ETAT_PT1 = 33
_COL_DECLASS_PT1 = 34
_COL_ETAT_PT2 = 43
_COL_DECLASS_PT2 = 44
_COL_ETAT_PT3 = 53
_COL_DECLASS_PT3 = 54
_COL_DIST_ELEC = 67

_PLACEHOLDERS_NUM_VOIE = {'9999', '99999'}
_CODES_DECLASSE = {'00', '01'}
_CODES_INACC_ENV = {'IN8', 'IN9'}


# ===========================================================================
#  DATACLASSES
# ===========================================================================

@dataclass
class GespotRecord:
    """Un appui GESPOT après normalisation."""
    num: str
    voie: str
    num_voie: str
    centre: str
    dist_elec: str
    source_file: str

    type_calc: str
    strategie_calc: str
    milieu_calc: str
    pres_elect_calc: str
    inacc_calc: str
    recalage_raw: str
    etat_no_yellow: str
    etat_usable: str
    ctrl_visuel: str

    num_voie_needs_check: bool = False


@dataclass
class GespotLoadResult:
    """Résultat de chargement d'un répertoire GESPOT."""
    records: Dict[str, GespotRecord] = field(default_factory=dict)
    anomalies: List[dict] = field(default_factory=list)
    file_counts: Dict[str, int] = field(default_factory=dict)


# ===========================================================================
#  RÈGLES MÉTIER
# ===========================================================================

def _br01_calc_type(type_raw: str, caracs: List[str]) -> str:
    """BR-01 : Calcule le type d'appui final depuis type + CARAC1-5."""
    base = type_raw.strip()
    if not base:
        return ''
    has_anc = any(c.strip().upper() == 'ANC' for c in caracs)
    is_mc = base.upper().startswith('MC')
    if is_mc and has_anc:
        return f'{base} ANC MIN'
    if is_mc:
        return f'{base} MIN'
    if has_anc:
        return f'{base} ANC'
    return base


def _br02_calc_strategie(strategique: str, support_pc: str,
                         caracs: List[str], whitelist_upper: Set[str]) -> str:
    """BR-02 : Calcule la stratégie depuis strategique + support_pc + CARAC1-5."""
    if strategique.strip().upper() != 'O':
        return 'Non'
    sp = support_pc.strip()
    if sp:
        return sp
    for carac in caracs:
        v = carac.strip()
        if v and v.upper() != 'ANC' and v.upper() in whitelist_upper:
            return v
    return 'Non'


def _br03_calc_env(envs: List[str]) -> Tuple[str, bool]:
    """BR-03 : Extrait milieu environnant et détecte inaccessibilité ENV."""
    has_inacc = False
    milieu = ''
    for v in envs:
        code = v.strip().upper()
        if code in _CODES_INACC_ENV:
            has_inacc = True
        elif code and not milieu:
            milieu = v.strip()
    return milieu, has_inacc


def _br04_calc_etat(etat_pts: List[Tuple[str, str]]) -> Tuple[str, str, str]:
    """BR-04 : Détermine etat_no_yellow, etat_usable, ctrl_visuel."""
    for etat, declass in etat_pts:
        if etat.strip() in _CODES_DECLASSE:
            ctrl = 'Non' if declass.strip() else 'Oui'
            return 'Non', 'Non', ctrl
    return 'Oui', 'Oui', 'Oui'


def _br04bis_calc_pres_elect(pres_elect_raw: str) -> str:
    """BR-04 bis : Normalise pres_elect (vide → 'Non')."""
    v = pres_elect_raw.strip()
    return v if v else 'Non'


def _br05_calc_inacc(inacc_raw: str, has_inacc_env: bool) -> str:
    """BR-05 : Calcule l'inaccessibilité véhicule."""
    if inacc_raw.strip().upper() == 'O' or has_inacc_env:
        return 'Oui'
    return 'Non'


def _br06_calc_recalage(recalage_raw: str) -> str:
    """BR-06 : Inversion RECALAGE → verticalité attendue C6."""
    if recalage_raw.strip().upper() == 'O':
        return 'Non'
    return 'Oui'


# ===========================================================================
#  PARSING CSV
# ===========================================================================

def _row_to_cols(row: List[str], indices: List[int]) -> List[str]:
    return [row[i] if i < len(row) else '' for i in indices]


def _parse_one_csv(filepath: str) -> Tuple[List[List[str]], str]:
    """Lit un fichier CSV GESPOT, retourne (rows, encoding_used).

    Raises:
        ValueError: si décodage impossible ou header invalide.
    """
    for enc in ('utf-8-sig', 'latin-1'):
        try:
            with open(filepath, encoding=enc, newline='') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader, None)
                if not _validate_header(header):
                    raise ValueError(
                        f"Header CSV invalide (colonnes attendues: {_GESPOT_HEADER_MARKERS})"
                    )
                rows = list(reader)
            return rows, enc
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Impossible de decoder {filepath}")


def _validate_header(header: Optional[List[str]]) -> bool:
    """Vérifie que le header contient les colonnes marqueurs GESPOT."""
    if not header:
        return False
    upper_cols = {c.strip().upper() for c in header}
    return _GESPOT_HEADER_MARKERS.issubset(upper_cols)


def _build_record(row: List[str], source: str,
                  whitelist_upper: set) -> GespotRecord:
    """Construit un GespotRecord depuis une ligne CSV brute."""
    def col(i):
        return row[i].strip() if i < len(row) else ''

    caracs = [col(i) for i in _COL_CARAC]
    envs = [col(i) for i in _COL_ENV]
    etat_pts = [
        (col(_COL_ETAT_PT1), col(_COL_DECLASS_PT1)),
        (col(_COL_ETAT_PT2), col(_COL_DECLASS_PT2)),
        (col(_COL_ETAT_PT3), col(_COL_DECLASS_PT3)),
    ]

    milieu, has_inacc_env = _br03_calc_env(envs)
    etat_no_yellow, etat_usable, ctrl_visuel = _br04_calc_etat(etat_pts)
    num_voie_raw = col(_COL_NUM_VOIE)

    return GespotRecord(
        num=col(_COL_NUM),
        voie=col(_COL_VOIE),
        num_voie=num_voie_raw,
        centre=col(_COL_CENTRE),
        dist_elec=col(_COL_DIST_ELEC),
        source_file=source,
        type_calc=_br01_calc_type(col(_COL_TYPE), caracs),
        strategie_calc=_br02_calc_strategie(
            col(_COL_STRATEGIQUE), col(_COL_SUPPORT_PC), caracs, whitelist_upper),
        milieu_calc=milieu,
        pres_elect_calc=_br04bis_calc_pres_elect(col(_COL_PRES_ELECT)),
        inacc_calc=_br05_calc_inacc(col(_COL_INACC), has_inacc_env),
        recalage_raw=col(_COL_RECALAGE),
        etat_no_yellow=etat_no_yellow,
        etat_usable=etat_usable,
        ctrl_visuel=ctrl_visuel,
        num_voie_needs_check=(
            bool(num_voie_raw) and num_voie_raw not in _PLACEHOLDERS_NUM_VOIE
        ),
    )


# ===========================================================================
#  CHARGEMENT DU RÉPERTOIRE
# ===========================================================================

def load_gespot_dir(gespot_dir: str,
                   whitelist: Optional[set] = None) -> GespotLoadResult:
    """Charge tous les CSV d'un répertoire GESPOT.

    Args:
        gespot_dir: Chemin absolu du dossier GESPOT/GSPOT.
        whitelist:  Ensemble de codes stratégie valides (depuis Bases col M).

    Returns:
        GespotLoadResult avec records indexés par NUM, anomalies, compteurs.
    """
    if whitelist is None:
        whitelist = set()

    whitelist_upper = {v.upper() for v in whitelist}

    result = GespotLoadResult()
    try:
        entries = os.listdir(gespot_dir)
    except OSError as e:
        result.anomalies.append({
            'source': 'GESPOT', 'fichier': '', 'num': '',
            'type': 'FICHIER_IGNORE',
            'detail': f"Impossible de lire le dossier: {e}",
            'action': 'Verifier que le dossier GESPOT existe et est accessible',
        })
        return result

    csv_files = sorted(f for f in entries if f.lower().endswith('.csv'))
    if not csv_files:
        result.anomalies.append({
            'source': 'GESPOT', 'fichier': '', 'num': '',
            'type': 'FICHIER_IGNORE',
            'detail': 'Aucun fichier CSV trouve dans le dossier GESPOT',
            'action': 'Placer les fichiers GESPOT (CSV ; delimiter) dans le dossier',
        })
        return result

    raw_by_num: Dict[str, List[dict]] = {}

    for fname in csv_files:
        path = os.path.join(gespot_dir, fname)
        try:
            rows, _ = _parse_one_csv(path)
        except Exception as e:
            result.anomalies.append({
                'source': 'GESPOT', 'fichier': fname, 'num': '',
                'type': 'FICHIER_IGNORE', 'detail': str(e),
                'action': 'Verifier encodage et format CSV',
            })
            continue

        valid = 0
        for row in rows:
            num_raw = row[_COL_NUM].strip() if _COL_NUM < len(row) else ''
            if not num_raw:
                continue
            num_key = normalize_appui_num(num_raw)
            if not num_key:
                continue
            valid += 1
            raw_by_num.setdefault(num_key, []).append({'row': row, 'file': fname})

        result.file_counts[fname] = valid

    _resolve_duplicates(raw_by_num, result, whitelist_upper)
    return result


def _resolve_duplicates(raw_by_num: Dict[str, List[dict]],
                        result: GespotLoadResult,
                        whitelist_upper: set) -> None:
    """Résout les doublons NUM et alimente records + anomalies."""
    for num, entries in raw_by_num.items():
        if len(entries) == 1:
            rec = _build_record(entries[0]['row'], entries[0]['file'], whitelist_upper)
            result.records[num] = rec
            if rec.num_voie_needs_check:
                result.anomalies.append({
                    'source': 'GESPOT', 'fichier': entries[0]['file'], 'num': num,
                    'type': 'ADRESSE_A_VALIDER',
                    'detail': f"num_voie={rec.num_voie}",
                    'action': 'Verifier si num_voie doit etre concatene a VOIE',
                })
            continue

        payloads = [str(e['row']) for e in entries]
        if len(set(payloads)) == 1:
            rec = _build_record(entries[0]['row'], entries[0]['file'], whitelist_upper)
            result.records[num] = rec
            result.anomalies.append({
                'source': 'GESPOT', 'fichier': entries[0]['file'], 'num': num,
                'type': 'DOUBLON_IDENTIQUE',
                'detail': f"Apparu {len(entries)} fois, consolide",
                'action': 'Aucune (consolide automatiquement)',
            })
        else:
            files = ', '.join(sorted({e['file'] for e in entries}))
            result.anomalies.append({
                'source': 'GESPOT', 'fichier': files, 'num': num,
                'type': 'DOUBLON_CONFLICTUEL',
                'detail': f"Valeurs differentes sur {len(entries)} lignes",
                'action': 'Corriger la source GESPOT avant de relancer',
            })
