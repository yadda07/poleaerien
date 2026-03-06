# -*- coding: utf-8 -*-
"""
Preflight Data Quality Checks — Vérifications qualité données.

Module autonome appelable depuis:
- batch_orchestrator._preflight_data_quality() (auto au démarrage)
- dialog_v2._run_diagnostic() (bouton Diagnostic)

Chaque check retourne un dict:
  {'id': 'PF-01', 'level': 'BLOCKER'|'WARN'|'OK', 'message': str}
"""

import os
import csv
from typing import List, Dict, Optional

from qgis.core import (
    QgsVectorLayer, NULL
)


def run_data_quality_checks(
    lyr_pot: Optional[QgsVectorLayer],
    lyr_cap: Optional[QgsVectorLayer],
    lyr_com: Optional[QgsVectorLayer],
    module_keys: List[str],
    auto_field_fn=None,
    be_type: str = 'nge',
    gracethd_dir: str = '',
    comac_dir: str = '',
    capft_dir: str = '',
) -> List[Dict]:
    """Execute all preflight data quality checks.
    
    Args:
        lyr_pot: infra_pt_pot layer (or None)
        lyr_cap: etude_cap_ft layer (or None)
        lyr_com: etude_comac layer (or None)
        module_keys: list of module keys selected (['comac', 'capft', ...])
        auto_field_fn: callable(layer) -> field_name for study name field detection
        be_type: 'nge' or 'axione'
        gracethd_dir: path to GraceTHD directory (Axione only)
        comac_dir: path to COMAC Excel directory
        capft_dir: path to CAP_FT directory
    
    Returns:
        List of check results, each: {'id', 'level', 'message'}
        level is 'BLOCKER', 'WARN', or 'OK'
    """
    results = []
    mods = set(module_keys)

    # === PF-01: Doublons noms d'etudes ===
    for lyr, label, concerned in [
        (lyr_cap, 'etude_cap_ft', {'capft', 'c6bd', 'police_c6'}),
        (lyr_com, 'etude_comac', {'comac'}),
    ]:
        if not lyr or not lyr.isValid() or not (mods & concerned):
            continue
        field_name = auto_field_fn(lyr) if auto_field_fn else None
        if not field_name:
            results.append({
                'id': 'PF-01', 'level': 'WARN',
                'message': f'{label}: champ etude non detecte — verification doublons ignoree. '
                           f'Champs disponibles: {[f.name() for f in lyr.fields()][:8]}'
            })
            continue
        names = {}
        for feat in lyr.getFeatures():
            val = feat[field_name]
            if val and val != NULL:
                name = str(val).strip()
                names[name] = names.get(name, 0) + 1
        doublons = {n: c for n, c in names.items() if c > 1}
        if doublons:
            for name, count in doublons.items():
                results.append({
                    'id': 'PF-01', 'level': 'WARN',
                    'message': f'{label}: nom d\'etude "{name}" present sur {count} polygones '
                               f'(normal si etude multi-communes).'
                })
        else:
            results.append({
                'id': 'PF-01', 'level': 'OK',
                'message': f'{label}: aucun doublon d\'etude'
            })

    # === PF-01b + PF-03 + PF-04: infra_pt_pot qualite (single pass) ===
    if lyr_pot and lyr_pot.isValid():
        field_names = {f.name() for f in lyr_pot.fields()}
        has_inf_num = 'inf_num' in field_names
        inf_nums = {}
        nb_null = 0
        nb_no_geom = 0
        total_pot = lyr_pot.featureCount()

        for feat in lyr_pot.getFeatures():
            if not feat.hasGeometry():
                nb_no_geom += 1
            if has_inf_num:
                val = feat['inf_num']
                if val and val != NULL and str(val).strip():
                    key = str(val).strip()
                    inf_nums[key] = inf_nums.get(key, 0) + 1
                else:
                    nb_null += 1

        # PF-01b
        dup_inf = {n: c for n, c in inf_nums.items() if c > 1}
        if dup_inf:
            sample = list(dup_inf.keys())[:5]
            results.append({
                'id': 'PF-01b', 'level': 'WARN',
                'message': f'infra_pt_pot: {len(dup_inf)} inf_num en doublon '
                           f'(ex: {", ".join(sample)}). Resultats potentiellement faux.'
            })
        else:
            results.append({
                'id': 'PF-01b', 'level': 'OK',
                'message': f'infra_pt_pot: aucun inf_num en doublon ({len(inf_nums)} uniques)'
            })

        # PF-03
        if nb_null > 0:
            results.append({
                'id': 'PF-03', 'level': 'WARN',
                'message': f'infra_pt_pot: {nb_null}/{total_pot} poteaux sans inf_num '
                           f'(seront ignores dans la comparaison).'
            })
        else:
            results.append({
                'id': 'PF-03', 'level': 'OK',
                'message': f'infra_pt_pot: tous les poteaux ont un inf_num ({total_pot} total)'
            })

        # PF-04
        if nb_no_geom > 0:
            results.append({
                'id': 'PF-04', 'level': 'WARN',
                'message': f'infra_pt_pot: {nb_no_geom}/{total_pot} poteaux sans geometrie '
                           f'(matching spatial cables impossible pour ceux-ci).'
            })
        else:
            results.append({
                'id': 'PF-04', 'level': 'OK',
                'message': f'infra_pt_pot: tous les poteaux ont une geometrie ({total_pot} total)'
            })

    # === PF-02: Polygones sans geometrie ===
    for lyr, label, concerned in [
        (lyr_cap, 'etude_cap_ft', {'capft'}),
        (lyr_com, 'etude_comac', {'comac'}),
    ]:
        if not lyr or not lyr.isValid() or not (mods & concerned):
            continue
        total = lyr.featureCount()
        no_geom = sum(1 for f in lyr.getFeatures() if not f.hasGeometry())
        if no_geom > 0:
            results.append({
                'id': 'PF-02', 'level': 'WARN',
                'message': f'{label}: {no_geom}/{total} polygones sans geometrie '
                           f'(poteaux dans ces zones non detectes).'
            })
        else:
            results.append({
                'id': 'PF-02', 'level': 'OK',
                'message': f'{label}: {total} polygones valides'
            })

    # === PF-05: CRS coherent ===
    crs_map = {}
    for lyr, label in [(lyr_pot, 'infra_pt_pot'), (lyr_cap, 'etude_cap_ft'), (lyr_com, 'etude_comac')]:
        if lyr and lyr.isValid() and lyr.crs().isValid():
            crs_map[label] = lyr.crs().authid()
    unique_crs = set(crs_map.values())
    if len(unique_crs) > 1:
        details = ', '.join(f'{k}={v}' for k, v in crs_map.items())
        results.append({
            'id': 'PF-05', 'level': 'BLOCKER',
            'message': f'CRS incoherent entre couches: {details}. '
                       f'Toutes les couches doivent avoir le meme CRS (EPSG:2154 attendu).'
        })
    elif len(unique_crs) == 1 and len(crs_map) >= 2:
        results.append({
            'id': 'PF-05', 'level': 'OK',
            'message': f'CRS coherent: {unique_crs.pop()} ({len(crs_map)} couches)'
        })

    # === PF-06: Dossier COMAC ===
    if 'comac' in mods and comac_dir and os.path.isdir(comac_dir):
        comac_xlsx = 0
        for _root, _, _files in os.walk(comac_dir):
            comac_xlsx += sum(1 for f in _files if f.endswith('.xlsx') and '~$' not in f)
        if comac_xlsx == 0:
            results.append({
                'id': 'PF-06', 'level': 'WARN',
                'message': f'Dossier COMAC: 0 fichier .xlsx trouve dans {os.path.basename(comac_dir)}.'
            })
        else:
            results.append({
                'id': 'PF-06', 'level': 'OK',
                'message': f'Dossier COMAC: {comac_xlsx} fichier(s) .xlsx'
            })

    # === PF-07: Dossier CAP_FT ===
    if 'capft' in mods and capft_dir and os.path.isdir(capft_dir):
        fiche_count = 0
        for _root, _, _files in os.walk(capft_dir):
            fiche_count += sum(1 for f in _files if 'FicheAppui' in f and f.endswith('.xlsx'))
        if fiche_count == 0:
            results.append({
                'id': 'PF-07', 'level': 'WARN',
                'message': 'Dossier CAP_FT: aucun fichier FicheAppui_*.xlsx detecte.'
            })
        else:
            results.append({
                'id': 'PF-07', 'level': 'OK',
                'message': f'Dossier CAP_FT: {fiche_count} FicheAppui detecte(s)'
            })

    # === PF-09 to PF-10e: GraceTHD MCD (Axione only) ===
    if be_type == 'axione' and gracethd_dir and os.path.isdir(gracethd_dir):
        results.extend(_check_gracethd_mcd(gracethd_dir))

    return results


def _check_gracethd_mcd(gdir: str) -> List[Dict]:
    """GraceTHD MCD integrity checks (PF-09 to PF-10e)."""
    results = []

    # PF-09: Required/optional files
    required = {'t_noeud.shp', 't_cableline.shp'}
    optional = {'t_cable.csv', 't_ptech.csv', 't_ebp.csv'}
    missing_req = [f for f in required if not os.path.isfile(os.path.join(gdir, f))]
    missing_opt = [f for f in optional if not os.path.isfile(os.path.join(gdir, f))]

    if missing_req:
        results.append({
            'id': 'PF-09', 'level': 'BLOCKER',
            'message': f'GraceTHD: fichiers requis manquants: {", ".join(missing_req)}. '
                       f'Placez-les dans {os.path.basename(gdir)}/.'
        })
        return results  # Can't check joins without required files
    else:
        results.append({'id': 'PF-09', 'level': 'OK', 'message': 'GraceTHD: fichiers requis presents'})

    if missing_opt:
        results.append({
            'id': 'PF-09', 'level': 'WARN',
            'message': f'GraceTHD: fichiers optionnels absents: {", ".join(missing_opt)}'
        })

    # PF-10 + PF-10b: Cables DI + join integrity
    t_cable_path = os.path.join(gdir, 't_cable.csv')
    t_cableline_path = os.path.join(gdir, 't_cableline.shp')
    if os.path.isfile(t_cable_path) and os.path.isfile(t_cableline_path):
        try:
            with open(t_cable_path, 'r', encoding='utf-8', errors='replace') as f:
                cables = list(csv.DictReader(f, delimiter=';'))
            cables_di = [r for r in cables if r.get('cb_typelog', '').upper() == 'DI']
            nb_di = len(cables_di)

            if nb_di == 0:
                results.append({
                    'id': 'PF-10', 'level': 'BLOCKER',
                    'message': f'GraceTHD: 0 cables DI dans t_cable.csv ({len(cables)} cables total). '
                               f'Verification cables impossible.'
                })
            else:
                results.append({
                    'id': 'PF-10', 'level': 'OK',
                    'message': f'GraceTHD t_cable.csv: {nb_di} cables DI / {len(cables)} total'
                })

                # PF-10b: Join cables -> cableline
                cl_layer = QgsVectorLayer(t_cableline_path, 'cl_check', 'ogr')
                if cl_layer.isValid():
                    cl_codes = {str(f['cl_cb_code']).strip() for f in cl_layer.getFeatures() if f['cl_cb_code']}
                    del cl_layer
                    cb_codes_di = {r.get('cb_code', '').strip() for r in cables_di}
                    missing_join = cb_codes_di - cl_codes
                    if missing_join:
                        results.append({
                            'id': 'PF-10b', 'level': 'WARN',
                            'message': f'GraceTHD jointure: {len(missing_join)}/{nb_di} cables DI '
                                       f'sans geometrie (cb_code absent de t_cableline.shp).'
                        })
                    else:
                        results.append({
                            'id': 'PF-10b', 'level': 'OK',
                            'message': f'GraceTHD jointure: {nb_di}/{nb_di} cables DI ont une geometrie'
                        })

                # PF-10e: Capacites
                valid_capas = {6, 12, 24, 36, 48, 72, 96, 144, 288, 576}
                invalid = sum(1 for r in cables_di if _invalid_capa(r, valid_capas))
                if invalid:
                    results.append({
                        'id': 'PF-10e', 'level': 'WARN',
                        'message': f'GraceTHD: {invalid} cables DI avec capacite non standard.'
                    })

        except Exception as e:
            results.append({
                'id': 'PF-10', 'level': 'WARN',
                'message': f'GraceTHD: erreur lecture t_cable.csv: {e}'
            })

    # PF-10c: Join BPE (t_ebp -> t_ptech -> t_noeud)
    t_ebp = os.path.join(gdir, 't_ebp.csv')
    t_ptech = os.path.join(gdir, 't_ptech.csv')
    t_noeud = os.path.join(gdir, 't_noeud.shp')
    if all(os.path.isfile(p) for p in [t_ebp, t_ptech, t_noeud]):
        try:
            with open(t_ebp, 'r', encoding='utf-8', errors='replace') as f:
                ebp_rows = list(csv.DictReader(f, delimiter=';'))
            with open(t_ptech, 'r', encoding='utf-8', errors='replace') as f:
                ptech_rows = list(csv.DictReader(f, delimiter=';'))

            pt_codes = {r.get('pt_code', '').strip() for r in ptech_rows}
            pt_nd_map = {r.get('pt_code', '').strip(): r.get('pt_nd_code', '').strip() for r in ptech_rows}

            nd_layer = QgsVectorLayer(t_noeud, 'nd_check', 'ogr')
            nd_codes = {str(f['nd_code']).strip() for f in nd_layer.getFeatures() if f['nd_code']} if nd_layer.isValid() else set()
            del nd_layer

            nb_broken = 0
            for ebp in ebp_rows:
                bp_pt = ebp.get('bp_pt_code', '').strip()
                if bp_pt not in pt_codes:
                    nb_broken += 1
                elif pt_nd_map.get(bp_pt, '') not in nd_codes:
                    nb_broken += 1

            if nb_broken > 0:
                results.append({
                    'id': 'PF-10c', 'level': 'WARN',
                    'message': f'GraceTHD jointure BPE: {nb_broken}/{len(ebp_rows)} BPE '
                               f'sans geometrie (jointure t_ebp->t_ptech->t_noeud cassee).'
                })
            else:
                results.append({
                    'id': 'PF-10c', 'level': 'OK',
                    'message': f'GraceTHD jointure BPE: {len(ebp_rows)} BPE valides'
                })
        except Exception as e:
            results.append({
                'id': 'PF-10c', 'level': 'WARN',
                'message': f'GraceTHD: erreur verification BPE: {e}'
            })

    return results


def _invalid_capa(row: dict, valid: set) -> bool:
    try:
        c = int(row.get('cb_capafo', '0') or '0')
        return c not in valid and c != 0
    except (ValueError, TypeError):
        return True


def format_check_log(check: Dict) -> tuple:
    """Format a check result for log display.
    
    Returns:
        (message_str, log_level_str) where log_level_str is 'success'|'warning'|'error'
    """
    level = check['level']
    msg = check['message']
    if level == 'OK':
        return f"  [OK] {msg}", 'success'
    elif level == 'WARN':
        return f"  [WARN] {msg}", 'warning'
    else:
        return f"  [ERREUR] {msg}", 'error'
