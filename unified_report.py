# -*- coding: utf-8 -*-
"""
Unified Excel Report - Professional diagnostic workbook.

Generates a single standardized Excel file with:
  1. TABLEAU DE BORD: executive summary with KPIs per module
  2. Module sheets: consistent formatting, borders, filters, freeze panes

Style: Calibri 10pt, dark blue headers (#1F4E79), thin grey borders,
       status fills (green=OK, amber=warning, red=error).
"""

import os
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from qgis.core import QgsMessageLog, Qgis


# ======================================================================
#  UNIFIED STYLE SYSTEM
# ======================================================================

_TAB = {
    'maj': 'F59E0B', 'capft': '34D399', 'comac': 'EA580C',
    'c6bd': 'EC4899', 'police_c6': 'A855F7', 'c6c3a': '3B82F6',
}
_NAMES = {
    'maj': '0. MAJ BD', 'capft': '1. VERIF CAP_FT', 'comac': '2. VERIF COMAC',
    'c6bd': '3. C6 vs BD', 'police_c6': '4. POLICE C6', 'c6c3a': '5. C6-C3A-BD',
}

_F_TITLE = Font(name='Calibri', size=14, bold=True, color='1F4E79')
_F_HEAD = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
_F_DATA = Font(name='Calibri', size=10)
_F_BOLD = Font(name='Calibri', size=10, bold=True)
_F_PCT = Font(name='Calibri', size=11, bold=True)

_P_HEAD = PatternFill('solid', fgColor='1F4E79')
_P_OK = PatternFill('solid', fgColor='C6EFCE')
_P_WARN = PatternFill('solid', fgColor='FFEB9C')
_P_ERR = PatternFill('solid', fgColor='FFC7CE')
_P_CRIT = PatternFill('solid', fgColor='FF6B6B')
_P_INFO = PatternFill('solid', fgColor='D6E4F0')

_BRD = Border(
    left=Side('thin', color='D9D9D9'), right=Side('thin', color='D9D9D9'),
    top=Side('thin', color='D9D9D9'), bottom=Side('thin', color='D9D9D9'),
)
_AL_C = Alignment(horizontal='center', vertical='center')
_AL_L = Alignment(horizontal='left', vertical='center')
_AL_W = Alignment(horizontal='left', vertical='center', wrap_text=True)
_AL_CV = Alignment(horizontal='center', vertical='center', wrap_text=True)

_F_MOD = Font(name='Calibri', size=10, bold=True, color='1F4E79')


# ======================================================================
#  HELPERS
# ======================================================================

def _init_sheet(ws, key, headers, widths):
    """Apply standard header, tab color, freeze panes, auto-filter."""
    ws.sheet_properties.tabColor = _TAB.get(key, '808080')
    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = _F_HEAD
        cell.fill = _P_HEAD
        cell.alignment = _AL_C
        cell.border = _BRD
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}1'


def _row(ws, r, vals, font=None, fill=None, align=None):
    """Write one data row with consistent styling."""
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = font or _F_DATA
        cell.alignment = align or _AL_L
        cell.border = _BRD
        if fill:
            cell.fill = fill


def _fill_row(ws, r, fill, ncols):
    for c in range(1, ncols + 1):
        ws.cell(row=r, column=c).fill = fill
        ws.cell(row=r, column=c).border = _BRD


def _df_sheet(wb, name, df, key):
    """Write DataFrame to a new sheet with standard formatting."""
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = _TAB.get(key, '808080')
    ncols = len(df.columns)
    for c, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=c, value=str(col_name))
        cell.font = _F_HEAD
        cell.fill = _P_HEAD
        cell.alignment = _AL_C
        cell.border = _BRD
        ws.column_dimensions[get_column_letter(c)].width = max(14, len(str(col_name)) + 4)
    for ri, row_data in enumerate(df.itertuples(index=False), 2):
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font = _F_DATA
            cell.border = _BRD
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(ncols)}1'
    return ws


# ======================================================================
#  KPI EXTRACTION (for dashboard)
# ======================================================================

def _kpi_maj(r):
    ft = r.get('liste_ft', [0, None, 0])
    bt = r.get('liste_bt', [0, None, 0])
    ok = (ft[2] if len(ft) > 2 else 0) + (bt[2] if len(bt) > 2 else 0)
    nok = (ft[0] if ft else 0) + (bt[0] if bt else 0)
    return ok, nok, f"FT:{ft[2] if len(ft)>2 else 0} id/{ft[0] if ft else 0} introuv | BT:{bt[2] if len(bt)>2 else 0} id/{bt[0] if bt else 0} introuv"

def _kpi_capft(r):
    res = r.get('resultats')
    if not res: return 0, 0, ''
    ne = sum(len(v) for v in res[0].values())
    nq = sum(len(v) for v in res[1].values())
    ok = len(res[2])
    return ok, ne + nq, f"{ok} OK | {ne} absents QGIS | {nq} absents Excel"

def _kpi_comac(r):
    base_ok, base_nok, base_msg = _kpi_capft(r)
    # Ajouter câbles + boîtiers au bilan
    verif_cables = r.get('verif_cables')
    cable_nok = 0
    if verif_cables:
        cable_nok = sum(1 for e in verif_cables if e.get('statut') in ('ECART', 'ABSENT_BDD'))
    verif_boitiers = r.get('verif_boitiers')
    boitier_nok = 0
    if verif_boitiers:
        boitier_nok = sum(1 for v in verif_boitiers.values() if v.get('statut') == 'ERREUR')
    total_nok = base_nok + cable_nok + boitier_nok
    parts = [base_msg] if base_msg else []
    if cable_nok:
        parts.append(f"{cable_nok} ecarts cables")
    if boitier_nok:
        parts.append(f"{boitier_nok} err boitiers")
    return base_ok, total_nok, ' | '.join(parts)

def _kpi_c6bd(r):
    df = r.get('final_df')
    if df is None or df.empty: return 0, 0, ''
    nok = int(df['Statut'].astype(str).str.contains('ABSENT').sum()) if 'Statut' in df.columns else 0
    ok = len(df) - nok
    out = len(r.get('df_poteaux_out', pd.DataFrame()))
    return ok, nok, f"{ok} trouves | {nok} absents | {out} hors perimetre"

def _kpi_police(r):
    st = r.get('stats', [])
    ok = sum(s.get('nb_ok', 0) for s in st)
    nok = sum(s.get('nb_ecart', 0) + s.get('nb_absent', 0) + s.get('nb_boitier_err', 0) for s in st)
    return ok, nok, f"{ok} OK | {nok} anomalies sur {len(st)} etudes"

def _kpi_c6c3a(r):
    df = r.get('df_final')
    if df is None or df.empty: return 0, 0, ''
    nok = sum(int((df[c] == 'ABSENT').sum()) for c in df.columns if ('inf_num' in c or 'Excel' in c))
    return max(0, len(df) - nok), nok, f"{len(df)} lignes | {nok} ABSENT"

_KPI = {'maj': _kpi_maj, 'capft': _kpi_capft, 'comac': _kpi_comac,
        'c6bd': _kpi_c6bd, 'police_c6': _kpi_police, 'c6c3a': _kpi_c6c3a}


# ======================================================================
#  SUB-CHECKS (detailed verifications per module for dashboard)
# ======================================================================

def _checks_maj(r):
    ft = r.get('liste_ft', [0, None, 0])
    bt = r.get('liste_bt', [0, None, 0])
    ok_ft = ft[2] if len(ft) > 2 else 0
    nok_ft = ft[0] if ft else 0
    ok_bt = bt[2] if len(bt) > 2 else 0
    nok_bt = bt[0] if bt else 0
    return [
        ("Correspondance FT (Excel/QGIS)", ok_ft, nok_ft,
         f"{ok_ft} identifies, {nok_ft} introuvables"),
        ("Correspondance BT (Excel/QGIS)", ok_bt, nok_bt,
         f"{ok_bt} identifies, {nok_bt} introuvables"),
    ]


def _checks_capft(r):
    res = r.get('resultats')
    if not res:
        return [("Analyse appuis", 0, 0, "Pas de resultats")]
    ne = sum(len(v) for v in res[0].values())
    nq = sum(len(v) for v in res[1].values())
    ok = len(res[2])
    return [
        ("Appuis presents dans QGIS", ok, ne,
         f"{ok} trouves, {ne} absents QGIS"),
        ("Appuis presents dans fiches", ok, nq,
         f"{ok} trouves, {nq} absents fiches appuis"),
    ]


def _checks_comac(r):
    res = r.get('resultats')
    checks = []
    if res:
        ne = sum(len(v) for v in res[0].values())
        nq = sum(len(v) for v in res[1].values())
        ok = len(res[2])
        checks.append(("Appuis presents dans QGIS", ok, ne,
                        f"{ok} trouves, {ne} absents QGIS"))
        checks.append(("Appuis presents dans Excel", ok, nq,
                        f"{ok} trouves, {nq} absents Excel"))
    dico_secu = r.get('dico_verif_secu')
    if dico_secu:
        ok_p, nok_p, ok_h, nok_h = 0, 0, 0, 0
        for liste_v in dico_secu.values():
            for v in liste_v:
                portee = v.get('portee', 0)
                hauteur = v.get('hauteur_sol', 0)
                if portee > 0:
                    vp = v.get('verif_portee')
                    if vp:
                        if vp.get('valide', True):
                            ok_p += 1
                        else:
                            nok_p += 1
                if hauteur > 0:
                    vh = v.get('verif_hauteur_sol')
                    if vh:
                        if vh.get('valide', True):
                            ok_h += 1
                        else:
                            nok_h += 1
        if ok_p + nok_p > 0:
            checks.append(("Verif portees NFC 11201", ok_p, nok_p,
                            f"{ok_p} OK, {nok_p} depassements"))
        if ok_h + nok_h > 0:
            checks.append(("Verif hauteurs sol", ok_h, nok_h,
                            f"{ok_h} OK, {nok_h} insuffisants"))
    verif_cables = r.get('verif_cables')
    if verif_cables:
        cables_only = [e for e in verif_cables if e.get('statut')]
        if cables_only:
            ok_c = sum(1 for e in cables_only if e.get('statut') == 'OK')
            nok_c = len(cables_only) - ok_c
            checks.append(("Verif cables COMAC vs BDD", ok_c, nok_c,
                            f"{ok_c} OK, {nok_c} ecarts"))
    dico_boitier = r.get('dico_boitier_comac', {})
    verif_boitiers = r.get('verif_boitiers')
    if dico_boitier:
        nb_oui = sum(1 for v in dico_boitier.values() if str(v).lower() == 'oui')
        nb_non = sum(1 for v in dico_boitier.values() if str(v).lower() == 'non')
        if verif_boitiers:
            ok_b = sum(1 for v in verif_boitiers.values() if v.get('statut') == 'OK')
            nok_b = sum(1 for v in verif_boitiers.values() if v.get('statut') == 'ERREUR')
            checks.append(("Verif boitiers COMAC vs BPE", ok_b, nok_b,
                            f"{ok_b} OK, {nok_b} absents BPE ({nb_non} Non)"))
        else:
            checks.append(("Verif boitiers COMAC vs BPE", nb_non, 0,
                            f"{nb_oui} Oui, {nb_non} Non, aucune verif BPE necessaire"))
    if not checks:
        checks = [("Analyse COMAC", 0, 0, "Pas de resultats")]
    return checks


def _checks_c6bd(r):
    checks = []
    df = r.get('final_df')
    if df is not None and not df.empty:
        nok = int(df['Statut'].astype(str).str.contains('ABSENT').sum()) \
            if 'Statut' in df.columns else 0
        ok = len(df) - nok
        checks.append(("Correspondance appuis C6/QGIS", ok, nok,
                        f"{ok} trouves, {nok} absents"))
    out_df = r.get('df_poteaux_out')
    n_out = len(out_df) if out_df is not None and not out_df.empty else 0
    checks.append(("Poteaux hors perimetre", 0, n_out,
                    f"{n_out} hors perimetre" if n_out > 0 else "Aucun"))
    verif = r.get('verif_etudes')
    if verif:
        sans_c6 = len(verif.get('etudes_sans_c6', []))
        c6_sans = len(verif.get('c6_sans_etude', []))
        nok_v = sans_c6 + c6_sans
        checks.append(("Coherence etudes/fichiers C6", 0, nok_v,
                        f"{sans_c6} etudes sans C6, {c6_sans} C6 sans etude"))
    if not checks:
        checks = [("Analyse C6 vs BD", 0, 0, "Pas de resultats")]
    return checks


def _checks_police(r):
    st = r.get('stats', [])
    if not st:
        return [("Analyse Police C6", 0, 0, "Pas de resultats")]
    ok_cables = sum(s.get('nb_ok', 0) for s in st)
    ecarts = sum(s.get('nb_ecart', 0) for s in st)
    absents = sum(s.get('nb_absent', 0) for s in st)
    boitier_err = sum(s.get('nb_boitier_err', 0) for s in st)
    total_appuis = sum(s.get('appuis_c6', 0) for s in st)
    return [
        ("Comparaison cables C6/BDD", ok_cables, ecarts,
         f"{ok_cables} OK, {ecarts} ecarts sur {len(st)} etudes"),
        ("Presence appuis dans BDD", total_appuis - absents, absents,
         f"{absents} absents BDD" if absents > 0 else "Tous presents"),
        ("Verification boitiers (BPE)", total_appuis - boitier_err, boitier_err,
         f"{boitier_err} erreurs boitier" if boitier_err > 0 else "Tous conformes"),
    ]


def _checks_c6c3a(r):
    df = r.get('df_final')
    if df is None or df.empty:
        return [("Analyse C6-C3A-BD", 0, 0, "Pas de resultats")]
    checks = []
    for col_check, label in [
        ("inf_num (ETUDES_QGIS)", "Presence dans etudes QGIS"),
        ("inf_num (C3A)", "Presence dans C3A"),
        ("Excel (C6)", "Presence dans C6 Excel"),
    ]:
        if col_check in df.columns:
            nok = int((df[col_check] == 'ABSENT').sum())
            ok = len(df) - nok
            checks.append((label, ok, nok, f"{ok} trouves, {nok} absents"))
    if not checks:
        nok = sum(int((df[c] == 'ABSENT').sum())
                  for c in df.columns if ('inf_num' in c or 'Excel' in c))
        ok = max(0, len(df) - nok)
        checks = [("Correspondance globale", ok, nok,
                    f"{len(df)} lignes, {nok} absents")]
    return checks


_CHECKS = {
    'maj': _checks_maj, 'capft': _checks_capft, 'comac': _checks_comac,
    'c6bd': _checks_c6bd, 'police_c6': _checks_police, 'c6c3a': _checks_c6c3a,
}


# ======================================================================
#  TABLEAU DE BORD
# ======================================================================

def _write_dashboard(wb, batch_results):
    ws = wb.active
    ws.title = "TABLEAU DE BORD"
    ws.sheet_properties.tabColor = '1F4E79'

    ws.merge_cells('A1:H1')
    t = ws.cell(row=1, column=1, value="RAPPORT DE DIAGNOSTIC - POLE AERIEN")
    t.font = _F_TITLE
    t.alignment = _AL_C

    ws.cell(row=2, column=1, value="Date :").font = _F_BOLD
    ws.cell(row=2, column=2, value=datetime.now().strftime("%d/%m/%Y %H:%M")).font = _F_DATA
    ws.cell(row=3, column=1, value="Modules :").font = _F_BOLD
    ws.cell(row=3, column=2, value=str(len(batch_results))).font = _F_DATA

    hdrs = ["MODULE", "VERIFICATION", "STATUT", "TOTAL",
            "CONFORMES", "NON CONFORMES", "% CONFORMITE", "DETAIL"]
    wds = [20, 32, 16, 10, 14, 16, 16, 50]
    for c, (h, w) in enumerate(zip(hdrs, wds), 1):
        cell = ws.cell(row=5, column=c, value=h)
        cell.font = _F_HEAD
        cell.fill = _P_HEAD
        cell.alignment = _AL_C
        cell.border = _BRD
        ws.column_dimensions[get_column_letter(c)].width = w

    row = 6
    t_ok, t_nok = 0, 0

    for key in ['maj', 'capft', 'comac', 'c6bd', 'police_c6', 'c6c3a']:
        res = batch_results.get(key)
        if res is None:
            continue

        checks = _CHECKS.get(key, lambda r: [("Analyse", 0, 0, "")])(res)
        if not checks:
            continue

        kpi_ok, kpi_nok, _ = _KPI.get(key, lambda r: (0, 0, ''))(res)
        t_ok += kpi_ok
        t_nok += kpi_nok

        start_row = row
        for i, (check_name, ok, nok, detail) in enumerate(checks):
            total = ok + nok
            pct = (ok / total * 100) if total > 0 else 100
            status = "CONFORME" if nok == 0 else "NON CONFORME"

            mod_name = _NAMES.get(key, key) if i == 0 else ""

            _row(ws, row, [mod_name, check_name, status, total, ok, nok,
                           f"{pct:.0f}%", detail])

            sf = _P_OK if nok == 0 else (_P_WARN if pct >= 80 else _P_ERR)
            ws.cell(row=row, column=3).fill = sf
            ws.cell(row=row, column=3).alignment = _AL_C

            pf = _P_OK if pct == 100 else (_P_WARN if pct >= 80 else _P_ERR)
            ws.cell(row=row, column=7).fill = pf
            ws.cell(row=row, column=7).alignment = _AL_C
            ws.cell(row=row, column=7).font = _F_PCT

            for c in [4, 5, 6]:
                ws.cell(row=row, column=c).alignment = _AL_C

            if i == 0:
                ws.cell(row=row, column=1).font = _F_MOD

            row += 1

        if len(checks) > 1:
            ws.merge_cells(start_row=start_row, start_column=1,
                           end_row=row - 1, end_column=1)
            ws.cell(row=start_row, column=1).alignment = _AL_CV
            ws.cell(row=start_row, column=1).font = _F_MOD

    gt = t_ok + t_nok
    gp = (t_ok / gt * 100) if gt > 0 else 100
    _row(ws, row, ["TOTAL", "", "", gt, t_ok, t_nok, f"{gp:.0f}%", ""],
         font=_F_BOLD)
    for c in range(1, 9):
        ws.cell(row=row, column=c).fill = _P_INFO
        ws.cell(row=row, column=c).border = Border(
            top=Side('medium', color='1F4E79'), bottom=Side('medium', color='1F4E79'),
            left=_BRD.left, right=_BRD.right)
    ws.cell(row=row, column=7).alignment = _AL_C
    ws.cell(row=row, column=7).font = _F_PCT

    lr = row + 2
    ws.cell(row=lr, column=1, value="LEGENDE :").font = _F_BOLD
    for i, (f, lbl) in enumerate([(_P_OK, "Conforme"), (_P_WARN, "Avertissement"),
                                   (_P_ERR, "Non conforme"), (_P_CRIT, "Critique")]):
        ws.cell(row=lr + 1 + i, column=1).fill = f
        ws.cell(row=lr + 1 + i, column=1).border = _BRD
        ws.cell(row=lr + 1 + i, column=2, value=lbl).font = _F_DATA

    ws.freeze_panes = 'A6'



# ======================================================================
#  MAJ BD
# ======================================================================

def write_maj(wb, result):
    """Write MAJ BD sheets: resume, FT, BT, introuvables."""
    ft = result.get('liste_ft', [0, pd.DataFrame(), 0, pd.DataFrame()])
    bt = result.get('liste_bt', [0, pd.DataFrame(), 0, pd.DataFrame()])
    nok_ft = ft[0] if ft else 0
    df_nok_ft = ft[1] if len(ft) > 1 else pd.DataFrame()
    ok_ft = ft[2] if len(ft) > 2 else 0
    df_ok_ft = ft[3] if len(ft) > 3 else pd.DataFrame()
    nok_bt = bt[0] if bt else 0
    df_nok_bt = bt[1] if len(bt) > 1 else pd.DataFrame()
    ok_bt = bt[2] if len(bt) > 2 else 0
    df_ok_bt = bt[3] if len(bt) > 3 else pd.DataFrame()
    key = 'maj'

    ws = wb.create_sheet("MAJ_RESUME")
    _init_sheet(ws, key, ["ELEMENT", "NOMBRE", "STATUT"], [40, 12, 18])
    data = [
        ("FT identifies", ok_ft, "OK" if ok_ft > 0 else ""),
        ("FT introuvables (Excel sans QGIS)", nok_ft,
         "ANOMALIE" if nok_ft > 0 else "OK"),
        ("BT identifies", ok_bt, "OK" if ok_bt > 0 else ""),
        ("BT introuvables (Excel sans QGIS)", nok_bt,
         "ANOMALIE" if nok_bt > 0 else "OK"),
    ]
    for r, (label, val, st) in enumerate(data, 2):
        _row(ws, r, [label, val, st])
        ws.cell(row=r, column=2).alignment = _AL_C
        ws.cell(row=r, column=3).alignment = _AL_C
        if st == "ANOMALIE":
            ws.cell(row=r, column=3).fill = _P_ERR
        elif st == "OK":
            ws.cell(row=r, column=3).fill = _P_OK

    if isinstance(df_ok_ft, pd.DataFrame) and not df_ok_ft.empty:
        _df_sheet(wb, "MAJ_FT", df_ok_ft.reset_index(), key)
    if isinstance(df_ok_bt, pd.DataFrame) and not df_ok_bt.empty:
        _df_sheet(wb, "MAJ_BT", df_ok_bt.reset_index(), key)
    if isinstance(df_nok_ft, pd.DataFrame) and not df_nok_ft.empty:
        _df_sheet(wb, "MAJ_FT_INTROUV", df_nok_ft, key)
    if isinstance(df_nok_bt, pd.DataFrame) and not df_nok_bt.empty:
        _df_sheet(wb, "MAJ_BT_INTROUV", df_nok_bt, key)


# ======================================================================
#  CAP_FT
# ======================================================================

def _write_analyse_sheet(wb, sheet_name, key, resultats, label_nok_qgis):
    """Shared writer for CAP_FT and COMAC analyse sheets."""
    introuvables_excel = resultats[0]
    introuvables_qgis = resultats[1]
    existants = resultats[2]

    ws = wb.create_sheet(sheet_name)
    hdrs = ["INF_NUM QGIS", "ETUDE QGIS", "INF_NUM EXCEL", "NOM FICHIER", "STATUT"]
    _init_sheet(ws, key, hdrs, [22, 30, 22, 45, 35])

    r = 2
    for fichier, appuis in introuvables_excel.items():
        for inf_num in appuis:
            _row(ws, r, ["", "", inf_num, fichier, "ABSENT QGIS"], fill=_P_ERR)
            r += 1
    for etude, appuis in introuvables_qgis.items():
        for inf_num in appuis:
            _row(ws, r, [inf_num, etude, "", "", label_nok_qgis], fill=_P_WARN)
            r += 1
    for _, values in existants.items():
        vals = list(values) + ["OK"]
        _row(ws, r, vals, fill=_P_OK)
        r += 1
    return ws


def write_capft(wb, result):
    """Write CAP_FT analysis sheet."""
    resultats = result.get('resultats')
    if not resultats:
        return
    _write_analyse_sheet(wb, "CAPFT_ANALYSE", 'capft', resultats,
                         "ABSENT FICHES APPUIS")


# ======================================================================
#  COMAC
# ======================================================================

def write_comac(wb, result):
    """Write COMAC analysis sheets."""
    resultats = result.get('resultats')
    if not resultats:
        return

    key = 'comac'
    _write_analyse_sheet(wb, "COMAC_ANALYSE", key, resultats, "ABSENT EXCELS")

    dico_verif_secu = result.get('dico_verif_secu')
    if dico_verif_secu:
        ws = wb.create_sheet("COMAC_SECURITE")
        hdrs = ["FICHIER", "POTEAU", "PORTEE (m)", "CAPACITE FO", "TYPE LIGNE",
                "PORTEE MAX (m)", "DEPASSEMENT (m)", "HAUTEUR SOL (m)",
                "VERIF PORTEE", "VERIF HAUTEUR"]
        _init_sheet(ws, key, hdrs, [40, 20, 14, 14, 15, 14, 16, 16, 28, 28])
        r = 2
        for fichier, liste_v in dico_verif_secu.items():
            for v in liste_v:
                portee = v.get('portee', 0)
                capacite = v.get('capacite_fo', 0)
                hauteur = v.get('hauteur_sol', 0)
                if portee == 0 and capacite == 0 and hauteur == 0:
                    continue
                vp = v.get('verif_portee')
                p_max = vp.get('portee_max', 0) if vp else 0
                dep = vp.get('depassement', 0) if vp else 0
                p_ok = vp.get('valide', True) if vp else True
                vh = v.get('verif_hauteur_sol')
                h_ok = vh.get('valide', True) if vh else True
                st_p = "OK" if p_ok else f"DEPASSEMENT (+{dep:.1f}m)"
                st_h = ""
                if hauteur > 0:
                    st_h = "OK" if h_ok else f"INSUFFISANT ({hauteur:.1f}m)"
                vals = [fichier, v.get('poteau', ''), portee, capacite,
                        v.get('type_ligne_fo', ''), p_max, dep,
                        hauteur if hauteur > 0 else '',
                        st_p if portee > 0 else '', st_h]
                _row(ws, r, vals)
                if portee > 0:
                    ws.cell(row=r, column=9).fill = _P_OK if p_ok else _P_CRIT
                if hauteur > 0:
                    ws.cell(row=r, column=10).fill = _P_OK if h_ok else _P_CRIT
                r += 1

    verif_cables = result.get('verif_cables')
    if verif_cables:
        ws = wb.create_sheet("COMAC_CABLES")
        has_boitier = any(e.get('boitier_comac') for e in verif_cables)
        hdrs = ["APPUI", "NB COMAC", "REFS COMAC", "CAPAS COMAC",
                "NB BDD", "CAPAS BDD", "STATUT", "MESSAGE"]
        widths = [20, 12, 40, 25, 12, 25, 14, 55]
        if has_boitier:
            hdrs.extend(["BOITIER COMAC", "BPE TYPE", "BOITIER STATUT"])
            widths.extend([15, 15, 15])
        _init_sheet(ws, key, hdrs, widths)
        r = 2
        for e in verif_cables:
            st = e.get('statut', '')
            vals = [
                e.get('num_appui', ''), e.get('nb_cables_comac', 0),
                e.get('cables_comac', ''),
                '+'.join(e.get('capas_comac', [])),
                e.get('nb_cables_bdd', 0),
                '+'.join(str(c) for c in e.get('capas_bdd', [])),
                st, e.get('message', '')]
            if has_boitier:
                vals.extend([
                    e.get('boitier_comac', ''),
                    e.get('bpe_noe_type', ''),
                    e.get('boitier_statut', '')])
            fill = _P_OK if st == 'OK' else (_P_CRIT if st == 'ECART' else (
                _P_WARN if st == 'ABSENT_BDD' else None))
            _row(ws, r, vals, fill=fill)
            if has_boitier:
                bst = e.get('boitier_statut', '')
                if bst == 'ERREUR':
                    for c in range(9, 12):
                        ws.cell(row=r, column=c).fill = _P_CRIT
                elif bst == 'OK':
                    for c in range(9, 12):
                        ws.cell(row=r, column=c).fill = _P_OK
            r += 1


# ======================================================================
#  C6 vs BD
# ======================================================================

def write_c6bd(wb, result):
    """Write C6 vs BD analysis sheets."""
    final_df = result.get('final_df')
    poteaux_out = result.get('df_poteaux_out')
    verif_etudes = result.get('verif_etudes')
    key = 'c6bd'

    if final_df is not None and not final_df.empty:
        ws = _df_sheet(wb, "C6BD_ANALYSE", final_df, key)
        for col in range(1, len(final_df.columns) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 20
        if "Statut" in final_df.columns:
            ncols = len(final_df.columns)
            for idx, row_t in enumerate(final_df.itertuples(), start=2):
                statut = str(getattr(row_t, 'Statut', ''))
                if "ABSENT" in statut:
                    _fill_row(ws, idx, _P_WARN, ncols)

    if poteaux_out is not None and not poteaux_out.empty:
        ws_out = _df_sheet(wb, "C6BD_HORS_PERIM", poteaux_out, key)
        ncols = len(poteaux_out.columns)
        for r in range(2, len(poteaux_out) + 2):
            _fill_row(ws_out, r, _P_CRIT, ncols)

    if verif_etudes:
        sans_c6 = verif_etudes.get('etudes_sans_c6', [])
        c6_sans = verif_etudes.get('c6_sans_etude', [])
        ml = max(len(sans_c6), len(c6_sans), 1)
        df_v = pd.DataFrame({
            'ETUDES CAP FT SANS C6': sans_c6 + [''] * (ml - len(sans_c6)),
            'FICHIERS C6 SANS ETUDE': c6_sans + [''] * (ml - len(c6_sans)),
        })
        ws_v = _df_sheet(wb, "C6BD_VERIF_ETUDES", df_v, key)
        ws_v.column_dimensions['A'].width = 40
        ws_v.column_dimensions['B'].width = 40
        for r in range(2, ml + 2):
            for c in [1, 2]:
                if ws_v.cell(row=r, column=c).value:
                    ws_v.cell(row=r, column=c).fill = _P_WARN


# ======================================================================
#  Police C6
# ======================================================================

def write_police(wb, result):
    """Write Police C6 sheets: recap + detail per appui."""
    stats = result.get('stats', [])
    if not stats:
        return

    key = 'police_c6'

    # Recap
    ws = wb.create_sheet("PLC6_RECAP")
    hdrs = ["ETUDE", "APPUIS C6", "OK", "ECARTS", "ABSENTS BDD",
            "BOITIER ERR", "% CONFORMITE"]
    _init_sheet(ws, key, hdrs, [35, 14, 10, 10, 14, 14, 16])
    for r, s in enumerate(stats, 2):
        total = s.get('appuis_c6', 0)
        ok = s.get('nb_ok', 0)
        nok = s.get('nb_ecart', 0) + s.get('nb_absent', 0) + s.get('nb_boitier_err', 0)
        pct = (ok / total * 100) if total > 0 else 100
        _row(ws, r, [s['etude'], total, ok, s.get('nb_ecart', 0),
                      s.get('nb_absent', 0), s.get('nb_boitier_err', 0),
                      f"{pct:.0f}%"])
        for c in [2, 3, 4, 5, 6, 7]:
            ws.cell(row=r, column=c).alignment = _AL_C
        fill = _P_OK if nok == 0 else (_P_WARN if pct >= 80 else _P_ERR)
        ws.cell(row=r, column=7).fill = fill
        if nok > 0:
            _fill_row(ws, r, _P_ERR, 7)
            ws.cell(row=r, column=7).fill = fill

    # Detail
    ws_d = wb.create_sheet("PLC6_DETAIL")
    hdrs_d = ["ETUDE", "N APPUI", "CABLES C6", "NB C6", "NB BDD",
              "CAPA C6", "CAPA BDD", "STATUT", "DETAIL",
              "BOITIER C6", "BPE TYPE", "BOITIER STATUT"]
    _init_sheet(ws_d, key, hdrs_d,
                [35, 15, 40, 10, 10, 18, 18, 14, 50, 12, 15, 15])

    dr = 2
    for s in stats:
        etude = s['etude']
        for a in s.get('detail', []):
            capas_c6 = a.get('capas_c6', [])
            capas_bdd = a.get('capas_bdd', [])
            vals = [etude, a['num_appui'], a.get('cables_c6', ''),
                    a['nb_cables_c6'], a['nb_cables_bdd'],
                    '+'.join(str(c) for c in capas_c6) if capas_c6 else '',
                    '+'.join(str(c) for c in capas_bdd) if capas_bdd else '',
                    a['statut'], a.get('message', ''),
                    a.get('boitier_c6', ''), a.get('bpe_noe_type', ''),
                    a.get('boitier_statut', '')]
            st = a['statut']
            fill = _P_OK if st == 'OK' else (_P_WARN if st == 'ABSENT_BDD' else _P_ERR)
            _row(ws_d, dr, vals, fill=fill)

            bst = a.get('boitier_statut', '')
            if bst == 'ERREUR':
                for c in range(10, 13):
                    ws_d.cell(row=dr, column=c).fill = _P_CRIT
            elif bst == 'OK':
                for c in range(10, 13):
                    ws_d.cell(row=dr, column=c).fill = _P_OK
            dr += 1


# ======================================================================
#  C6-C3A
# ======================================================================

def _c6c3a_highlight(ws, df, check_cols):
    """Highlight rows containing ABSENT in any of the check columns."""
    ncols = len(df.columns)
    for idx in range(len(df)):
        row_data = df.iloc[idx]
        if any(col in row_data and row_data[col] == "ABSENT" for col in check_cols):
            _fill_row(ws, idx + 2, _P_WARN, ncols)


def write_c6c3a(wb, result):
    """Write C6-C3A analysis sheets."""
    key = 'c6c3a'

    df_final = result.get('df_final')
    if df_final is not None and not df_final.empty:
        ws = _df_sheet(wb, "C6C3A_ANALYSE", df_final, key)
        for c in range(1, len(df_final.columns) + 1):
            ws.column_dimensions[get_column_letter(c)].width = 25
        _c6c3a_highlight(ws, df_final,
                         ["inf_num (ETUDES_QGIS)", "inf_num (C3A)", "Excel (C6)"])

    df_rempl = result.get('df_final_rempl')
    if df_rempl is not None and not df_rempl.empty:
        ws2 = _df_sheet(wb, "C6C3A_REMPL", df_rempl, key)
        for c in range(1, len(df_rempl.columns) + 1):
            ws2.column_dimensions[get_column_letter(c)].width = 25
        _c6c3a_highlight(ws2, df_rempl,
                         ["inf_num (ETUDES_QGIS)", "inf_num (C3A)",
                          "Excel (C6)", "Fichier (C7) Excel"])


# ======================================================================
#  Main entry point
# ======================================================================

_WRITERS = {
    'maj': write_maj,
    'capft': write_capft,
    'comac': write_comac,
    'c6bd': write_c6bd,
    'police_c6': write_police,
    'c6c3a': write_c6c3a,
}


def generate_unified_report(batch_results, export_dir):
    """
    Generate a single Excel workbook from all module results.

    Args:
        batch_results: dict {module_key: result_dict}
        export_dir: directory to save the file

    Returns:
        str: path to the generated file
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"rapport_batch_{ts}.xlsx"
    filepath = os.path.join(export_dir, filename)

    wb = openpyxl.Workbook()

    # 1. Dashboard (first sheet)
    _write_dashboard(wb, batch_results)

    # 2. Module sheets
    for module_key in ['maj', 'capft', 'comac', 'c6bd', 'police_c6', 'c6c3a']:
        result = batch_results.get(module_key)
        if result is None:
            continue
        writer = _WRITERS.get(module_key)
        if writer:
            try:
                writer(wb, result)
            except (KeyError, ValueError, TypeError, AttributeError) as exc:
                QgsMessageLog.logMessage(
                    f"[REPORT] Erreur ecriture {module_key}: {exc}",
                    "PoleAerien", Qgis.Warning
                )

    wb.save(filepath)
    return filepath
