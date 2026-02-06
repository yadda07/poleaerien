# -*- coding: utf-8 -*-
"""
Async Tasks Module - Non-blocking execution with smooth progress.

Architecture: QgsTask + signals for UI updates.

THREADING SAFETY:
- All QGIS API calls (layer access, feature iteration) MUST happen on main thread
- Worker threads receive pre-extracted data via qgis_data parameter
- Only pure Python operations (Excel parsing, dict comparison) run in workers
- Use signals to communicate results back to main thread
"""

from qgis.PyQt.QtCore import QObject, pyqtSignal, QTimer
from qgis.core import QgsTask, QgsApplication, QgsMessageLog, Qgis
import traceback
import sip


def _sip_is_deleted(obj):
    try:
        return sip.isdeleted(obj)
    except RuntimeError:
        return True


class TaskSignals(QObject):
    """Signals for async task communication"""
    progress = pyqtSignal(int)           # Progress value 0-100
    message = pyqtSignal(str, str)       # (message, color)
    finished = pyqtSignal(dict)          # Result dict
    error = pyqtSignal(str)              # Error message


class SmoothProgressController(QObject):
    """
    Smooth progress interpolation for fluid UI feedback.
    Animates between target values instead of jumping.
    """
    
    valueChanged = pyqtSignal(int)
    
    def __init__(self, progress_bar=None, interval_ms=30, step=2):
        super().__init__()
        self._progress_bar = progress_bar
        self._current = 0
        self._target = 0
        self._step = step
        
        self._timer = QTimer(self)
        self._timer.setInterval(interval_ms)
        self._timer.timeout.connect(self._interpolate)
    
    def set_progress_bar(self, progress_bar):
        """Attach to a QProgressBar"""
        self._progress_bar = progress_bar
    
    def set_target(self, value):
        """Set target value, animation will interpolate to it (never regress)"""
        new_target = max(0, min(100, value))
        # Ne jamais régresser: garder le max entre l'ancienne et la nouvelle cible
        self._target = max(self._target, new_target)
        if not self._timer.isActive():
            self._timer.start()
    
    def set_immediate(self, value):
        """Set value immediately without animation"""
        self._current = value
        self._target = value
        self._update_bar()
    
    def reset(self):
        """Reset to 0"""
        self._timer.stop()
        self._current = 0
        self._target = 0
        self._update_bar()
    
    def _interpolate(self):
        """Smooth interpolation step (only forward, never regress)"""
        if self._current < self._target:
            self._current = min(self._current + self._step, self._target)
            self._update_bar()
        # Ne jamais régresser: ignorer si current > target
        if self._current >= self._target and self._target >= 100:
            self._timer.stop()
    
    def _update_bar(self):
        """Update progress bar and emit signal"""
        if self._progress_bar:
            self._progress_bar.setValue(int(self._current))
        self.valueChanged.emit(int(self._current))


class AsyncTaskBase(QgsTask):
    """
    Base class for async operations.
    Subclasses implement execute() method.
    """
    
    def __init__(self, name, params=None):
        super().__init__(name, QgsTask.CanCancel)
        self.params = params or {}
        self.signals = TaskSignals()
        self.result = {}
        self.exception = None
    
    def run(self):
        """Execute in background thread"""
        try:
            return self.execute()
        except ValueError as e:
            tb = traceback.format_exc()
            self.exception = str(e)
            QgsMessageLog.logMessage(f"Erreur validation: {e}\n{tb}", "PoleAerien", Qgis.Warning)
            return False
        except Exception as e:
            tb = traceback.format_exc()
            self.exception = f"{e}\n{tb}"
            QgsMessageLog.logMessage(f"Erreur task: {e}\n{tb}", "PoleAerien", Qgis.Critical)
            return False
    
    def execute(self):
        """Override in subclass - main logic"""
        raise NotImplementedError
    
    def finished(self, success):
        """Callback on main thread"""
        if success:
            self.signals.finished.emit(self.result)
        else:
            self.signals.error.emit(self.exception or "Annule")
    
    def cancel(self):
        """Override cancel to be completely safe during unload.
        
        QgsTask.cancel() is called by QGIS TaskManager during cleanup.
        The C++ object may already be deleted at this point.
        We simply pass to avoid any SIP access.
        """
        pass
    
    def emit_progress(self, value):
        """Helper to emit progress"""
        self.signals.progress.emit(value)
    
    def emit_message(self, msg, color="black"):
        """Helper to emit message"""
        self.signals.message.emit(msg, color)


class ExcelExportTask(AsyncTaskBase):
    """Async task for Excel export (openpyxl).

    Runs pure Python export code in background to avoid UI freezes.
    """

    def __init__(self, name, export_fn, args=None, kwargs=None, payload=None):
        super().__init__(name)
        self._export_fn = export_fn
        self._args = args or []
        self._kwargs = kwargs or {}
        self._payload = payload or {}

    def execute(self):
        if self.isCanceled():
            return False

        self.emit_progress(92)
        self.emit_message("Export Excel...", "grey")

        try:
            self._export_fn(*self._args, **self._kwargs)
        except Exception as e:
            raise RuntimeError(f"ExcelExportTask.execute: export echoue (cause: {e})") from e

        if self.isCanceled():
            return False

        self.emit_progress(99)
        self.result = dict(self._payload)
        self.result['export_done'] = True
        return True


class CapFtTask(AsyncTaskBase):
    """Async task for CAP_FT analysis.
    
    ARCH-FIX: Reçoit données pré-extraites du main thread.
    Ne fait que traitement pur (comparaison, lecture Excel).
    """
    
    def __init__(self, params, qgis_data=None):
        """
        Args:
            params: Configuration (chemins, fichier_export)
            qgis_data: Données extraites sur main thread:
                - doublons, hors_etude (validation)
                - dico_qgis (poteaux QGIS)
        """
        super().__init__("Analyse CAP_FT", params)
        self.qgis_data = qgis_data or {}
    
    def execute(self):
        if self.isCanceled():
            return False
        
        self.emit_progress(5)
        self.emit_message("Verification donnees...", "grey")
        
        # Données pré-extraites sur main thread
        doublons = self.qgis_data.get('doublons', [])
        hors_etude = self.qgis_data.get('hors_etude', [])
        
        if self.isCanceled():
            return False
        
        if doublons:
            self.result = {'error_type': 'doublons', 'data': doublons}
            return True
        
        if hors_etude:
            self.result = {'error_type': 'hors_etude', 'data': hors_etude}
            return True
        
        self.emit_progress(15)
        self.emit_message("Donnees QGIS chargees...", "grey")
        
        dico_qgis = self.qgis_data.get('dico_qgis', {})
        dico_poteaux_prives = self.qgis_data.get('dico_poteaux_prives', {})
        
        if self.isCanceled():
            return False
        
        self.emit_progress(40)
        self.emit_message("Lecture fichiers Excel...", "grey")
        
        # Lecture Excel - thread-safe (pas d'appel QGIS)
        from .CapFt import CapFt
        cap = CapFt()
        dico_excel = cap.LectureFichiersExcelsCap_ft(self.params['chemin_cap_ft'])
        
        if self.isCanceled():
            return False
        
        self.emit_progress(60)
        self.emit_message("Comparaison donnees...", "grey")
        
        # Traitement pur - thread-safe
        resultats = cap.traitementResultatFinauxCapFt(dico_qgis, dico_excel)
        
        if self.isCanceled():
            return False
        
        self.emit_progress(90)
        
        self.result = {
            'success': True,
            'resultats': resultats,
            'dico_excel_introuvable': resultats[0],
            'dico_qgis_introuvable': resultats[1],
            'fichier_export': self.params['fichier_export'],
            'dico_qgis': dico_qgis,
            'dico_poteaux_prives': dico_poteaux_prives,
            'dico_excel': dico_excel,
            'pending_export': True
        }
        return True


class ComacTask(AsyncTaskBase):
    """Async task for COMAC analysis.
    
    ARCH-FIX: Reçoit données pré-extraites du main thread.
    Ne fait que traitement pur (comparaison, lecture Excel).
    """
    
    def __init__(self, params, qgis_data=None):
        """
        Args:
            params: Configuration (chemins, fichier_export)
            qgis_data: Données extraites sur main thread:
                - doublons, hors_etude (validation)
                - dico_qgis (poteaux QGIS)
        """
        super().__init__("Analyse COMAC", params)
        self.qgis_data = qgis_data or {}
    
    def execute(self):
        if self.isCanceled():
            return False
        
        self.emit_progress(5)
        self.emit_message("Verification donnees...", "grey")
        
        # Données pré-extraites sur main thread
        doublons = self.qgis_data.get('doublons', [])
        hors_etude = self.qgis_data.get('hors_etude', [])
        
        if self.isCanceled():
            return False
        
        if doublons:
            self.result = {'error_type': 'doublons', 'data': doublons}
            return True
        
        if hors_etude:
            self.result = {'error_type': 'hors_etude', 'data': hors_etude}
            return True
        
        self.emit_progress(15)
        self.emit_message("Donnees QGIS chargees...", "grey")
        
        dico_qgis = self.qgis_data.get('dico_qgis', {})
        dico_poteaux_prives = self.qgis_data.get('dico_poteaux_prives', {})
        
        if self.isCanceled():
            return False
        
        self.emit_progress(40)
        self.emit_message("Lecture fichiers Excel...", "grey")
        
        # Lecture Excel - thread-safe (pas d'appel QGIS)
        from .Comac import Comac
        com = Comac()
        doublons, erreurs_lecture, dico_excel, dico_verif_secu = com.LectureFichiersExcelsComac(
            self.params['chemin_comac'],
            self.params.get('zone_climatique', 'ZVN')
        )
        
        if self.isCanceled():
            return False
        
        if doublons:
            self.result = {'error_type': 'fichiers_doublons', 'data': doublons}
            return True
        
        if erreurs_lecture:
            self.result = {'error_type': 'erreur_lecture', 'data': erreurs_lecture}
            return True
        
        if self.isCanceled():
            return False
        
        self.emit_progress(60)
        self.emit_message("Comparaison donnees...", "grey")
        
        # Traitement pur - thread-safe
        resultats = com.traitementResultatFinaux(dico_qgis, dico_excel)
        
        if self.isCanceled():
            return False
        
        self.emit_progress(90)
        
        self.result = {
            'success': True,
            'resultats': resultats,
            'dico_excel_introuvable': resultats[0],
            'dico_qgis_introuvable': resultats[1],
            'fichier_export': self.params['fichier_export'],
            'dico_qgis': dico_qgis,
            'dico_poteaux_prives': dico_poteaux_prives,
            'dico_excel': dico_excel,
            'dico_verif_secu': dico_verif_secu,
            'pending_export': True
        }
        return True


class C6BdTask(AsyncTaskBase):
    """Async task for C6 vs BD comparison.
    
    ARCH-FIX: Reçoit données pré-extraites du main thread.
    Ne fait que traitement pur (lecture Excel, comparaison pandas).
    """
    
    def __init__(self, params, qgis_data=None):
        """
        Args:
            params: Configuration (chemins, fichier_export)
            qgis_data: Données extraites sur main thread:
                - df_qgis (DataFrame poteaux QGIS)
        """
        super().__init__("Comparaison C6 vs BD", params)
        self.qgis_data = qgis_data or {}
    
    def execute(self):
        if self.isCanceled():
            return False
        
        self.emit_progress(10)
        self.emit_message("Lecture fichiers C6...", "grey")
        
        # Lecture Excel - thread-safe (pas d'appel QGIS)
        import pandas as pd
        import os
        from .C6_vs_Bd import C6_vs_Bd
        from .core_utils import normalize_appui_num
        
        c6bd = C6_vs_Bd()
        df = pd.DataFrame(columns=["N° appui", "Nature des travaux", "Études", "Excel"])
        
        try:
            df1 = c6bd.LectureFichiersExcelsC6(df, self.params['repertoire_c6'])
        except Exception as e:
            raise RuntimeError(f"Lecture C6 echouee: {e}") from e
        
        if self.isCanceled():
            return False
        
        self.emit_progress(40)
        self.emit_message("Fusion donnees...", "grey")
        
        # Données pré-extraites sur main thread
        df2 = self.qgis_data.get('df_qgis')
        if df2 is None:
            raise ValueError("Donnees QGIS manquantes")
        
        # Traitement pur pandas - thread-safe
        def _uniq_join(sr):
            vals = [str(v).strip() for v in sr.tolist() 
                    if v is not None and str(v).strip() and str(v).strip() != "nan"]
            if not vals:
                return ""
            return ", ".join(sorted(set(vals)))
        
        df1 = df1.copy()
        df2 = df2.copy()
        
        df1["appui_key"] = df1["N° appui"].apply(normalize_appui_num)
        df2["appui_key"] = df2["N° appui"].apply(normalize_appui_num)
        
        df1 = df1[df1["appui_key"].astype(bool)]
        df2 = df2[df2["appui_key"].astype(bool)]
        
        if self.isCanceled():
            return False
        
        self.emit_progress(55)
        
        df1_g = (
            df1.groupby("appui_key", dropna=False)
            .agg({
                "N° appui": "first",
                "Excel": _uniq_join,
                "Études": _uniq_join,
                "Nature des travaux": _uniq_join,
            })
            .reset_index()
            .rename(columns={
                "Excel": "Fichiers (Excel)",
                "Études": "Études (Excel)",
                "Nature des travaux": "Nature (Excel)",
            })
        )
        
        df2_g = (
            df2.groupby("appui_key", dropna=False)
            .agg({
                "inf_num (QGIS)": "first",
                "Études": _uniq_join,
                "Nature des travaux": _uniq_join,
                "Etat": _uniq_join,
            })
            .reset_index()
            .rename(columns={
                "Études": "Études (QGIS)",
                "Nature des travaux": "Nature (QGIS)",
            })
        )
        
        if self.isCanceled():
            return False
        
        self.emit_progress(70)
        self.emit_message("Calcul statuts...", "grey")
        
        final_df = pd.merge(df1_g, df2_g, on=["appui_key"], how="outer")
        # Convertir en string avant fillna pour éviter erreur de type Int64 vs string
        final_df["N° appui"] = final_df["N° appui"].astype(str).replace('<NA>', '').replace('nan', '')
        final_df["N° appui"] = final_df.apply(
            lambda r: r["N° appui"] if r["N° appui"] else str(r["inf_num (QGIS)"] or ""), axis=1
        )
        
        def _statut(row):
            def _val(v):
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return ""
                s = str(v).strip()
                return "" if s.lower() == "nan" else s
            
            excel_val = _val(row.get("Fichiers (Excel)"))
            qgis_val = _val(row.get("inf_num (QGIS)"))
            has_excel = bool(excel_val)
            has_qgis = bool(qgis_val)
            
            if has_excel and has_qgis:
                et_x = set([v.strip() for v in _val(row.get("Études (Excel)")).split(",") if v.strip()])
                et_q = set([v.strip() for v in _val(row.get("Études (QGIS)")).split(",") if v.strip()])
                if et_x and et_q and et_x.isdisjoint(et_q):
                    return "A VERIFIER"
                return "OK"
            if has_excel and not has_qgis:
                return "ABSENT QGIS"
            if has_qgis and not has_excel:
                return "ABSENT EXCEL"
            return "ABSENT"
        
        final_df["Statut"] = final_df.apply(_statut, axis=1)
        final_df = final_df.sort_values(by=["Statut", "appui_key"], ascending=[True, True])
        
        cols = [
            "N° appui", "appui_key", "Statut",
            "Fichiers (Excel)", "Études (Excel)", "Nature (Excel)",
            "inf_num (QGIS)", "Études (QGIS)", "Nature (QGIS)", "Etat",
        ]
        final_df = final_df[[c for c in cols if c in final_df.columns]]
        
        if self.isCanceled():
            return False
        
        self.emit_progress(85)
        
        self.result = {
            'success': True,
            'final_df': final_df,
            'fichier_export': self.params['fichier_export'],
            'df_poteaux_out': self.qgis_data.get('df_poteaux_out'),
            'verif_etudes': self.qgis_data.get('verif_etudes'),
            'pending_export': True
        }
        return True


class PoliceC6Task(AsyncTaskBase):
    """Async task for Police C6 analysis.
    
    ARCHITECTURE:
    - Main thread: Extract QGIS data (appuis, SRO)
    - Worker thread: PostgreSQL query (fddcpi2), data comparison, Excel export
    
    Thread-safe: Ne fait aucun appel QGIS dans execute().
    """
    
    def __init__(self, params, qgis_data=None):
        """
        Args:
            params: Configuration
                - c6_files: Liste de (etude_name, fichier_c6)
                - export_path: Chemin export Excel
                - sro: Code SRO pour fddcpi2
            qgis_data: Données pré-extraites sur main thread
                - appuis: Liste de {'num_appui': str, 'geom': WKT/coords}
        """
        super().__init__("Analyse Police C6", params)
        self.qgis_data = qgis_data or {}
        self._cancelled = False
    
    def execute(self):
        if self.isCanceled():
            return False
        
        self.emit_progress(5)
        self.emit_message("Initialisation analyse Police C6...", "grey")
        
        import os
        
        # Import ici pour éviter import circulaire
        from .PoliceC6 import PoliceC6
        from .db_connection import DatabaseConnection
        from .cable_analyzer import compter_cables_par_appui
        
        from qgis.core import QgsGeometry
        
        c6_files = self.params.get('c6_files', [])
        sro = self.params.get('sro', '')
        export_path = self.params.get('export_path', '')
        
        # Désérialiser WKB → QgsGeometry (thread-safe)
        appuis_data = []
        for appui in self.qgis_data.get('appuis', []):
            geom = None
            wkb = appui.get('geom_wkb')
            if wkb:
                geom = QgsGeometry()
                geom.fromWkb(wkb)
            appuis_data.append({
                'num_appui': appui.get('num_appui', ''),
                'feature_id': appui.get('feature_id'),
                'geom': geom
            })
        
        if not c6_files:
            self.result = {'error': 'Aucun fichier C6 fourni'}
            return True
        
        if not sro:
            self.result = {'error': 'SRO non défini'}
            return True
        
        if self.isCanceled():
            return False
        
        # 1. Connexion PostgreSQL et récupération câbles via DatabaseConnection
        self.emit_progress(10)
        self.emit_message(f"Connexion PostgreSQL et appel fddcpi2({sro})...", "grey")
        
        try:
            db = DatabaseConnection()
            if not db.connect():
                self.result = {'error': 'Connexion PostgreSQL impossible'}
                return True
            
            cables_all = db.execute_fddcpi2(sro)
            
            # Garder câbles de distribution aériens + façade (cab_type='CDI' + posemode 1 ou 2)
            cables = [c for c in cables_all if c.cab_type == 'CDI' and c.posemode in (1, 2)]
            nb_exclu = len(cables_all) - len(cables)
            
            self.emit_message(
                f"  {len(cables)} câbles CDI aériens/façade retenus ({nb_exclu} exclus sur {len(cables_all)})",
                "blue"
            )
            
            # Récupérer les BPE du SRO pour vérification boîtier
            bpe_list = db.query_bpe_by_sro(sro)
            db.disconnect()
            
            # Préparer géométries BPE pour matching spatial
            bpe_geoms = []
            for bpe in bpe_list:
                if bpe['geom_wkt']:
                    g = QgsGeometry.fromWkt(bpe['geom_wkt'])
                    if g and not g.isNull():
                        bpe_geoms.append({'geom': g, 'noe_type': bpe['noe_type'], 'gid': bpe['gid']})
            
            self.emit_message(
                f"  {len(bpe_geoms)} BPE récupérés pour vérification boîtier",
                "blue"
            )
        except Exception as e:
            self.result = {'error': f'Erreur fddcpi2: {e}'}
            return True
        
        if self.isCanceled():
            return False
        
        # 2. Traitement de chaque étude
        self.emit_progress(20)
        stats_globales = []
        police = PoliceC6()
        
        total = len(c6_files)
        for i, (etude_name, c6_file) in enumerate(c6_files):
            if self.isCanceled():
                break
            
            progress = 20 + int((i / total) * 60)
            self.emit_progress(progress)
            self.emit_message(f"[>] {etude_name}...", "blue")
            
            try:
                # Lire C6
                donnees_c6, _, boitier_c6 = police.lire_annexe_c6(c6_file)
                if not donnees_c6:
                    self.emit_message(f"  [!] C6 vide ou illisible", "orange")
                    continue
                
                total_cables_c6 = sum(len(v) for v in donnees_c6.values())
                nb_boitier_c6 = len(boitier_c6)
                self.emit_message(
                    f"  C6: {len(donnees_c6)} appuis, {total_cables_c6} câbles, {nb_boitier_c6} boîtiers",
                    "grey"
                )
                
                # Compter câbles par appui (tolérance 0.5m)
                cables_par_appui = compter_cables_par_appui(cables, appuis_data, tolerance=0.5)
                
                # Vérifier boîtiers: pour chaque appui avec PB/PEO, chercher BPE dans rayon 1m
                verif_boitier = self._verifier_boitiers(
                    boitier_c6, appuis_data, bpe_geoms
                )
                
                # Comparer C6 vs câbles BDD
                comparaison = police.comparer_c6_cables(donnees_c6, cables_par_appui)
                
                # Injecter résultats boîtier dans comparaison
                for entry in comparaison:
                    num = entry['num_appui']
                    bv = verif_boitier.get(num, {})
                    entry['boitier_c6'] = bv.get('boitier_c6', '')
                    entry['bpe_trouve'] = bv.get('bpe_trouve', False)
                    entry['bpe_noe_type'] = bv.get('bpe_noe_type', '')
                    entry['boitier_statut'] = bv.get('statut', '')
                
                nb_ok = sum(1 for c in comparaison if c['statut'] == 'OK')
                nb_ecart = sum(1 for c in comparaison if c['statut'] == 'ECART')
                nb_absent = sum(1 for c in comparaison if c['statut'] == 'ABSENT_BDD')
                nb_boitier_err = sum(1 for c in comparaison if c.get('boitier_statut') == 'ERREUR')
                
                self.emit_message(
                    f"  Résultat: {nb_ok} OK, {nb_ecart} écarts, {nb_absent} absents BDD"
                    f"{f', {nb_boitier_err} boîtier(s) absent(s)' if nb_boitier_err else ''}",
                    "green" if nb_ecart == 0 and nb_absent == 0 and nb_boitier_err == 0 else "orange"
                )
                
                # Lister les appuis en anomalie
                for c in comparaison:
                    if c['statut'] != 'OK':
                        self.emit_message(
                            f"    [{c['statut']}] Appui {c['num_appui']}: "
                            f"C6={c['nb_cables_c6']} câbles, BDD={c['nb_cables_bdd']} | {c['message']}",
                            "red" if c['statut'] == 'ABSENT_BDD' else "orange"
                        )
                    if c.get('boitier_statut') == 'ERREUR':
                        self.emit_message(
                            f"    [BOITIER] Appui {c['num_appui']}: "
                            f"C6={c['boitier_c6']} mais aucun BPE trouvé dans rayon 1m",
                            "red"
                        )
                
                stats_globales.append({
                    'etude': etude_name,
                    'appuis_c6': len(donnees_c6),
                    'nb_ok': nb_ok,
                    'nb_ecart': nb_ecart,
                    'nb_absent': nb_absent,
                    'nb_boitier_err': nb_boitier_err,
                    'detail': comparaison
                })
                
            except Exception as e:
                self.emit_message(f"  [X] Erreur: {e}", "red")
                QgsMessageLog.logMessage(f"Erreur {etude_name}: {e}", "PoleAerien", Qgis.Warning)
        
        if self.isCanceled():
            return False
        
        # 3. Export Excel
        self.emit_progress(85)
        
        if export_path and stats_globales:
            self.emit_message("Export Excel...", "grey")
            try:
                excel_file = self._export_excel(stats_globales, export_path)
                self.emit_message(f"✓ Export: {os.path.basename(excel_file)}", "green")
            except Exception as e:
                self.emit_message(f"[!] Erreur export: {e}", "orange")
        
        self.emit_progress(100)
        
        # Sérialiser câbles pour chargement couche QGIS (main thread)
        cables_for_layer = []
        for cable in cables:
            cables_for_layer.append({
                'gid_dc2': cable.gid_dc2,
                'gid_dc': cable.gid_dc,
                'cab_capa': cable.cab_capa,
                'cab_type': cable.cab_type,
                'cb_etiquet': cable.cb_etiquet,
                'posemode': cable.posemode,
                'length': cable.length,
                'geom_wkt': cable.geom_wkt
            })
        
        self.result = {
            'success': True,
            'stats': stats_globales,
            'total_etudes': len(c6_files),
            'etudes_traitees': len(stats_globales),
            'cables': cables_for_layer,
            'sro': sro
        }
        return True
    
    def _verifier_boitiers(self, boitier_c6, appuis_data, bpe_geoms):
        """
        Vérifie la présence de BPE pour chaque appui déclarant un boîtier en C6.
        
        Args:
            boitier_c6: Dict[num_appui, str] - type boîtier C6 (PB/PEO)
            appuis_data: Liste de dicts avec 'num_appui' et 'geom' (QgsGeometry)
            bpe_geoms: Liste de dicts avec 'geom', 'noe_type', 'gid'
        
        Returns:
            Dict[num_appui, {boitier_c6, bpe_trouve, bpe_noe_type, statut}]
        """
        result = {}
        
        if not boitier_c6:
            return result
        
        # Index appuis par numéro pour accès rapide
        appuis_by_num = {}
        for appui in appuis_data:
            num = appui.get('num_appui', '')
            if num and appui.get('geom'):
                appuis_by_num[num] = appui['geom']
        
        for num_appui, type_boitier in boitier_c6.items():
            appui_geom = appuis_by_num.get(num_appui)
            
            entry = {
                'boitier_c6': type_boitier,
                'bpe_trouve': False,
                'bpe_noe_type': '',
                'statut': ''
            }
            
            if not appui_geom:
                entry['statut'] = 'ERREUR'
                entry['bpe_noe_type'] = 'appui non localisé'
                result[num_appui] = entry
                continue
            
            # Chercher BPE dans un rayon de 1m
            bpe_proche = None
            dist_min = 999999
            for bpe in bpe_geoms:
                dist = appui_geom.distance(bpe['geom'])
                if dist < 1.0 and dist < dist_min:
                    dist_min = dist
                    bpe_proche = bpe
            
            if bpe_proche:
                entry['bpe_trouve'] = True
                entry['bpe_noe_type'] = bpe_proche['noe_type']
                entry['statut'] = 'OK'
            else:
                entry['statut'] = 'ERREUR'
            
            result[num_appui] = entry
        
        return result

    def _export_excel(self, stats, export_path):
        """Export des résultats vers Excel avec récap + détail par appui."""
        import os
        from datetime import datetime
        
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment
        except ImportError:
            raise RuntimeError("openpyxl non installé")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if os.path.isdir(export_path):
            filename = os.path.join(export_path, f"PoliceC6_Export_{timestamp}.xlsx")
        else:
            filename = export_path
        
        wb = openpyxl.Workbook()
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        ok_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        ecart_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        absent_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        
        # === Feuille 1: Récapitulatif par étude ===
        ws_recap = wb.active
        ws_recap.title = "Récapitulatif"
        
        recap_headers = ["Étude", "Appuis C6", "OK", "Écarts", "Absents BDD", "Boîtier absent"]
        for col, header in enumerate(recap_headers, 1):
            cell = ws_recap.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
        
        for row, stat in enumerate(stats, 2):
            ws_recap.cell(row=row, column=1, value=stat['etude'])
            ws_recap.cell(row=row, column=2, value=stat['appuis_c6'])
            ws_recap.cell(row=row, column=3, value=stat.get('nb_ok', 0))
            ws_recap.cell(row=row, column=4, value=stat.get('nb_ecart', 0))
            ws_recap.cell(row=row, column=5, value=stat.get('nb_absent', 0))
            ws_recap.cell(row=row, column=6, value=stat.get('nb_boitier_err', 0))
            
            # Coloriser si anomalies
            has_anomaly = (stat.get('nb_ecart', 0) > 0 or stat.get('nb_absent', 0) > 0
                           or stat.get('nb_boitier_err', 0) > 0)
            if has_anomaly:
                for c in range(1, 7):
                    ws_recap.cell(row=row, column=c).fill = ecart_fill
        
        # Ajuster largeurs
        ws_recap.column_dimensions['A'].width = 35
        for col_letter in ['B', 'C', 'D', 'E', 'F']:
            ws_recap.column_dimensions[col_letter].width = 15
        
        # === Feuille 2: Détail par appui ===
        ws_detail = wb.create_sheet("Détail par appui")
        
        detail_headers = [
            "Étude", "N° Appui", "Câbles C6", "Nb C6",
            "Nb BDD", "Capa C6", "Capa BDD", "Statut", "Détail",
            "Boîtier C6", "BPE Type", "Boîtier Statut"
        ]
        for col, header in enumerate(detail_headers, 1):
            cell = ws_detail.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
        
        boitier_err_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
        
        detail_row = 2
        for stat in stats:
            etude = stat['etude']
            for appui in stat.get('detail', []):
                capas_c6 = appui.get('capas_c6', [])
                capas_bdd = appui.get('capas_bdd', [])
                
                ws_detail.cell(row=detail_row, column=1, value=etude)
                ws_detail.cell(row=detail_row, column=2, value=appui['num_appui'])
                ws_detail.cell(row=detail_row, column=3, value=appui.get('cables_c6', ''))
                ws_detail.cell(row=detail_row, column=4, value=appui['nb_cables_c6'])
                ws_detail.cell(row=detail_row, column=5, value=appui['nb_cables_bdd'])
                ws_detail.cell(row=detail_row, column=6, value='+'.join(str(c) for c in capas_c6) if capas_c6 else '')
                ws_detail.cell(row=detail_row, column=7, value='+'.join(str(c) for c in capas_bdd) if capas_bdd else '')
                ws_detail.cell(row=detail_row, column=8, value=appui['statut'])
                ws_detail.cell(row=detail_row, column=9, value=appui.get('message', ''))
                ws_detail.cell(row=detail_row, column=10, value=appui.get('boitier_c6', ''))
                ws_detail.cell(row=detail_row, column=11, value=appui.get('bpe_noe_type', ''))
                ws_detail.cell(row=detail_row, column=12, value=appui.get('boitier_statut', ''))
                
                # Coloriser selon statut
                if appui['statut'] == 'OK':
                    fill = ok_fill
                elif appui['statut'] == 'ABSENT_BDD':
                    fill = absent_fill
                else:
                    fill = ecart_fill
                
                for c in range(1, 10):
                    ws_detail.cell(row=detail_row, column=c).fill = fill
                
                # Coloriser colonnes boîtier si erreur
                if appui.get('boitier_statut') == 'ERREUR':
                    for c in range(10, 13):
                        ws_detail.cell(row=detail_row, column=c).fill = boitier_err_fill
                elif appui.get('boitier_statut') == 'OK':
                    for c in range(10, 13):
                        ws_detail.cell(row=detail_row, column=c).fill = ok_fill
                
                detail_row += 1
        
        # Ajuster largeurs
        ws_detail.column_dimensions['A'].width = 35
        ws_detail.column_dimensions['B'].width = 15
        ws_detail.column_dimensions['C'].width = 40
        ws_detail.column_dimensions['D'].width = 10
        ws_detail.column_dimensions['E'].width = 10
        ws_detail.column_dimensions['F'].width = 18
        ws_detail.column_dimensions['G'].width = 18
        ws_detail.column_dimensions['H'].width = 15
        ws_detail.column_dimensions['I'].width = 55
        ws_detail.column_dimensions['J'].width = 12
        ws_detail.column_dimensions['K'].width = 15
        ws_detail.column_dimensions['L'].width = 15
        
        wb.save(filename)
        return filename


def run_async_task(task):
    """Submit task to QGIS task manager"""
    QgsApplication.taskManager().addTask(task)
    return task
