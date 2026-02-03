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


def run_async_task(task):
    """Submit task to QGIS task manager"""
    QgsApplication.taskManager().addTask(task)
    return task
