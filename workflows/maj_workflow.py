# -*- coding: utf-8 -*-
"""
Workflow pour la mise à jour FT/BT.
Orchestre l'extraction des données et l'exécution asynchrone.
"""

import time
from qgis.PyQt.QtCore import QObject, pyqtSignal, QTimer
from qgis.core import QgsApplication, Qgis, QgsMessageLog, QgsFeatureRequest, QgsExpression, NULL
from ..Maj_Ft_Bt import MajFtBt, MajFtBtTask, MajUpdateTask
from ..maj_sql_background import MajSqlBackgroundTask, get_layer_db_uri, reload_layer
from ..qgis_utils import validate_same_crs

class MajWorkflow(QObject):
    """
    Contrôleur pour le flux de Mise à Jour FT/BT.
    Gère:
    1. L'extraction des données QGIS (Main Thread, non-bloquant via QTimer)
    2. Le lancement de la tâche asynchrone (Worker Thread)
    3. La communication des résultats via signaux
    """
    
    # Signaux relayés vers l'UI
    progress_changed = pyqtSignal(int)
    message_received = pyqtSignal(str, str)  # message, couleur
    analysis_finished = pyqtSignal(dict)
    analysis_error = pyqtSignal(str)
    
    # CRITICAL-001 FIX: Signaux pour MAJ BD asynchrone
    update_finished = pyqtSignal(dict)  # {'ft_updated': int, 'bt_updated': int}
    update_error = pyqtSignal(str)
    
    BATCH_SIZE = 50  # Features par lot
    UPDATE_BATCH_SIZE = 1  # MAJ BD par lot (UI responsive)
    
    # Cache statique partagé entre instances
    _coords_cache = None
    _cache_key = None
    
    def __init__(self):
        super().__init__()
        self.maj_logic = MajFtBt()
        self.current_task = None
        self.update_task = None
        self._update_state = None
        
        # État extraction incrémentale
        self._extraction_state = None
    
    @staticmethod
    def _compute_cache_key(lyr_pot, lyr_cap, lyr_com):
        """Calcule une clé de cache basée sur les couches."""
        return (
            lyr_pot.id(), lyr_pot.featureCount(),
            lyr_cap.id(), lyr_cap.featureCount(),
            lyr_com.id(), lyr_com.featureCount()
        )
    
    @staticmethod
    def clear_cache():
        """Vide le cache (à appeler si les données changent)."""
        MajWorkflow._coords_cache = None
        MajWorkflow._cache_key = None

    def start_analysis(self, lyr_pot, lyr_cap, lyr_com, excel_path):
        """Lance l'analyse avec extraction non-bloquante (avec cache)."""
        if not lyr_pot or not lyr_cap or not lyr_com:
            self.analysis_error.emit("Couches invalides ou manquantes")
            return

        try:
            self.maj_logic.valider_nom_fichier(excel_path)
        except ValueError as e:
            self.analysis_error.emit(str(e))
            return

        try:
            validate_same_crs(lyr_pot, lyr_cap, "MAJ_FT_BT")
            validate_same_crs(lyr_pot, lyr_com, "MAJ_FT_BT")
        except ValueError as e:
            self.analysis_error.emit(str(e))
            return

        # Vérifier le cache
        cache_key = self._compute_cache_key(lyr_pot, lyr_cap, lyr_com)
        if MajWorkflow._cache_key == cache_key and MajWorkflow._coords_cache is not None:
            # CACHE HIT - utiliser les données en cache (instantané!)
            QgsMessageLog.logMessage("[PERF-MAJ] ======= CACHE HIT - extraction instantanée =======", "PoleAerien", Qgis.Info)
            self.message_received.emit("Données en cache...", "green")
            self.progress_changed.emit(15)
            
            # Lancer directement la tâche async avec les données en cache
            params = {'fichier_excel': excel_path}
            self.current_task = MajFtBtTask(params, MajWorkflow._coords_cache)
            self.current_task.signals.progress.connect(self.progress_changed)
            self.current_task.signals.message.connect(self.message_received)
            self.current_task.signals.finished.connect(self.analysis_finished)
            self.current_task.signals.error.connect(self.analysis_error)
            QgsApplication.taskManager().addTask(self.current_task)
            return

        # CACHE MISS - faire l'extraction complète
        QgsMessageLog.logMessage("[PERF-MAJ] ======= EXTRACTION COORDS (non-bloquant) =======", "PoleAerien", Qgis.Info)
        
        # Initialiser l'état d'extraction
        self._extraction_state = {
            't0': time.perf_counter(),
            'excel_path': excel_path,
            'cache_key': cache_key,
            'lyr_pot': lyr_pot,
            'lyr_cap': lyr_cap,
            'lyr_com': lyr_com,
            'phase': 'pot_ft',
            'iterator': None,
            'poteaux_ft': [],
            'poteaux_bt': [],
            'etudes_cap_ft': [],
            'etudes_comac': []
        }
        
        self.message_received.emit("Extraction POT-FT...", "grey")
        self.progress_changed.emit(2)
        
        # Créer l'iterator pour POT-FT
        self._extraction_state['iterator'] = lyr_pot.getFeatures(
            QgsFeatureRequest(QgsExpression("inf_type LIKE 'POT-FT'"))
        )
        self._extraction_state['t_phase'] = time.perf_counter()
        
        # Démarrer extraction par lots via QTimer
        QTimer.singleShot(0, self._process_extraction_batch)

    def _process_extraction_batch(self):
        """Traite un lot de features puis cède le contrôle à l'event loop."""
        state = self._extraction_state
        if state is None:
            return
        
        phase = state['phase']
        iterator = state['iterator']
        count = 0
        
        try:
            while count < self.BATCH_SIZE:
                try:
                    feat = next(iterator)
                except StopIteration:
                    # Phase terminée, passer à la suivante
                    self._finish_phase()
                    return
                
                if phase == 'pot_ft':
                    if feat.hasGeometry():
                        pt = feat.geometry().asPoint()
                        state['poteaux_ft'].append({
                            'x': pt.x(), 'y': pt.y(),
                            'gid': feat["gid"], 'inf_num': feat["inf_num"]
                        })
                elif phase == 'pot_bt':
                    if feat.hasGeometry():
                        pt = feat.geometry().asPoint()
                        state['poteaux_bt'].append({
                            'x': pt.x(), 'y': pt.y(),
                            'gid': feat["gid"], 'inf_num': feat["inf_num"]
                        })
                elif phase == 'cap_ft':
                    if feat.hasGeometry():
                        raw_etude = feat["nom_etudes"]
                        etude = str(raw_etude).upper() if raw_etude and raw_etude != NULL else ""
                        geom = feat.geometry()
                        if geom.isMultipart():
                            poly = geom.asMultiPolygon()[0][0] if geom.asMultiPolygon() else []
                        else:
                            poly = geom.asPolygon()[0] if geom.asPolygon() else []
                        vertices = [(p.x(), p.y()) for p in poly]
                        bbox = geom.boundingBox()
                        state['etudes_cap_ft'].append({
                            'nom_etudes': etude, 'vertices': vertices,
                            'bbox': (bbox.xMinimum(), bbox.yMinimum(), bbox.xMaximum(), bbox.yMaximum())
                        })
                elif phase == 'comac':
                    if feat.hasGeometry():
                        raw_etude = feat["nom_etudes"]
                        etude = str(raw_etude).upper() if raw_etude and raw_etude != NULL else ""
                        geom = feat.geometry()
                        if geom.isMultipart():
                            poly = geom.asMultiPolygon()[0][0] if geom.asMultiPolygon() else []
                        else:
                            poly = geom.asPolygon()[0] if geom.asPolygon() else []
                        vertices = [(p.x(), p.y()) for p in poly]
                        bbox = geom.boundingBox()
                        state['etudes_comac'].append({
                            'nom_etudes': etude, 'vertices': vertices,
                            'bbox': (bbox.xMinimum(), bbox.yMinimum(), bbox.xMaximum(), bbox.yMaximum())
                        })
                count += 1
            
            # Continuer avec le prochain lot après avoir cédé à l'event loop
            QTimer.singleShot(0, self._process_extraction_batch)
            
        except Exception as e:
            self.analysis_error.emit(f"Erreur extraction: {e}")
            self._extraction_state = None

    def _finish_phase(self):
        """Termine une phase et passe à la suivante."""
        state = self._extraction_state
        phase = state['phase']
        t_phase = time.perf_counter() - state['t_phase']
        
        if phase == 'pot_ft':
            QgsMessageLog.logMessage(f"[PERF-MAJ] >> POT-FT: {len(state['poteaux_ft'])} en {t_phase:.3f}s", "PoleAerien", Qgis.Info)
            state['phase'] = 'pot_bt'
            state['iterator'] = state['lyr_pot'].getFeatures(
                QgsFeatureRequest(QgsExpression("inf_type LIKE 'POT-BT'"))
            )
            state['t_phase'] = time.perf_counter()
            self.message_received.emit("Extraction POT-BT...", "grey")
            self.progress_changed.emit(8)
            QTimer.singleShot(0, self._process_extraction_batch)
            
        elif phase == 'pot_bt':
            QgsMessageLog.logMessage(f"[PERF-MAJ] >> POT-BT: {len(state['poteaux_bt'])} en {t_phase:.3f}s", "PoleAerien", Qgis.Info)
            state['phase'] = 'cap_ft'
            state['iterator'] = state['lyr_cap'].getFeatures()
            state['t_phase'] = time.perf_counter()
            self.message_received.emit("Extraction CAP_FT...", "grey")
            self.progress_changed.emit(12)
            QTimer.singleShot(0, self._process_extraction_batch)
            
        elif phase == 'cap_ft':
            QgsMessageLog.logMessage(f"[PERF-MAJ] >> CAP_FT: {len(state['etudes_cap_ft'])} en {t_phase:.3f}s", "PoleAerien", Qgis.Info)
            state['phase'] = 'comac'
            state['iterator'] = state['lyr_com'].getFeatures()
            state['t_phase'] = time.perf_counter()
            self.message_received.emit("Extraction COMAC...", "grey")
            self.progress_changed.emit(16)
            QTimer.singleShot(0, self._process_extraction_batch)
            
        elif phase == 'comac':
            QgsMessageLog.logMessage(f"[PERF-MAJ] >> COMAC: {len(state['etudes_comac'])} en {t_phase:.3f}s", "PoleAerien", Qgis.Info)
            # Toutes les phases terminées - lancer la task async
            self._launch_async_task()

    def _launch_async_task(self):
        """Lance la tâche asynchrone après extraction complète."""
        state = self._extraction_state
        t_total = time.perf_counter() - state['t0']
        QgsMessageLog.logMessage(f"[PERF-MAJ] Extraction TOTAL: {t_total:.3f}s", "PoleAerien", Qgis.Info)
        
        raw_data = {
            'poteaux_ft': state['poteaux_ft'],
            'poteaux_bt': state['poteaux_bt'],
            'etudes_cap_ft': state['etudes_cap_ft'],
            'etudes_comac': state['etudes_comac']
        }
        params = {'fichier_excel': state['excel_path']}
        
        # Stocker dans le cache pour les prochains appels
        MajWorkflow._coords_cache = raw_data
        MajWorkflow._cache_key = state.get('cache_key')
        QgsMessageLog.logMessage("[PERF-MAJ] Données mises en cache", "PoleAerien", Qgis.Info)
        
        self._extraction_state = None  # Libérer état temporaire
        
        self.current_task = MajFtBtTask(params, raw_data)
        self.current_task.signals.progress.connect(self.progress_changed)
        self.current_task.signals.message.connect(self.message_received)
        self.current_task.signals.finished.connect(self.analysis_finished)
        self.current_task.signals.error.connect(self.analysis_error)
        
        QgsApplication.taskManager().addTask(self.current_task)

    def apply_updates_ft(self, layer_name, data_list, progress_callback=None):
        """Applique les mises à jour FT finales."""
        return self.maj_logic.miseAjourFinalDesDonneesFT(layer_name, data_list, progress_callback)

    def apply_updates_bt(self, layer_name, data_list, progress_callback=None):
        """Applique les mises à jour BT finales."""
        if hasattr(self.maj_logic, 'miseAjourFinalDesDonneesBT'):
            return self.maj_logic.miseAjourFinalDesDonneesBT(layer_name, data_list, progress_callback)
        else:
            return self.maj_logic.miseAjourFinalDesDonnees(layer_name, data_list)

    def start_updates_incremental(self, layer_name, data_ft, data_bt, progress_callback,
                                  finished_callback, error_callback, should_cancel=None):
        """Lance une MAJ BD incrémentale sur le main thread pour garder l'UI réactive."""
        self._update_state = {
            "layer_name": layer_name,
            "data_ft": data_ft,
            "data_bt": data_bt,
            "progress_callback": progress_callback,
            "finished_callback": finished_callback,
            "error_callback": error_callback,
            "should_cancel": should_cancel,
            "phase": "ft",
            "ft_state": None,
            "bt_state": None,
            "ft_iter": None,
            "bt_iter": None,
            "ft_count": 0,
            "bt_count": 0,
            "ft_total": len(data_ft) if data_ft is not None and not data_ft.empty else 0,
            "bt_total": len(data_bt) if data_bt is not None and not data_bt.empty else 0,
        }
        QTimer.singleShot(0, self._process_update_batch)

    def _rollback_update_state(self, state):
        for key in ("ft_state", "bt_state"):
            sub = state.get(key) if state else None
            layer = sub.get("layer") if sub else None
            if layer and layer.isEditable():
                layer.rollBack()

    def _process_update_batch(self):
        state = self._update_state
        if state is None:
            return

        should_cancel = state.get("should_cancel")
        if should_cancel and should_cancel():
            self._rollback_update_state(state)
            err_cb = state.get("error_callback")
            if err_cb:
                err_cb("Annulation demandee")
            self._update_state = None
            return

        try:
            if state["phase"] == "ft":
                if state["ft_total"] == 0:
                    state["phase"] = "bt"
                else:
                    if state["ft_state"] is None:
                        state["ft_state"] = self.maj_logic._prepare_update_ft(
                            state["layer_name"], state["data_ft"]
                        )
                        state["ft_iter"] = state["data_ft"].iterrows()
                        if state["progress_callback"]:
                            state["progress_callback"](
                                f"MAJ FT: 0/{state['ft_total']} poteaux...", 60
                            )

                    processed = 0
                    while processed < self.UPDATE_BATCH_SIZE:
                        try:
                            gid, row = next(state["ft_iter"])
                        except StopIteration:
                            self.maj_logic._finalize_update_ft(state["ft_state"])
                            state["phase"] = "bt"
                            break

                        state["ft_count"] += 1
                        action = str(row.get("action", "")).upper()
                        self.maj_logic._apply_update_ft_row(state["ft_state"], gid, row)
                        if state["progress_callback"]:
                            pct = 60 + int((state["ft_count"] / state["ft_total"]) * 25)
                            state["progress_callback"](
                                f"MAJ FT: {state['ft_count']}/{state['ft_total']} ({action})...",
                                pct,
                            )
                        processed += 1

                    QTimer.singleShot(0, self._process_update_batch)
                    return

            if state["phase"] == "bt":
                if state["bt_total"] == 0:
                    state["phase"] = "done"
                else:
                    if state["bt_state"] is None:
                        state["bt_state"] = self.maj_logic._prepare_update_bt(
                            state["layer_name"], state["data_bt"]
                        )
                        state["bt_iter"] = state["data_bt"].iterrows()
                        if state["progress_callback"]:
                            state["progress_callback"](
                                f"MAJ BT: 0/{state['bt_total']} poteaux...", 88
                            )

                    processed = 0
                    while processed < self.UPDATE_BATCH_SIZE:
                        try:
                            gid, row = next(state["bt_iter"])
                        except StopIteration:
                            self.maj_logic._finalize_update_bt(state["bt_state"])
                            state["phase"] = "done"
                            break

                        state["bt_count"] += 1
                        self.maj_logic._apply_update_bt_row(state["bt_state"], gid, row)
                        if state["progress_callback"]:
                            pct = 88 + int((state["bt_count"] / state["bt_total"]) * 8) if state["bt_total"] else 88
                            state["progress_callback"](
                                f"MAJ BT: {state['bt_count']}/{state['bt_total']}...",
                                pct,
                            )
                        processed += 1

                    QTimer.singleShot(0, self._process_update_batch)
                    return

            if state["phase"] == "done":
                finish_cb = state.get("finished_callback")
                if finish_cb:
                    finish_cb(state["ft_count"], state["bt_count"])
                self._update_state = None

        except Exception as e:
            self._rollback_update_state(state)
            err_cb = state.get("error_callback")
            if err_cb:
                err_cb(str(e))
            self._update_state = None

    def start_updates(self, layer_name, data_ft, data_bt):
        """
        CRITICAL-001 FIX: Lance la MAJ BD en arrière-plan (non-bloquant).
        
        Remplace les appels synchrones apply_updates_ft/bt qui causaient
        le freeze UI après clic CONFIRMER.
        
        Args:
            layer_name (str): Nom de la couche infra_pt_pot
            data_ft (DataFrame): Données FT à mettre à jour
            data_bt (DataFrame): Données BT à mettre à jour
        """
        self.update_task = MajUpdateTask(layer_name, data_ft, data_bt)
        
        # Connexion des signaux de la task vers le workflow
        self.update_task.signals.progress.connect(self.progress_changed)
        self.update_task.signals.message.connect(self.message_received)
        self.update_task.signals.finished.connect(self.update_finished)
        self.update_task.signals.error.connect(self.update_error)
        
        QgsApplication.taskManager().addTask(self.update_task)

    def start_updates_sql_background(self, layer_name, data_ft, data_bt,
                                      progress_callback, finished_callback, error_callback):
        """
        UI-FREEZE-FIX: Lance la MAJ BD via SQL direct en background.
        
        Cette méthode exécute TOUTES les modifications en arrière-plan
        via connexion PostgreSQL directe, sans bloquer l'UI QGIS.
        L'UI reste 100% réactive pendant toute la durée de la MAJ.
        
        Args:
            layer_name: Nom de la couche infra_pt_pot
            data_ft: DataFrame des MAJ FT
            data_bt: DataFrame des MAJ BT
            progress_callback: callable(pct, msg) pour progression
            finished_callback: callable(result_dict) quand terminé
            error_callback: callable(error_msg) en cas d'erreur
        """
        # Récupérer l'URI de connexion PostgreSQL
        db_uri = get_layer_db_uri(layer_name)
        if db_uri is None:
            error_callback("Couche non PostgreSQL ou invalide")
            return
        
        # Créer et lancer la tâche background
        self._sql_bg_task = MajSqlBackgroundTask(layer_name, data_ft, data_bt, db_uri)
        self._sql_bg_callbacks = {
            'progress': progress_callback,
            'finished': finished_callback,
            'error': error_callback,
            'layer_name': layer_name
        }
        
        # Connecter les signaux
        self._sql_bg_task.signals.progress.connect(self._on_sql_bg_progress)
        self._sql_bg_task.signals.finished.connect(self._on_sql_bg_finished)
        self._sql_bg_task.signals.error.connect(self._on_sql_bg_error)
        
        QgsApplication.taskManager().addTask(self._sql_bg_task)
    
    def _on_sql_bg_progress(self, pct, msg):
        """Relaye la progression de la tâche SQL background."""
        cb = self._sql_bg_callbacks.get('progress')
        if cb:
            cb(msg, pct)
    
    def _on_sql_bg_finished(self, result):
        """Callback quand la MAJ SQL background est terminée."""
        layer_name = self._sql_bg_callbacks.get('layer_name')
        
        # Recharger la couche QGIS pour refléter les modifications
        if layer_name:
            reload_layer(layer_name)
        
        cb = self._sql_bg_callbacks.get('finished')
        if cb:
            cb(result.get('ft_updated', 0), result.get('bt_updated', 0))
    
    def _on_sql_bg_error(self, error_msg):
        """Callback en cas d'erreur de la tâche SQL background."""
        cb = self._sql_bg_callbacks.get('error')
        if cb:
            cb(error_msg)
    
    def cancel_sql_background(self):
        """Annule la tâche SQL background si en cours."""
        if hasattr(self, '_sql_bg_task') and self._sql_bg_task:
            self._sql_bg_task.cancel()
