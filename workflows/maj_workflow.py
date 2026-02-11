# -*- coding: utf-8 -*-
"""
Workflow pour la mise à jour FT/BT.
Orchestre l'extraction des données et l'exécution asynchrone.
"""

from qgis.PyQt.QtCore import QObject, pyqtSignal, QTimer
from qgis.core import QgsApplication, QgsFeatureRequest, QgsExpression, NULL
from ..Maj_Ft_Bt import MajFtBt, MajFtBtTask
from ..maj_sql_background import MajSqlBackgroundTask, get_layer_db_uri, reload_layer
from ..qgis_utils import validate_same_crs, detect_etude_field as _detect_etude_field

class MajWorkflow(QObject):
    """
    Contrôleur pour le flux de Mise à Jour FT/BT.
    Gère:
    1. L'extraction des données QGIS (Main Thread, non-bloquant via QTimer)
    2. Le lancement de la tâche asynchrone d'analyse (Worker Thread)
    3. La MAJ BD via SQL direct en background
    4. La communication des résultats via signaux
    """
    
    # Signaux relayés vers l'UI
    progress_changed = pyqtSignal(int)
    message_received = pyqtSignal(str, str)  # message, couleur
    analysis_finished = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)
    
    BATCH_SIZE = 50  # Features par lot d'extraction
    
    # Cache statique partagé entre instances
    _coords_cache = None
    _cache_key = None
    
    def __init__(self):
        super().__init__()
        self.maj_logic = MajFtBt()
        self.current_task = None
        self._sql_bg_task = None
        self._sql_bg_callbacks = {}
        
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

    @staticmethod
    def _detect_etude_field(layer):
        """Auto-detect study field name from a layer. Delegue a qgis_utils."""
        return _detect_etude_field(layer, context="MAJ")

    def start_analysis(self, lyr_pot, lyr_cap, lyr_com, excel_path,
                       col_cap=None, col_com=None):
        """Lance l'analyse avec extraction non-bloquante (avec cache).
        
        Args:
            col_cap: Nom du champ etude dans la couche CAP FT (auto-detect si None)
            col_com: Nom du champ etude dans la couche COMAC (auto-detect si None)
        """
        if not lyr_pot or not lyr_cap or not lyr_com:
            self.error_occurred.emit("Couches invalides ou manquantes")
            return

        try:
            self.maj_logic.valider_nom_fichier(excel_path)
        except ValueError as e:
            self.error_occurred.emit(str(e))
            return

        try:
            validate_same_crs(lyr_pot, lyr_cap, "MAJ_FT_BT")
            validate_same_crs(lyr_pot, lyr_com, "MAJ_FT_BT")
        except ValueError as e:
            self.error_occurred.emit(str(e))
            return

        # Vérifier le cache
        cache_key = self._compute_cache_key(lyr_pot, lyr_cap, lyr_com)
        if MajWorkflow._cache_key == cache_key and MajWorkflow._coords_cache is not None:
            # CACHE HIT - utiliser les données en cache (instantané!)
            self.message_received.emit("Données en cache...", "green")
            self.progress_changed.emit(15)
            
            # Lancer directement la tâche async avec les données en cache
            params = {'fichier_excel': excel_path}
            self.current_task = MajFtBtTask(params, MajWorkflow._coords_cache)
            self.current_task.signals.progress.connect(self.progress_changed)
            self.current_task.signals.message.connect(self.message_received)
            self.current_task.signals.finished.connect(self.analysis_finished)
            self.current_task.signals.error.connect(self.error_occurred)
            QgsApplication.taskManager().addTask(self.current_task)
            return

        # CACHE MISS - faire l'extraction complète
        
        # Auto-detect study field if not provided
        if not col_cap:
            col_cap = self._detect_etude_field(lyr_cap)
        if not col_com:
            col_com = self._detect_etude_field(lyr_com)
        
        if not col_cap:
            self.error_occurred.emit(
                f"Champ etude non detecte dans {lyr_cap.name()}. "
                f"Champs: {[f.name() for f in lyr_cap.fields()]}")
            return
        if not col_com:
            self.error_occurred.emit(
                f"Champ etude non detecte dans {lyr_com.name()}. "
                f"Champs: {[f.name() for f in lyr_com.fields()]}")
            return
        
        # Initialiser l'état d'extraction
        self._extraction_state = {
            'excel_path': excel_path,
            'cache_key': cache_key,
            'lyr_pot': lyr_pot,
            'lyr_cap': lyr_cap,
            'lyr_com': lyr_com,
            'col_cap': col_cap,
            'col_com': col_com,
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
                        raw_etude = feat[state['col_cap']]
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
                        raw_etude = feat[state['col_com']]
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
            self.error_occurred.emit(f"Erreur extraction: {e}")
            self._extraction_state = None

    def _finish_phase(self):
        """Termine une phase et passe à la suivante."""
        state = self._extraction_state
        phase = state['phase']
        
        if phase == 'pot_ft':
            state['phase'] = 'pot_bt'
            state['iterator'] = state['lyr_pot'].getFeatures(
                QgsFeatureRequest(QgsExpression("inf_type LIKE 'POT-BT'"))
            )
            self.message_received.emit("Extraction POT-BT...", "grey")
            self.progress_changed.emit(8)
            QTimer.singleShot(0, self._process_extraction_batch)
            
        elif phase == 'pot_bt':
            state['phase'] = 'cap_ft'
            state['iterator'] = state['lyr_cap'].getFeatures()
            self.message_received.emit("Extraction CAP_FT...", "grey")
            self.progress_changed.emit(12)
            QTimer.singleShot(0, self._process_extraction_batch)
            
        elif phase == 'cap_ft':
            state['phase'] = 'comac'
            state['iterator'] = state['lyr_com'].getFeatures()
            self.message_received.emit("Extraction COMAC...", "grey")
            self.progress_changed.emit(16)
            QTimer.singleShot(0, self._process_extraction_batch)
            
        elif phase == 'comac':
            # Toutes les phases terminées - lancer la task async
            self._launch_async_task()

    def _launch_async_task(self):
        """Lance la tâche asynchrone après extraction complète."""
        state = self._extraction_state
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
        
        self._extraction_state = None  # Libérer état temporaire
        
        self.current_task = MajFtBtTask(params, raw_data)
        self.current_task.signals.progress.connect(self.progress_changed)
        self.current_task.signals.message.connect(self.message_received)
        self.current_task.signals.finished.connect(self.analysis_finished)
        self.current_task.signals.error.connect(self.error_occurred)
        
        QgsApplication.taskManager().addTask(self.current_task)

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
        """Relaye la progression de la tâche SQL background.
        
        Note: Le signal émet (int pct, str msg). Le callback attend (str msg, int pct).
        """
        cb = self._sql_bg_callbacks.get('progress')
        if cb:
            cb(msg, pct)  # Inversion intentionnelle: signal=(pct,msg) → callback=(msg,pct)
    
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
