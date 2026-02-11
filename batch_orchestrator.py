# -*- coding: utf-8 -*-
"""
Batch Orchestrator - Bridges DialogV2 with existing workflows.

Translates detected project paths + QGIS layer selections into
workflow.start_analysis() calls, and routes workflow signals
back to the BatchRunner callback system.
"""

import os
import time

from qgis.PyQt.QtCore import QObject
from qgis.core import QgsMessageLog, Qgis, QgsTask, QgsApplication, QgsProject

from .project_detector import DetectionResult
from .db_layer_loader import DbLayerLoader
from .qgis_utils import detect_etude_field as _detect_etude_field, reset_crs_cache


class _LoadProjectLayersTask(QgsTask):
    """Background task: PG connection + layer creation for project mode.

    Network I/O (connection, metadata fetch) runs in worker thread.
    Project registration (addMapLayer) deferred to main thread callback.
    """

    def __init__(self, sro):
        super().__init__(f"Chargement couches BDD (SRO={sro})")
        self.sro = sro
        self.lyr_pot = None
        self.lyr_cap = None
        self.lyr_com = None
        self.loader = None
        self.error_msg = None
        self.counts = {}

    def run(self):
        self.loader = DbLayerLoader()
        if not self.loader.connect():
            self.error_msg = "Mode Projet: connexion PostgreSQL impossible"
            return False

        if self.isCanceled():
            return False

        self.lyr_pot = self.loader.load_infra_pt_pot(
            self.sro, add_to_project=False
        )
        if not self.lyr_pot or not self.lyr_pot.isValid():
            self.error_msg = "Mode Projet: couche infra_pt_pot non chargee"
            return False
        self.counts['pot'] = self.lyr_pot.featureCount()

        if self.isCanceled():
            return False

        self.lyr_cap = self.loader.load_etude_cap_ft(
            self.sro, add_to_project=False
        )
        if self.lyr_cap and self.lyr_cap.isValid():
            self.counts['cap'] = self.lyr_cap.featureCount()
        else:
            self.lyr_cap = None

        if self.isCanceled():
            return False

        self.lyr_com = self.loader.load_etude_comac(
            self.sro, add_to_project=False
        )
        if self.lyr_com and self.lyr_com.isValid():
            self.counts['com'] = self.lyr_com.featureCount()
        else:
            self.lyr_com = None

        return True

    def finished(self, result):
        pass


_VALIDATION_LABELS = {
    'hors_etude': 'poteaux hors perimetre etude',
    'doublons': 'etudes en doublon',
    'fichiers_doublons': 'fichiers en doublon',
    'erreur_lecture': 'erreurs de lecture',
}


def _format_validation_error(result):
    """Format a validation error result into a readable message."""
    err_type = result.get('error_type', '')
    data = result.get('data', [])
    label = _VALIDATION_LABELS.get(err_type, err_type)
    if isinstance(data, list):
        count = len(data)
        return f"{count} {label}"
    return f"{label}: {data}"


class BatchOrchestrator(QObject):
    """Connects DialogV2 → BatchRunner → Workflows.

    Usage (in PoleAerien.__init__):
        self.batch_orch = BatchOrchestrator(
            dialog_v2, batch_runner,
            maj_workflow, capft_workflow, comac_workflow,
            c6bd_workflow, police_workflow, c6c3a_workflow,
            iface
        )
    """

    def __init__(self, dialog, runner,
                 maj_wf, capft_wf, comac_wf,
                 c6bd_wf, police_wf, c6c3a_wf,
                 iface, parent=None):
        super().__init__(parent)
        self._dlg = dialog
        self._runner = runner
        self._iface = iface

        # Workflows
        self._maj_wf = maj_wf
        self._capft_wf = capft_wf
        self._comac_wf = comac_wf
        self._c6bd_wf = c6bd_wf
        self._police_wf = police_wf
        self._c6c3a_wf = c6c3a_wf

        # Track current batch callback for the active module
        self._current_done_cb = None
        self._module_start_times = {}
        self._batch_total = 0
        self._batch_index = 0
        self._batch_results = {}

        # Shared fddcpi2 cache: avoids duplicate PostgreSQL calls
        # when both COMAC and Police C6 run in the same batch.
        # Key: sro (str), Value: list[CableSegment]
        self._fddcpi_cache = {}

        # Shared SRO + appuis WKB cache: avoids duplicate QGIS extraction
        # when both COMAC and Police C6 need the same data from infra_pt_pot.
        # Format: {'sro': str, 'appuis_wkb': list[dict]}
        self._sro_appuis_cache = {}

        # Project mode: DB-loaded layers (temporary or persistent)
        self._db_loader = None
        self._pm_lyr_pot = None
        self._pm_lyr_cap = None
        self._pm_lyr_com = None

        # Async layer loading state
        self._layer_load_task = None
        self._pending_module_keys = None

        # Register launchers
        self._runner.set_launcher('maj', self._launch_maj)
        self._runner.set_launcher('capft', self._launch_capft)
        self._runner.set_launcher('comac', self._launch_comac)
        self._runner.set_launcher('c6bd', self._launch_c6bd)
        self._runner.set_launcher('police_c6', self._launch_police_c6)
        self._runner.set_launcher('c6c3a', self._launch_c6c3a)

        # Connect runner signals -> dialog
        self._runner.module_started.connect(self._on_module_started)
        self._runner.module_finished.connect(self._on_module_finished)
        self._runner.batch_finished.connect(self._on_batch_finished)
        self._runner.batch_cancelled.connect(self._on_batch_cancelled)
        self._runner.log_message.connect(self._dlg.log_message)

        # Connect workflow progress_changed for intra-module granularity
        for wf in (self._maj_wf, self._capft_wf, self._comac_wf,
                    self._c6bd_wf, self._police_wf, self._c6c3a_wf):
            if hasattr(wf, 'progress_changed'):
                wf.progress_changed.connect(self._on_workflow_progress)

        # Connect dialog signals
        self._dlg.start_requested.connect(self._on_start)
        self._dlg.cancel_requested.connect(self._on_cancel)

    # ------------------------------------------------------------------
    #  Start / Cancel
    # ------------------------------------------------------------------
    def _on_start(self, module_keys):
        self._batch_results = {}
        self._fddcpi_cache = {}
        self._sro_appuis_cache = {}
        reset_crs_cache()
        self._dlg.textBrowser.clear()

        # --- Preflight checks (all-at-once before any module runs) ---
        issues = self._preflight_checks(module_keys)
        if issues:
            self._dlg.log_message(
                f"Verification pre-lancement : {len(issues)} probleme(s) detecte(s)",
                'error'
            )
            for issue in issues:
                self._dlg.log_message(f"  {issue}", 'warning')
            self._dlg.log_message(
                "Corrigez ces problemes puis relancez l'analyse.", 'error'
            )
            return

        self._dlg.set_running(True)

        # Project mode: load layers from DB asynchronously, then start batch
        if self._dlg.is_project_mode:
            self._pending_module_keys = module_keys
            self._start_async_layer_loading()
            return

        self._runner.start(module_keys)

    def _preflight_checks(self, module_keys):
        """Validate ALL prerequisites upfront. Returns list of actionable messages (empty = OK)."""
        issues = []
        det = self._detection()
        is_pm = self._dlg.is_project_mode

        # --- Export directory ---
        export_dir = self._export_dir()
        if export_dir and not os.path.isdir(export_dir):
            issues.append(
                f"Dossier d'export introuvable : {export_dir}"
            )
        elif export_dir:
            try:
                test_path = os.path.join(export_dir, '.poleaerien_test')
                with open(test_path, 'w') as _f:
                    pass
                os.remove(test_path)
            except OSError:
                issues.append(
                    f"Dossier d'export non inscriptible : {export_dir}"
                )

        # --- Layer checks (QGIS mode only) ---
        if not is_pm:
            lyr_pot = self._lyr_pot()
            lyr_cap = self._lyr_capft()
            lyr_com = self._lyr_comac()

            needs_pot = set(module_keys) & {'maj', 'capft', 'comac', 'c6bd', 'police_c6', 'c6c3a'}
            needs_cap = set(module_keys) & {'maj', 'capft', 'c6bd', 'c6c3a'}
            needs_com = set(module_keys) & {'maj', 'comac'}

            if needs_pot and (not lyr_pot or not lyr_pot.isValid()):
                issues.append(
                    "Couche infra_pt_pot manquante ou invalide. "
                    "Chargez la couche poteaux dans le projet QGIS "
                    "et selectionnez-la dans la liste deroulante."
                )
            elif needs_pot and lyr_pot and lyr_pot.featureCount() == 0:
                issues.append(
                    f"Couche {lyr_pot.name()} est vide (0 entites). "
                    "Verifiez le filtre applique sur la couche."
                )

            if needs_cap and (not lyr_cap or not lyr_cap.isValid()):
                issues.append(
                    "Couche etude_cap_ft manquante ou invalide. "
                    "Chargez la couche de zonage CAP FT dans le projet QGIS."
                )

            if needs_com and (not lyr_com or not lyr_com.isValid()):
                issues.append(
                    "Couche etude_comac manquante ou invalide. "
                    "Chargez la couche de zonage COMAC dans le projet QGIS."
                )

        # --- Project mode: SRO ---
        if is_pm and not self._dlg.sro:
            issues.append(
                "Mode Projet: SRO non derivable du nom de dossier. "
                "Format attendu: XXXXX-YYY-ZZZ-NNNNN (ex: 63041-B1I-PMZ-00003)"
            )

        # --- Per-module resource checks ---
        if 'maj' in module_keys and not det.has_ftbt:
            issues.append(
                "Module MAJ: fichier FT-BT KO non detecte. "
                "Placez un fichier Excel contenant 'FT-BT KO' dans le nom a la racine du projet."
            )

        if 'capft' in module_keys and not det.has_capft:
            issues.append(
                "Module CAP_FT: dossier CAP FT non detecte. "
                "Noms acceptes: CAP FT, CAPFT, ETUDE CAP FT, KPFT."
            )

        if 'comac' in module_keys and not det.has_comac:
            issues.append(
                "Module COMAC: dossier COMAC non detecte. "
                "Noms acceptes: COMAC, Export COMAC, ETUDE COMAC."
            )

        c6_modules = set(module_keys) & {'c6bd', 'police_c6', 'c6c3a'}
        if c6_modules and not det.has_c6:
            issues.append(
                f"Module(s) {', '.join(sorted(c6_modules))}: dossier C6 (= CAP FT) non detecte. "
                "Les annexes C6 sont les fichiers d'etude dans le dossier CAP FT."
            )

        if 'c6c3a' in module_keys:
            if not det.has_c7 and not det.has_c3a:
                issues.append(
                    "Module C6-C3A-C7-BD: ni fichier C7 ni C3A detecte. "
                    "Au moins un des deux est necessaire."
                )

        return issues

    def _on_cancel(self):
        # Cancel layer loading task if in progress
        if self._layer_load_task:
            self._layer_load_task.cancel()
        self._runner.cancel()
        # Cancel active workflow
        for wf in (self._maj_wf, self._capft_wf, self._comac_wf,
                    self._c6bd_wf, self._police_wf, self._c6c3a_wf):
            if hasattr(wf, 'cancel'):
                try:
                    wf.cancel()
                except Exception:
                    pass

    # ------------------------------------------------------------------
    #  Runner signal handlers
    # ------------------------------------------------------------------
    def _on_module_started(self, key, idx, total):
        self._batch_index = idx
        self._batch_total = total
        self._module_start_times[key] = time.time()
        # Push progress to start of this module's sub-range
        if total > 0:
            pct = int((idx / total) * 100)
            self._dlg.set_progress(max(pct, 2))

    def _on_module_finished(self, key, success, msg):
        # Push progress to end of this module's sub-range
        if self._batch_total > 0:
            pct = int(((self._batch_index + 1) / self._batch_total) * 100)
            self._dlg.set_progress(pct)

    def _on_workflow_progress(self, module_pct):
        """Map per-module 0-100 progress into the batch sub-range."""
        total = self._batch_total
        idx = self._batch_index
        if total <= 0:
            return
        range_start = (idx / total) * 100
        range_end = ((idx + 1) / total) * 100
        mapped = range_start + (module_pct / 100.0) * (range_end - range_start)
        self._dlg.set_progress(int(mapped))

    def _on_batch_finished(self, results):
        # --- Post-batch recap with metrics ---
        self._log_batch_recap(results)

        # Generate unified Excel report
        if self._batch_results:
            try:
                from .unified_report import generate_unified_report
                export_dir = self._export_dir()
                filepath = generate_unified_report(self._batch_results, export_dir)
                self._dlg.log_result_link(filepath)
            except Exception as exc:
                self._dlg.log_message(f"Erreur rapport: {exc}", 'error')

        # Project mode cleanup: remove temp layers if user didn't ask to keep them
        self._cleanup_project_mode_layers()
        self._dlg.reset_after_batch()

    def _log_batch_recap(self, runner_results):
        """Log a structured recap with metrics from each completed module."""
        if not self._batch_results:
            return

        self._dlg.log_message("--- RECAPITULATIF ---", 'info')

        elapsed = {}
        for key, t0 in self._module_start_times.items():
            elapsed[key] = time.time() - t0

        for key, result in self._batch_results.items():
            name = {
                'maj': 'MAJ BD', 'capft': 'CAP_FT', 'comac': 'COMAC',
                'c6bd': 'C6 vs BD', 'police_c6': 'POLICE C6', 'c6c3a': 'C6-C3A-C7-BD',
            }.get(key, key)

            dur = elapsed.get(key)
            dur_str = f" ({dur:.1f}s)" if dur else ""

            metrics = self._extract_metrics(key, result)
            if metrics:
                self._dlg.log_message(f"  {name}{dur_str}: {metrics}", 'info')
            else:
                run_res = runner_results.get(key, {})
                status = 'OK' if run_res.get('success') else 'ERREUR'
                self._dlg.log_message(f"  {name}{dur_str}: {status}", 'info')

    def _extract_metrics(self, key, result):
        """Extract human-readable metrics from a module result dict."""
        parts = []

        if key == 'maj':
            liste_ft = result.get('liste_ft', [])
            liste_bt = result.get('liste_bt', [])
            ft_count = liste_ft[2] if len(liste_ft) > 2 else 0
            bt_count = liste_bt[2] if len(liste_bt) > 2 else 0
            if ft_count or bt_count:
                parts.append(f"{ft_count} FT, {bt_count} BT a mettre a jour")

        elif key == 'capft':
            resultats = result.get('resultats', {})
            if resultats:
                total = sum(len(v) for v in resultats.values()) if isinstance(resultats, dict) else 0
                parts.append(f"{len(resultats)} etudes, {total} poteaux analyses")

        elif key == 'comac':
            resultats = result.get('resultats', {})
            verif = result.get('verif_cables', [])
            if resultats:
                parts.append(f"{len(resultats)} etudes")
            if verif:
                nb_ecart = sum(1 for v in verif if v.get('statut') == 'ECART')
                nb_absent = sum(1 for v in verif if v.get('statut') == 'ABSENT_BDD')
                if nb_ecart or nb_absent:
                    parts.append(f"{nb_ecart} ecarts cables, {nb_absent} absents BDD")
                else:
                    parts.append("cables OK")

        elif key == 'police_c6':
            stats = result.get('stats', [])
            if stats:
                total_ecart = sum(s.get('nb_ecart', 0) for s in stats)
                total_absent = sum(s.get('nb_absent', 0) for s in stats)
                total_boitier = sum(s.get('nb_boitier_err', 0) for s in stats)
                parts.append(f"{len(stats)} etudes")
                anomalies = total_ecart + total_absent + total_boitier
                if anomalies:
                    parts.append(f"{anomalies} anomalie(s)")
                else:
                    parts.append("aucune anomalie")

        elif key == 'c6bd':
            if result.get('success'):
                parts.append("OK")

        elif key == 'c6c3a':
            if result.get('success'):
                parts.append("OK")

        return ", ".join(parts) if parts else ""

    def _on_batch_cancelled(self):
        self._cleanup_project_mode_layers()
        self._dlg.reset_after_batch()
        self._dlg.log_message("Batch annulé.", 'warning')

    # ------------------------------------------------------------------
    #  Helper: get common params from dialog
    # ------------------------------------------------------------------
    def _detection(self) -> DetectionResult:
        return self._dlg.detection

    def _export_dir(self) -> str:
        return self._dlg.get_export_dir()

    def _lyr_pot(self):
        if self._pm_lyr_pot:
            return self._pm_lyr_pot
        return self._dlg.comboInfraPtPot.currentLayer()

    def _lyr_capft(self):
        if self._pm_lyr_cap:
            return self._pm_lyr_cap
        return self._dlg.comboEtudeCapFt.currentLayer()

    def _lyr_comac(self):
        if self._pm_lyr_com:
            return self._pm_lyr_com
        return self._dlg.comboEtudeComac.currentLayer()

    # ------------------------------------------------------------------
    #  Project Mode: async DB layer loading
    # ------------------------------------------------------------------
    def _start_async_layer_loading(self):
        """Launch background task to load PostgreSQL layers for project mode.

        Layer creation (network I/O) runs in a worker thread.
        On completion, layers are registered in QgsProject on the main thread,
        then the batch starts.
        """
        sro = self._dlg.sro
        if not sro:
            self._dlg.log_message("Mode Projet: SRO non disponible", 'error')
            self._pending_module_keys = None
            self._dlg.reset_after_batch()
            return

        self._dlg.log_message(
            f"Mode Projet: chargement depuis BDD (SRO={sro})", 'info'
        )

        self._layer_load_task = _LoadProjectLayersTask(sro)
        self._layer_load_task.taskCompleted.connect(self._on_layers_loaded)
        self._layer_load_task.taskTerminated.connect(self._on_layers_load_failed)
        QgsApplication.taskManager().addTask(self._layer_load_task)

    def _on_layers_loaded(self):
        """Main thread callback: register layers in QgsProject, start batch."""
        task = self._layer_load_task
        self._layer_load_task = None

        if not task:
            return

        self._db_loader = task.loader

        # Always add layers to QGIS project so that get_layer_safe() and
        # other name-based lookups work inside workflows.
        # Cleanup after batch removes them if user didn't ask to keep.

        # infra_pt_pot (mandatory - guaranteed valid by task.run())
        self._pm_lyr_pot = task.lyr_pot
        QgsProject.instance().addMapLayer(self._pm_lyr_pot)
        self._db_loader._loaded_layers.append(self._pm_lyr_pot.id())
        self._dlg.log_message(
            f"  infra_pt_pot: {task.counts.get('pot', '?')} poteaux", 'success'
        )

        # etude_cap_ft (optional)
        if task.lyr_cap:
            self._pm_lyr_cap = task.lyr_cap
            QgsProject.instance().addMapLayer(self._pm_lyr_cap)
            self._db_loader._loaded_layers.append(self._pm_lyr_cap.id())
            self._dlg.log_message(
                f"  etude_cap_ft: {task.counts.get('cap', '?')} zones",
                'success'
            )
        else:
            self._dlg.log_message(
                "  etude_cap_ft: non disponible (modules CAP FT/MAJ limites)",
                'warning'
            )
            self._pm_lyr_cap = None

        # etude_comac (optional)
        if task.lyr_com:
            self._pm_lyr_com = task.lyr_com
            QgsProject.instance().addMapLayer(self._pm_lyr_com)
            self._db_loader._loaded_layers.append(self._pm_lyr_com.id())
            self._dlg.log_message(
                f"  etude_comac: {task.counts.get('com', '?')} zones",
                'success'
            )
        else:
            self._dlg.log_message(
                "  etude_comac: non disponible (module COMAC/MAJ limites)",
                'warning'
            )
            self._pm_lyr_com = None

        # Start the batch now that layers are ready
        keys = self._pending_module_keys
        self._pending_module_keys = None
        if keys:
            self._runner.start(keys)

    def _on_layers_load_failed(self):
        """Main thread callback: layer loading failed or was cancelled."""
        task = self._layer_load_task
        self._layer_load_task = None
        self._pending_module_keys = None

        msg = (task.error_msg if task and task.error_msg
               else "Mode Projet: erreur chargement couches")
        self._dlg.log_message(msg, 'error')
        self._dlg.reset_after_batch()

    def _cleanup_project_mode_layers(self):
        """Remove temporary DB layers if user did not ask to keep them."""
        if not self._db_loader:
            return

        if not self._dlg.load_layers_in_qgis:
            self._db_loader.cleanup_layers()
            QgsMessageLog.logMessage(
                "Mode Projet: couches temporaires supprimees",
                "PoleAerien", Qgis.Info
            )

        self._pm_lyr_pot = None
        self._pm_lyr_cap = None
        self._pm_lyr_com = None
        self._db_loader = None

    def _auto_field(self, layer, patterns=None):
        """Auto-detect study field name from layer. Delegue a qgis_utils."""
        result = _detect_etude_field(layer, context="BatchOrchestrator")
        if result:
            return result
        # Fallback: first text field
        if layer and layer.isValid():
            for field in layer.fields():
                if field.typeName().lower() in ('string', 'text', 'varchar'):
                    return field.name()
        return ''

    # ------------------------------------------------------------------
    #  Module launchers
    # ------------------------------------------------------------------
    def _launch_maj(self, on_done):
        """Launch MAJ FT/BT module."""
        self._current_done_cb = on_done
        det = self._detection()

        lyr_pot = self._lyr_pot()
        lyr_cap = self._lyr_capft()
        lyr_com = self._lyr_comac()

        if not lyr_pot or not lyr_cap or not lyr_com:
            on_done(False, "Couches QGIS non sélectionnées")
            return

        if not det.has_ftbt:
            on_done(False, "Fichier FT-BT KO non détecté")
            return

        # Disconnect previous, connect for this batch
        self._safe_connect_once(
            self._maj_wf.analysis_finished, self._on_maj_done
        )
        self._safe_connect_once(
            self._maj_wf.error_occurred, self._on_maj_error
        )
        self._safe_connect_once(
            self._maj_wf.message_received, self._relay_message
        )

        col_cap = self._auto_field(lyr_cap)
        col_com = self._auto_field(lyr_com)
        self._maj_wf.start_analysis(
            lyr_pot, lyr_cap, lyr_com, det.ftbt_excel,
            col_cap=col_cap, col_com=col_com
        )

    def _on_maj_done(self, result):
        import pandas as pd

        cb = self._current_done_cb
        if not cb:
            return

        liste_ft = result.get('liste_ft', [0, None, 0, None])
        liste_bt = result.get('liste_bt', [0, None, 0, None])
        ft_count = liste_ft[2] if len(liste_ft) > 2 else 0
        bt_count = liste_bt[2] if len(liste_bt) > 2 else 0
        self._batch_results['maj'] = result

        # Extract DataFrames of matched rows (index=gid)
        data_ft = liste_ft[3] if len(liste_ft) > 3 and isinstance(liste_ft[3], pd.DataFrame) else pd.DataFrame()
        data_bt = liste_bt[3] if len(liste_bt) > 3 and isinstance(liste_bt[3], pd.DataFrame) else pd.DataFrame()

        if data_ft.empty and data_bt.empty:
            cb(True, f"Analyse: {ft_count} FT, {bt_count} BT (aucune MAJ)")
            return

        # Get layer for SQL update
        lyr = self._lyr_pot()
        if not lyr:
            cb(True, f"Analyse: {ft_count} FT, {bt_count} BT (couche introuvable)")
            return

        layer_name = lyr.name()
        self._dlg.log_message("Mise a jour BD en cours...", "info")

        # Save state for async callbacks
        self._maj_pending_cb = cb
        self._maj_ft_count = ft_count
        self._maj_bt_count = bt_count

        def _on_sql_progress(msg, pct):
            self._dlg.log_message(msg, "info")

        def _on_sql_finished(ft_updated, bt_updated):
            self._dlg.log_message(
                f"MAJ BD terminee: {ft_updated} FT, {bt_updated} BT mis a jour",
                "success"
            )
            done_cb = self._maj_pending_cb
            if done_cb:
                done_cb(True,
                        f"Analyse: {self._maj_ft_count} FT, {self._maj_bt_count} BT"
                        f" | MAJ: {ft_updated} FT, {bt_updated} BT")

        def _on_sql_error(err):
            self._dlg.log_message(f"Erreur MAJ SQL: {err}", "error")
            done_cb = self._maj_pending_cb
            if done_cb:
                done_cb(False, f"Erreur MAJ: {err}")

        self._maj_wf.start_updates_sql_background(
            layer_name, data_ft, data_bt,
            _on_sql_progress, _on_sql_finished, _on_sql_error
        )

    def _on_maj_error(self, err):
        cb = self._current_done_cb
        if cb:
            cb(False, str(err))

    # ------------------------------------------------------------------
    def _launch_capft(self, on_done):
        """Launch CAP_FT module."""
        self._current_done_cb = on_done
        det = self._detection()

        lyr_pot = self._lyr_pot()
        lyr_cap = self._lyr_capft()
        col_cap = self._auto_field(lyr_cap)
        export_dir = self._export_dir()

        if not lyr_pot or not lyr_cap:
            on_done(False, "Couches QGIS non sélectionnées")
            return
        if not det.has_capft:
            on_done(False, "Dossier CAP FT non détecté")
            return

        self._safe_connect_once(
            self._capft_wf.export_finished, self._on_capft_done
        )
        self._safe_connect_once(
            self._capft_wf.analysis_finished, self._on_capft_analysis
        )
        self._safe_connect_once(
            self._capft_wf.error_occurred, self._on_capft_error
        )
        self._safe_connect_once(
            self._capft_wf.message_received, self._relay_message
        )

        self._capft_wf.start_analysis(
            lyr_pot, lyr_cap, col_cap,
            det.capft_dir, export_dir
        )

    def _on_capft_analysis(self, result):
        err_type = result.get('error_type', '')
        if err_type:
            cb = self._current_done_cb
            if cb:
                cb(False, _format_validation_error(result))
            return
        self._batch_results['capft'] = result
        cb = self._current_done_cb
        if cb:
            cb(True, 'OK')

    def _on_capft_done(self, result):
        pass

    def _on_capft_error(self, err):
        cb = self._current_done_cb
        if cb:
            cb(False, str(err))

    # ------------------------------------------------------------------
    def _launch_comac(self, on_done):
        """Launch COMAC module."""
        self._current_done_cb = on_done
        det = self._detection()

        lyr_pot = self._lyr_pot()
        lyr_comac = self._lyr_comac()
        col_comac = self._auto_field(lyr_comac)
        export_dir = self._export_dir()

        if not lyr_pot or not lyr_comac:
            on_done(False, "Couches QGIS non sélectionnées")
            return
        if not det.has_comac:
            on_done(False, "Dossier COMAC non détecté")
            return

        self._safe_connect_once(
            self._comac_wf.export_finished, self._on_comac_done
        )
        self._safe_connect_once(
            self._comac_wf.analysis_finished, self._on_comac_analysis
        )
        self._safe_connect_once(
            self._comac_wf.error_occurred, self._on_comac_error
        )
        self._safe_connect_once(
            self._comac_wf.message_received, self._relay_message
        )

        # Pass cached data if available (e.g. Police C6 ran first)
        fddcpi = self._fddcpi_cache.get('cables') if self._fddcpi_cache else None
        sro_appuis = self._sro_appuis_cache if self._sro_appuis_cache else None
        self._comac_wf.start_analysis(
            lyr_pot, lyr_comac, col_comac,
            det.comac_dir, export_dir,
            fddcpi_cache=fddcpi,
            sro_appuis_cache=sro_appuis
        )

    def _on_comac_analysis(self, result):
        err_type = result.get('error_type', '')
        if err_type:
            cb = self._current_done_cb
            if cb:
                cb(False, _format_validation_error(result))
            return
        # Cache fddcpi2 data for other modules in this batch
        cables_all = result.get('fddcpi_cables_all')
        sro = result.get('fddcpi_sro')
        if cables_all is not None and sro:
            self._fddcpi_cache = {'sro': sro, 'cables': cables_all}
        # Cache SRO + appuis WKB for Police C6
        appuis_wkb = result.get('appuis_wkb')
        if sro and appuis_wkb is not None and not self._sro_appuis_cache:
            self._sro_appuis_cache = {'sro': sro, 'appuis_wkb': appuis_wkb}
        self._batch_results['comac'] = result
        cb = self._current_done_cb
        if cb:
            cb(True, 'OK')

    def _on_comac_done(self, result):
        pass

    def _on_comac_error(self, err):
        cb = self._current_done_cb
        if cb:
            cb(False, str(err))

    # ------------------------------------------------------------------
    def _launch_c6bd(self, on_done):
        """Launch C6 vs BD module."""
        self._current_done_cb = on_done
        det = self._detection()

        lyr_pot = self._lyr_pot()
        lyr_cap = self._lyr_capft()
        export_dir = self._export_dir()

        if not lyr_pot or not lyr_cap:
            on_done(False, "Couches QGIS non sélectionnées")
            return
        if not det.has_c6:
            on_done(False, "Dossier C6 non détecté")
            return

        c6_path = det.c6_dir

        self._safe_connect_once(
            self._c6bd_wf.export_finished, self._on_c6bd_done
        )
        self._safe_connect_once(
            self._c6bd_wf.analysis_finished, self._on_c6bd_analysis
        )
        self._safe_connect_once(
            self._c6bd_wf.error_occurred, self._on_c6bd_error
        )
        self._safe_connect_once(
            self._c6bd_wf.message_received, self._relay_message
        )

        # col_cap=None → auto-detect in workflow
        self._c6bd_wf.start_analysis(
            lyr_pot, lyr_cap, None,
            c6_path, export_dir
        )

    def _on_c6bd_analysis(self, result):
        self._batch_results['c6bd'] = result
        cb = self._current_done_cb
        if cb:
            cb(True, 'OK')

    def _on_c6bd_done(self, result):
        pass

    def _on_c6bd_error(self, err):
        cb = self._current_done_cb
        if cb:
            cb(False, str(err))

    # ------------------------------------------------------------------
    def _launch_police_c6(self, on_done):
        """Launch Police C6 module."""
        self._current_done_cb = on_done
        det = self._detection()

        if not det.has_c6:
            on_done(False, "Dossier C6 non détecté")
            return

        # In project mode, skip check_layers_exist (layer comes from DB loader)
        if not self._dlg.is_project_mode:
            liste_absent = self._police_wf.check_layers_exist()
            if liste_absent:
                on_done(False, f"Couches manquantes: {', '.join(liste_absent)}")
                return

        self._safe_connect_once(
            self._police_wf.analysis_finished, self._on_police_done
        )
        self._safe_connect_once(
            self._police_wf.error_occurred, self._on_police_error
        )
        self._safe_connect_once(
            self._police_wf.message_received, self._relay_message
        )

        params = {}

        # Project mode: inject layer reference and SRO directly
        if self._pm_lyr_pot:
            params['layer_appuis'] = self._pm_lyr_pot
            params['sro'] = self._dlg.sro

        # Pass study layer info so auto-browse can get study names
        lyr_cap = self._lyr_capft()
        if lyr_cap:
            params['table_etude'] = lyr_cap.name()
            col = self._auto_field(lyr_cap)
            if col:
                params['colonne_etude'] = col

        # Inject cached data if available (e.g. COMAC ran first)
        if self._fddcpi_cache:
            params['fddcpi_cables_cache'] = self._fddcpi_cache.get('cables')
        if self._sro_appuis_cache:
            params['sro_appuis_cache'] = self._sro_appuis_cache

        c6_path = det.c6_dir
        self._police_wf.reset_logic()

        if os.path.isdir(c6_path):
            params['repertoire_c6'] = c6_path
            self._police_wf.run_analysis_auto_browse(params)
        else:
            params['fname'] = c6_path
            self._police_wf.run_analysis(params)

    def _on_police_done(self, result):
        # Cache fddcpi2 data for other modules in this batch
        cables_all = result.get('fddcpi_cables_all')
        sro = result.get('fddcpi_sro')
        if cables_all is not None and sro:
            self._fddcpi_cache = {'sro': sro, 'cables': cables_all}
        # Cache SRO + appuis WKB for COMAC
        appuis_wkb = result.get('appuis_wkb')
        if sro and appuis_wkb is not None and not self._sro_appuis_cache:
            self._sro_appuis_cache = {'sro': sro, 'appuis_wkb': appuis_wkb}
        self._batch_results['police_c6'] = result
        cb = self._current_done_cb
        if cb:
            cb(True, 'OK')

    def _on_police_error(self, err):
        cb = self._current_done_cb
        if cb:
            cb(False, str(err))

    # ------------------------------------------------------------------
    def _launch_c6c3a(self, on_done):
        """Launch C6-C3A-C7-BD module."""
        self._current_done_cb = on_done
        det = self._detection()

        if not det.has_c6:
            on_done(False, "Dossier C6 non détecté")
            return

        # Build params from detected files
        c6_file = ''
        if os.path.isdir(det.c6_dir):
            # Find first xlsx in c6 dir
            for f in sorted(os.listdir(det.c6_dir)):
                if f.lower().endswith('.xlsx'):
                    c6_file = os.path.join(det.c6_dir, f)
                    break
        else:
            c6_file = det.c6_dir

        if not c6_file or not os.path.isfile(c6_file):
            on_done(False, "Fichier C6 introuvable")
            return

        params = {
            'fichier_c6': c6_file,
            'fichier_c7': det.c7_file or '',
            'chemin_export': self._export_dir(),
            'mode_c3a': 'QGIS',
        }

        # Layer names from QGIS project
        lyr_pot = self._lyr_pot()
        lyr_cap = self._lyr_capft()
        if lyr_pot:
            params['table_infra'] = lyr_pot.name()
        if lyr_cap:
            params['table_decoupage'] = lyr_cap.name()
            params['champs_dcp'] = self._auto_field(lyr_cap)

        if det.has_c3a:
            params['fichier_c3a'] = det.c3a_file
            params['mode_c3a'] = 'EXCEL'

        self._safe_connect_once(
            self._c6c3a_wf.export_finished, self._on_c6c3a_done
        )
        self._safe_connect_once(
            self._c6c3a_wf.analysis_finished, self._on_c6c3a_analysis
        )
        self._safe_connect_once(
            self._c6c3a_wf.error_occurred, self._on_c6c3a_error
        )
        self._safe_connect_once(
            self._c6c3a_wf.message_received, self._relay_message
        )

        self._c6c3a_wf.start_analysis(params)

    def _on_c6c3a_analysis(self, result):
        if result.get('success'):
            self._batch_results['c6c3a'] = result
            cb = self._current_done_cb
            if cb:
                cb(True, 'OK')
        else:
            cb = self._current_done_cb
            if cb:
                cb(False, 'Erreur analyse')

    def _on_c6c3a_done(self, result):
        pass

    def _on_c6c3a_error(self, err):
        cb = self._current_done_cb
        if cb:
            cb(False, str(err))

    # ------------------------------------------------------------------
    #  Helpers
    # ------------------------------------------------------------------
    def _relay_message(self, msg, color):
        """Relay workflow messages to dialog log."""
        level_map = {
            'red': 'error',
            'orange': 'warning',
            'green': 'success',
            'grey': 'info',
            'gray': 'info',
            'blue': 'info',
        }
        level = level_map.get(color, 'info')
        self._dlg.log_message(msg, level)

    def _safe_connect_once(self, signal, slot):
        """Connect signal to slot, disconnecting previous connection first."""
        try:
            signal.disconnect(slot)
        except (TypeError, RuntimeError):
            pass
        signal.connect(slot)
