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
from qgis.core import QgsMessageLog, Qgis

from .project_detector import DetectionResult
from .db_layer_loader import DbLayerLoader


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

        # Project mode: DB-loaded layers (temporary or persistent)
        self._db_loader = None
        self._pm_lyr_pot = None
        self._pm_lyr_cap = None
        self._pm_lyr_com = None

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
        self._dlg.set_running(True)
        self._dlg.textBrowser.clear()

        # Project mode: load layers from DB before starting batch
        if self._dlg.is_project_mode:
            if not self._load_project_mode_layers():
                self._dlg.reset_after_batch()
                return

        self._runner.start(module_keys)

    def _on_cancel(self):
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
        # Generate unified Excel report
        if self._batch_results:
            try:
                from .unified_report import generate_unified_report
                export_dir = self._export_dir()
                filepath = generate_unified_report(self._batch_results, export_dir)
                fname = os.path.basename(filepath)
                self._dlg.log_message(f"Rapport: {fname}", 'success')
                os.startfile(filepath)
            except Exception as exc:
                self._dlg.log_message(f"Erreur rapport: {exc}", 'error')

        # Project mode cleanup: remove temp layers if user didn't ask to keep them
        self._cleanup_project_mode_layers()
        self._dlg.reset_after_batch()

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
    #  Project Mode: DB layer loading
    # ------------------------------------------------------------------
    def _load_project_mode_layers(self) -> bool:
        """Load filtered layers from PostgreSQL for project mode.

        Returns:
            True if infra_pt_pot loaded successfully (mandatory).
        """
        sro = self._dlg.sro
        if not sro:
            self._dlg.log_message("Mode Projet: SRO non disponible", 'error')
            return False

        self._dlg.log_message(
            f"Mode Projet: chargement depuis BDD (SRO={sro})", 'info'
        )

        self._db_loader = DbLayerLoader()
        if not self._db_loader.connect():
            self._dlg.log_message(
                "Mode Projet: connexion PostgreSQL impossible", 'error'
            )
            return False

        # Always add layers to QGIS project so that get_layer_safe() and
        # other name-based lookups work inside workflows.
        # Cleanup after batch removes them if user didn't ask to keep.

        # Load infra_pt_pot (mandatory)
        self._pm_lyr_pot = self._db_loader.load_infra_pt_pot(
            sro, add_to_project=True
        )
        if not self._pm_lyr_pot or not self._pm_lyr_pot.isValid():
            self._dlg.log_message(
                "Mode Projet: couche infra_pt_pot non chargee", 'error'
            )
            return False

        count_pot = self._pm_lyr_pot.featureCount()
        self._dlg.log_message(
            f"  infra_pt_pot: {count_pot} poteaux", 'success'
        )

        # Load etude_cap_ft (optional)
        self._pm_lyr_cap = self._db_loader.load_etude_cap_ft(
            sro, add_to_project=True
        )
        if self._pm_lyr_cap and self._pm_lyr_cap.isValid():
            count_cap = self._pm_lyr_cap.featureCount()
            self._dlg.log_message(
                f"  etude_cap_ft: {count_cap} zones", 'success'
            )
        else:
            self._dlg.log_message(
                "  etude_cap_ft: non disponible (modules CAP FT/MAJ limites)",
                'warning'
            )
            self._pm_lyr_cap = None

        # Load etude_comac (optional)
        self._pm_lyr_com = self._db_loader.load_etude_comac(
            sro, add_to_project=True
        )
        if self._pm_lyr_com and self._pm_lyr_com.isValid():
            count_com = self._pm_lyr_com.featureCount()
            self._dlg.log_message(
                f"  etude_comac: {count_com} zones", 'success'
            )
        else:
            self._dlg.log_message(
                "  etude_comac: non disponible (module COMAC/MAJ limites)",
                'warning'
            )
            self._pm_lyr_com = None

        return True

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
        """Auto-detect study field name from layer."""
        if not layer or not layer.isValid():
            return ''
        import re
        if patterns is None:
            patterns = [r'nom[_ ]?etudes', r'etudes', r'ref_fci']
        for field in layer.fields():
            for pat in patterns:
                if re.search(pat, field.name(), re.IGNORECASE):
                    return field.name()
        # Fallback: first text field
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
        cb = self._current_done_cb
        if cb:
            liste_ft = result.get('liste_ft', [0, None, 0, None])
            liste_bt = result.get('liste_bt', [0, None, 0, None])
            ft_count = liste_ft[2] if len(liste_ft) > 2 else 0
            bt_count = liste_bt[2] if len(liste_bt) > 2 else 0
            self._batch_results['maj'] = result
            msg = f"Analyse: {ft_count} FT, {bt_count} BT"
            cb(True, msg)

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

        # Pass cached fddcpi2 data if available (e.g. Police C6 ran first)
        fddcpi = self._fddcpi_cache.get('cables') if self._fddcpi_cache else None
        self._comac_wf.start_analysis(
            lyr_pot, lyr_comac, col_comac,
            det.comac_dir, export_dir,
            fddcpi_cache=fddcpi
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

        # Inject cached fddcpi2 data if available (e.g. COMAC ran first)
        if self._fddcpi_cache:
            params['fddcpi_cables_cache'] = self._fddcpi_cache.get('cables')

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
