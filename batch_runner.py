# -*- coding: utf-8 -*-
"""
Batch Runner - Sequential multi-module execution controller.

Executes selected modules one after another, collecting results
and reporting progress to the UI via signals.
"""

from qgis.PyQt.QtCore import QObject, pyqtSignal, QTimer


# Module registry: key → display name
MODULE_REGISTRY = {
    'maj':       '0. MAJ BD',
    'capft':     '1. VERIF CAP_FT',
    'comac':     '2. VERIF COMAC',
    'c6bd':      '3. C6 vs BD',
    'police_c6': '4. POLICE C6',
    'c6c3a':     '5. C6-C3A-C7-BD',
}


class BatchRunner(QObject):
    """Runs selected modules sequentially.

    Signals:
        module_started(str, int, int): module_key, current_index, total
        module_finished(str, bool, str): module_key, success, message
        batch_progress(int): overall percent 0-100
        batch_finished(dict): {module_key: {'success': bool, 'message': str}}
        batch_cancelled(): emitted on user cancel
        log_message(str, str): message, level ('info'|'success'|'warning'|'error')
    """

    module_started = pyqtSignal(str, int, int)
    module_finished = pyqtSignal(str, bool, str)
    batch_progress = pyqtSignal(int)
    batch_finished = pyqtSignal(dict)
    batch_cancelled = pyqtSignal()
    log_message = pyqtSignal(str, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._queue = []
        self._results = {}
        self._current_index = 0
        self._running = False
        self._cancelled = False

        # Module launchers: set by the orchestrator (PoleAerien.py)
        # Each launcher is a callable that starts the module and connects
        # its finished signal to self._on_module_done
        self._launchers = {}
        self._current_module_key = None

    @property
    def is_running(self) -> bool:
        return self._running

    def set_launcher(self, module_key: str, launcher_fn):
        """Register a launcher function for a module.

        Args:
            module_key: Key from MODULE_REGISTRY
            launcher_fn: Callable that starts the module analysis.
                         Must accept a callback: launcher_fn(on_done_callback)
                         where on_done_callback(success: bool, message: str)
        """
        self._launchers[module_key] = launcher_fn

    def start(self, module_keys: list):
        """Start batch execution of selected modules.

        Args:
            module_keys: Ordered list of module keys to execute.
        """
        if self._running:
            return

        valid_keys = [k for k in module_keys if k in self._launchers]
        if not valid_keys:
            self.log_message.emit("Aucun module valide sélectionné.", 'warning')
            return

        self._queue = list(valid_keys)
        self._results = {}
        self._current_index = 0
        self._running = True
        self._cancelled = False

        total = len(self._queue)
        names = ', '.join(MODULE_REGISTRY.get(k, k) for k in self._queue)
        self.log_message.emit(
            f"Lancement de {total} module(s) : {names}",
            'info'
        )
        self.batch_progress.emit(0)

        # Start first module (deferred to let signals connect)
        QTimer.singleShot(100, self._run_next)

    def cancel(self):
        """Cancel the batch execution."""
        if not self._running:
            return
        self._cancelled = True
        self._running = False
        self.log_message.emit("Exécution annulée par l'utilisateur.", 'warning')
        self.batch_cancelled.emit()

    def _run_next(self):
        """Launch the next module in the queue."""
        if self._cancelled or not self._running:
            return

        if self._current_index >= len(self._queue):
            self._finish_batch()
            return

        key = self._queue[self._current_index]
        total = len(self._queue)
        self._current_module_key = key

        name = MODULE_REGISTRY.get(key, key)
        self.log_message.emit(
            f"--- Module {self._current_index + 1}/{total} : {name} ---",
            'info'
        )
        self.module_started.emit(key, self._current_index, total)

        launcher = self._launchers.get(key)
        if launcher is None:
            self._on_module_done(False, f"Lanceur non configuré pour '{key}'")
            return

        try:
            launcher(self._on_module_done)
        except Exception as exc:
            self._on_module_done(False, f"Erreur lancement : {exc}")

    def _on_module_done(self, success: bool, message: str = ''):
        """Called when a module finishes (success or error).

        Args:
            success: True if module completed successfully.
            message: Optional result/error message.
        """
        if not self._running:
            return

        key = self._current_module_key
        name = MODULE_REGISTRY.get(key, key)

        self._results[key] = {'success': success, 'message': message}

        level = 'success' if success else 'error'
        status = 'OK' if success else 'ERREUR'
        self.log_message.emit(f"{name} : {status} {message}", level)
        self.module_finished.emit(key, success, message)

        # Update progress
        total = len(self._queue)
        self._current_index += 1
        pct = int((self._current_index / total) * 100) if total > 0 else 100
        self.batch_progress.emit(pct)

        # Next module (deferred to allow UI refresh)
        QTimer.singleShot(200, self._run_next)

    def _finish_batch(self):
        """Complete the batch and emit results."""
        self._running = False
        self._current_module_key = None

        successes = sum(1 for r in self._results.values() if r['success'])
        total = len(self._results)
        errors = total - successes

        if errors == 0:
            self.log_message.emit(
                f"Tous les modules terminés avec succès ({total}/{total}).",
                'success'
            )
        else:
            self.log_message.emit(
                f"Terminé : {successes}/{total} OK, {errors} erreur(s).",
                'warning'
            )

        self.batch_progress.emit(100)
        self.batch_finished.emit(dict(self._results))
