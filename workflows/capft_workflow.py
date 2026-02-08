# -*- coding: utf-8 -*-
"""
Workflow pour l'analyse CAP_FT.
Orchestre l'extraction des données, l'analyse asynchrone et l'export Excel.
"""

from qgis.PyQt.QtCore import QObject, pyqtSignal
from qgis.core import Qgis, QgsMessageLog
from ..CapFt import CapFt
from ..async_tasks import CapFtTask, ExcelExportTask, run_async_task
from ..core_utils import build_export_path
import os
import copy

class CapFtWorkflow(QObject):
    """
    Contrôleur pour le flux d'analyse CAP_FT.
    Gère:
    1. L'extraction des données QGIS (Main Thread)
    2. Le lancement de la tâche asynchrone d'analyse (Worker Thread)
    3. Le lancement de la tâche asynchrone d'export (Worker Thread)
    4. La communication des résultats via signaux
    """
    
    # Signaux relayés vers l'UI
    progress_changed = pyqtSignal(int)
    message_received = pyqtSignal(str, str)  # message, couleur
    analysis_finished = pyqtSignal(dict)
    export_finished = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.cap_logic = CapFt()
        self.current_task = None

    def start_analysis(self, lyr_pot, lyr_cap, col_cap, chemin_cap, chemin_export):
        """
        Lance l'analyse CAP_FT.
        
        Args:
            lyr_pot (QgsVectorLayer): Couche Poteaux (infra_pt_pot)
            lyr_cap (QgsVectorLayer): Couche Etude CAP_FT
            col_cap (str): Colonne identifiant l'étude dans la couche CAP_FT
            chemin_cap (str): Répertoire des fichiers CAP_FT (Fiches Appuis)
            chemin_export (str): Chemin pour le fichier Excel de sortie
        """
        if not lyr_pot or not lyr_cap:
            self.error_occurred.emit("Couches invalides ou manquantes")
            return

        if not os.path.isdir(chemin_cap):
            self.error_occurred.emit("Répertoire CAP_FT invalide")
            return

        fichier_export = build_export_path(chemin_export, "analyse_cap_ft.xlsx")

        # 1. Extraction des données (Main Thread)
        try:
            doublons, hors_etude = self.cap_logic.verificationsDonneesCapft(
                lyr_pot.name(), lyr_cap.name(), col_cap
            )
            dico_qgis, dico_poteaux_prives = self.cap_logic.liste_poteau_cap_ft(
                lyr_pot.name(), lyr_cap.name(), col_cap
            )
        except ValueError as e:
            self.error_occurred.emit(str(e))
            return
        except Exception as e:
            QgsMessageLog.logMessage(f"Erreur extraction CAP_FT: {e}", "PoleAerien", Qgis.Critical)
            self.error_occurred.emit(f"Erreur technique extraction: {e}")
            return

        # 2. Préparation des données pour la Task
        params = {
            'chemin_cap_ft': chemin_cap,
            'fichier_export': fichier_export
        }
        
        # Deep copy pour éviter mutation
        qgis_data = {
            'doublons': list(doublons),
            'hors_etude': list(hors_etude),
            'dico_qgis': copy.deepcopy(dico_qgis),
            'dico_poteaux_prives': copy.deepcopy(dico_poteaux_prives)
        }

        # 3. Lancement Task Asynchrone
        self.current_task = CapFtTask(params, qgis_data)
        
        self.current_task.signals.progress.connect(self.progress_changed)
        self.current_task.signals.message.connect(self.message_received)
        self.current_task.signals.finished.connect(self.analysis_finished)
        self.current_task.signals.error.connect(self.error_occurred)
        
        run_async_task(self.current_task)

    def start_export(self, result):
        """
        Lance l'export Excel suite à l'analyse.
        
        Args:
            result (dict): Résultat de l'analyse CAP_FT contenant les données à exporter
        """
        if not result.get('pending_export'):
            return

        self.current_task = ExcelExportTask(
            "Export CAP_FT",
            self.cap_logic.ecrireResultatsAnalyseExcelsCapFt,
            args=[result['resultats'], result['fichier_export']],
            payload=result,
        )
        
        self.current_task.signals.progress.connect(self.progress_changed)
        self.current_task.signals.message.connect(self.message_received)
        self.current_task.signals.finished.connect(self.export_finished)
        self.current_task.signals.error.connect(self.error_occurred)
        
        run_async_task(self.current_task)
