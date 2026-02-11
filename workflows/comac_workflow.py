# -*- coding: utf-8 -*-
"""
Workflow pour l'analyse COMAC.
Orchestre l'extraction des données, l'analyse asynchrone et l'export Excel.
"""

from qgis.PyQt.QtCore import QObject, pyqtSignal
from qgis.core import Qgis, QgsMessageLog
from ..Comac import Comac
from ..async_tasks import ComacTask, ExcelExportTask, run_async_task
from ..core_utils import build_export_path
from ..db_connection import extract_sro_from_layer
from ..cable_analyzer import extraire_appuis_wkb
import os
import copy

class ComacWorkflow(QObject):
    """
    Contrôleur pour le flux d'analyse COMAC.
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
        self.comac_logic = Comac()
        self.current_task = None

    def start_analysis(self, lyr_pot, lyr_comac, col_comac, chemin_comac, chemin_export,
                        fddcpi_cache=None, sro_appuis_cache=None):
        """
        Lance l'analyse COMAC.
        
        Args:
            lyr_pot (QgsVectorLayer): Couche Poteaux (infra_pt_pot)
            lyr_comac (QgsVectorLayer): Couche Etude COMAC
            col_comac (str): Colonne identifiant l'etude dans la couche COMAC
            chemin_comac (str): Repertoire des fichiers COMAC
            chemin_export (str): Chemin pour le fichier Excel de sortie
            fddcpi_cache (list|None): CableSegment list from previous fddcpi2 call (batch optimization)
            sro_appuis_cache (dict|None): {'sro': str, 'appuis_wkb': list} from another module (batch optimization)
        """
        if not lyr_pot or not lyr_comac:
            self.error_occurred.emit("Couches invalides ou manquantes")
            return

        if not os.path.isdir(chemin_comac):
            self.error_occurred.emit("Repertoire COMAC invalide")
            return

        # 1. Extraction des donnees (Main Thread) - passe unique
        try:
            doublons, hors_etude, dico_qgis, dico_poteaux_prives = (
                self.comac_logic.extraire_donnees_comac(
                    lyr_pot.name(), lyr_comac.name(), col_comac
                )
            )
        except ValueError as e:
            self.error_occurred.emit(str(e))
            return
        except Exception as e:
            QgsMessageLog.logMessage(f"Erreur extraction COMAC: {e}", "PoleAerien", Qgis.Critical)
            self.error_occurred.emit(f"Erreur technique extraction: {e}")
            return

        # 2. Extraction SRO + appuis WKB pour verif cables (Main Thread)
        # Utiliser le cache batch si disponible (evite double extraction)
        if sro_appuis_cache:
            sro = sro_appuis_cache.get('sro')
            appuis_wkb = sro_appuis_cache.get('appuis_wkb', [])
        else:
            sro = extract_sro_from_layer(lyr_pot)
            appuis_wkb = []
            if sro:
                appuis_wkb = extraire_appuis_wkb(lyr_pot)
            else:
                QgsMessageLog.logMessage(
                    "[COMAC] SRO non trouve - verif cables desactivee",
                    "PoleAerien", Qgis.Warning
                )

        # 3. Préparation des données pour la Task
        fichier_export = build_export_path(chemin_export, "ANALYSE_COMAC.xlsx")
        
        params = {
            'chemin_comac': chemin_comac,
            'fichier_export': fichier_export,
            'zone_climatique': 'ZVN',
            'sro': sro,
            'fddcpi_cables_cache': fddcpi_cache,
        }
        
        # Deep copy pour éviter mutation
        qgis_data = {
            'doublons': list(doublons),
            'hors_etude': list(hors_etude),
            'dico_qgis': copy.deepcopy(dico_qgis),
            'dico_poteaux_prives': copy.deepcopy(dico_poteaux_prives),
            'appuis': appuis_wkb
        }

        # 3. Lancement Task Asynchrone
        self.current_task = ComacTask(params, qgis_data)
        
        self.current_task.signals.progress.connect(self.progress_changed)
        self.current_task.signals.message.connect(self.message_received)
        self.current_task.signals.finished.connect(self.analysis_finished)
        self.current_task.signals.error.connect(self.error_occurred)
        
        # Utilisation de run_async_task pour compatibilité avec async_tasks.py
        run_async_task(self.current_task)

    def start_export(self, result):
        """
        Lance l'export Excel suite à l'analyse.
        
        Args:
            result (dict): Résultat de l'analyse COMAC contenant les données à exporter
        """
        if not result.get('pending_export'):
            return

        self.current_task = ExcelExportTask(
            "Export COMAC",
            self.comac_logic.ecrireResultatsAnalyseExcels,
            args=[
                result['resultats'],
                result['fichier_export'],
                result.get('dico_verif_secu'),
                result.get('verif_cables'),
                result.get('verif_boitiers')
            ],
            payload=result,
        )
        
        self.current_task.signals.progress.connect(self.progress_changed)
        self.current_task.signals.message.connect(self.message_received)
        self.current_task.signals.finished.connect(self.export_finished) # Signal spécifique export
        self.current_task.signals.error.connect(self.error_occurred)
        
        run_async_task(self.current_task)
