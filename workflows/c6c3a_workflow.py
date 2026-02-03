# -*- coding: utf-8 -*-
"""
Workflow pour la comparaison C6 vs C3A vs BD.
Orchestre la comparaison entre les fichiers Excel C6, C3A, C7 et les données QGIS.
"""

from qgis.PyQt.QtCore import QObject, pyqtSignal
from qgis.core import Qgis, QgsMessageLog, QgsApplication
from ..C6_vs_C3A_vs_Bd import C6_vs_C3A_vs_Bd
from ..async_tasks import ExcelExportTask, run_async_task
import os
import pandas as pd
import datetime
import traceback

class C6C3AWorkflow(QObject):
    """
    Contrôleur pour le flux de comparaison C6 vs C3A vs BD.
    
    Gère:
    1. Lecture des fichiers Excel (C6, C3A, C7)
    2. Extraction et comparaison avec les données QGIS
    3. Fusion des résultats
    4. Export Excel
    """
    
    # Signaux
    progress_changed = pyqtSignal(int)
    message_received = pyqtSignal(str, str)  # message, couleur
    analysis_finished = pyqtSignal(dict)
    export_finished = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.logic = C6_vs_C3A_vs_Bd()
        self.current_task = None

    def start_analysis(self, params):
        """
        Lance l'analyse complète.
        
        Args:
            params (dict): Paramètres d'analyse
                - fichier_c6 (str): Chemin fichier C6
                - fichier_c3a (str): Chemin fichier C3A (optionnel si mode QGIS)
                - fichier_c7 (str): Chemin fichier C7
                - chemin_export (str): Dossier export
                - mode_c3a (str): 'QGIS' ou 'EXCEL'
                - table_infra (str): Nom couche infra_pt_pot
                - table_cmd (str): Nom couche commande (si mode QGIS)
                - table_decoupage (str): Nom couche découpage
                - champs_dcp (str): Champ découpage
                - valeur_dcp (str): Valeur découpage
        """
        self.progress_changed.emit(5)
        self.message_received.emit("Démarrage de l'analyse C6 vs C3A vs BD...", "black")
        
        try:
            # Note: Pour l'instant, on exécute la logique d'analyse sur le thread principal
            # car elle mélange lecture Excel et accès QGIS via C6_vs_C3A_vs_Bd.
            # Idéalement, il faudrait séparer lecture Excel (Thread) et QGIS (Main).
            # Mais C6_vs_C3A_vs_Bd est une classe monolithique pour l'instant.
            
            fichier_c6 = params['fichier_c6']
            fichier_c3a = params.get('fichier_c3a', '')
            fichier_c7 = params['fichier_c7']
            mode_c3a = params['mode_c3a']
            
            # 1. Lecture C6
            self.message_received.emit(f"Lecture fichier C6: {os.path.basename(fichier_c6)}", "black")
            df_init = pd.DataFrame(columns=["N° appui", "Nature des travaux"], dtype="object") # dtype object pour compatibilité
            df_c6_excel, df_c6_excel_rempl = self.logic.LectureFichiersExcelsC6(df_init, fichier_c6)
            self.progress_changed.emit(20)
            
            # 2. Comparaison BD
            self.message_received.emit("Comparaison avec la Base de Données QGIS...", "black")
            df_bd, df_bd_rempl = self.logic.liste_poteau_etudes(
                params['table_infra'], 
                params['table_decoupage'], 
                params['champs_dcp'], 
                params['valeur_dcp']
            )
            self.progress_changed.emit(40)
            
            # 3. Comparaison C3A
            df_c3a, df_c3a_rempl = pd.DataFrame(), pd.DataFrame()
            if mode_c3a == 'QGIS':
                self.message_received.emit("Comparaison C3A via QGIS...", "black")
                df_c3a, df_c3a_rempl = self.logic.liste_poteau_c3a_qgis(
                    params['table_cmd'], 
                    params['table_decoupage'], 
                    params['champs_dcp'], 
                    params['valeur_dcp']
                )
            else:
                self.message_received.emit(f"Lecture fichier C3A: {os.path.basename(fichier_c3a)}", "black")
                df_c3a, df_c3a_rempl = self.logic.liste_poteau_c3a_excel(fichier_c3a)
            self.progress_changed.emit(60)
            
            # 4. Lecture C7
            self.message_received.emit(f"Lecture fichier C7: {os.path.basename(fichier_c7)}", "black")
            df_c7 = self.logic.lectureFichierC7(fichier_c7)
            self.progress_changed.emit(70)
            
            # 5. Fusion et vérifications
            if df_c6_excel.empty or df_bd.empty or df_c3a.empty:
                missing = []
                if df_c6_excel.empty: missing.append("Données C6")
                if df_bd.empty: missing.append("Données BD")
                if df_c3a.empty: missing.append("Données C3A")
                self.message_received.emit(f"Attention: Des données sont vides ou introuvables ({', '.join(missing)})", "orange")
            
            # Préparation dataframes finaux (fusion)
            # Logique copiée de comparaisonC6C3aBd
            
            # Fusion principale (Tous les appuis)
            df_final = pd.DataFrame()
            if not df_bd.empty:
                df_final = df_bd.copy()
            
            if not df_c3a.empty:
                if df_final.empty:
                    df_final = df_c3a.copy()
                else:
                    df_final = pd.merge(df_final, df_c3a, on="N° appui", how="outer")
            
            if not df_c6_excel.empty:
                if df_final.empty:
                    df_final = df_c6_excel.copy()
                else:
                    df_final = pd.merge(df_final, df_c6_excel, on="N° appui", how="outer")
            
            # Fusion remplacements (Uniquement "Nature des travaux" == OUI)
            df_final_rempl = pd.DataFrame()
            if not df_bd_rempl.empty:
                df_final_rempl = df_bd_rempl.copy()
                
            if not df_c3a_rempl.empty:
                if df_final_rempl.empty:
                    df_final_rempl = df_c3a_rempl.copy()
                else:
                    df_final_rempl = pd.merge(df_final_rempl, df_c3a_rempl, on="N° appui", how="outer")
            
            if not df_c6_excel_rempl.empty:
                if df_final_rempl.empty:
                    df_final_rempl = df_c6_excel_rempl.copy()
                else:
                    df_final_rempl = pd.merge(df_final_rempl, df_c6_excel_rempl, on="N° appui", how="outer")
            
            if not df_c7.empty:
                if not df_final_rempl.empty:
                     df_final_rempl = pd.merge(df_final_rempl, df_c7, on="N° appui", how="outer")
            
            # Remplissage des valeurs manquantes pour analyse
            for col in df_final.columns:
                if "inf_num" in col or "Excel" in col:
                    df_final[col] = df_final[col].fillna("ABSENT")
            
            for col in df_final_rempl.columns:
                if "inf_num" in col or "Excel" in col or "Fichier (C7)" in col:
                    df_final_rempl[col] = df_final_rempl[col].fillna("ABSENT")
            
            self.progress_changed.emit(85)
            
            # Préparation résultat
            result = {
                'df_final': df_final,
                'df_final_rempl': df_final_rempl,
                'chemin_export': params['chemin_export'],
                'success': True
            }
            
            self.analysis_finished.emit(result)
            
        except Exception as e:
            QgsMessageLog.logMessage(f"Erreur C6C3AWorkflow: {traceback.format_exc()}", "PoleAerien", Qgis.Critical)
            self.error_occurred.emit(str(e))

    def start_export(self, result):
        """
        Lance l'export Excel (tâche asynchrone).
        """
        df_final = result.get('df_final')
        df_final_rempl = result.get('df_final_rempl')
        chemin_export = result.get('chemin_export')
        
        if df_final is None or df_final_rempl is None:
            self.error_occurred.emit("Données manquantes pour l'export")
            return
            
        # On définit une fonction wrapper pour l'export qui fait les 2 exports
        def run_export(df_final, df_final_rempl, chemin_export):
            # 1. Export Analyse Globale
            date_str = datetime.datetime.now().strftime("%d-%m-%Y_%H_%M")
            file_name_1 = f"{chemin_export}{os.sep}ANALYSE_C6_C3A_BD_{date_str}.xlsx"
            nb_err_1 = self.logic.ecrictureExcel(df_final, file_name_1)
            
            # 2. Export Analyse Remplacement (avec C7)
            file_name_2 = f"{chemin_export}{os.sep}ANALYSE_C6_C7_C3A_BD_{date_str}.xlsx"
            nb_err_2 = self.logic.ecrictureExcelC6C7C3aBd(df_final_rempl, file_name_2)
            
            return {
                'file_1': file_name_1,
                'nb_err_1': nb_err_1,
                'file_2': file_name_2,
                'nb_err_2': nb_err_2
            }

        self.current_task = ExcelExportTask(
            "Export C6-C3A-BD",
            run_export,
            args=[df_final, df_final_rempl, chemin_export],
            payload={}
        )
        
        self.current_task.signals.progress.connect(self.progress_changed)
        self.current_task.signals.message.connect(self.message_received)
        self.current_task.signals.finished.connect(self.export_finished)
        self.current_task.signals.error.connect(self.error_occurred)
        
        run_async_task(self.current_task)
