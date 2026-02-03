#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
MAJ SQL Background - Exécution des mises à jour BD en arrière-plan.

Ce module permet d'exécuter les MAJ FT/BT via SQL direct dans un QgsTask,
sans bloquer l'UI QGIS. L'UI reste 100% réactive pendant toute la durée
de la mise à jour.

Architecture:
- MajSqlBackgroundTask: QgsTask qui exécute les UPDATE SQL en background
- Connexion PostgreSQL directe via QSqlDatabase (pas via QGIS provider)
- Signaux pour progression et résultat
- layer.reload() + triggerRepaint() sur main thread après MAJ
"""

import time
from qgis.PyQt.QtCore import QObject, pyqtSignal
from qgis.PyQt.QtSql import QSqlDatabase, QSqlQuery
from qgis.core import (
    Qgis, QgsTask, QgsMessageLog, QgsDataSourceUri, QgsProject
)
import pandas as pd


class MajSqlSignals(QObject):
    """Signaux pour communication avec le main thread."""
    progress = pyqtSignal(int, str)  # (pourcentage, message)
    finished = pyqtSignal(dict)  # résultat final
    error = pyqtSignal(str)  # message d'erreur


class MajSqlBackgroundTask(QgsTask):
    """
    Tâche asynchrone pour MAJ BD via SQL direct.
    
    Exécute les UPDATE SQL dans un worker thread, sans bloquer l'UI.
    Les modifications sont faites directement en base PostgreSQL,
    puis la couche QGIS est rechargée sur le main thread.
    """
    
    def __init__(self, layer_name, data_ft, data_bt, db_uri):
        """
        Args:
            layer_name: Nom de la couche infra_pt_pot
            data_ft: DataFrame des MAJ FT (index=gid)
            data_bt: DataFrame des MAJ BT (index=gid)
            db_uri: QgsDataSourceUri avec les infos de connexion
        """
        super().__init__("MAJ BD SQL Background", QgsTask.CanCancel)
        self.layer_name = layer_name
        self.data_ft = data_ft
        self.data_bt = data_bt
        self.db_uri = db_uri
        self.signals = MajSqlSignals()
        self.exception = None
        self.result = {'ft_updated': 0, 'bt_updated': 0, 'gids_pot_ac': []}
    
    def run(self):
        """Exécute les MAJ SQL en background (worker thread)."""
        try:
            t0 = time.perf_counter()
            
            # Connexion PostgreSQL
            conn_name = f"maj_bg_{id(self)}"
            db = QSqlDatabase.addDatabase("QPSQL", conn_name)
            db.setHostName(self.db_uri.host())
            db.setPort(int(self.db_uri.port()) if self.db_uri.port() else 5432)
            db.setDatabaseName(self.db_uri.database())
            db.setUserName(self.db_uri.username())
            db.setPassword(self.db_uri.password())
            
            if not db.open():
                raise RuntimeError(f"Connexion DB échouée: {db.lastError().text()}")
            
            schema = self.db_uri.schema().replace('"', '""')
            table = self.db_uri.table().replace('"', '""')
            
            # Récupérer les colonnes existantes dans la table
            self._existing_columns = self._get_table_columns(db, schema, table)
            
            try:
                # Démarrer transaction
                db.transaction()
                
                # MAJ FT
                if self.data_ft is not None and not self.data_ft.empty:
                    ft_count = self._update_ft_sql(db, schema, table)
                    self.result['ft_updated'] = ft_count
                
                if self.isCanceled():
                    db.rollback()
                    return False
                
                # MAJ BT
                if self.data_bt is not None and not self.data_bt.empty:
                    bt_count = self._update_bt_sql(db, schema, table)
                    self.result['bt_updated'] = bt_count
                
                if self.isCanceled():
                    db.rollback()
                    return False
                
                # Commit transaction
                if not db.commit():
                    raise RuntimeError(f"Commit échoué: {db.lastError().text()}")
                
                t1 = time.perf_counter()
                QgsMessageLog.logMessage(
                    f"[MAJ-SQL-BG] Terminé: {self.result['ft_updated']} FT, "
                    f"{self.result['bt_updated']} BT en {t1-t0:.1f}s",
                    "PoleAerien", Qgis.Info
                )
                
            finally:
                db.close()
                QSqlDatabase.removeDatabase(conn_name)
            
            return True
            
        except Exception as e:
            self.exception = str(e)
            QgsMessageLog.logMessage(
                f"[MAJ-SQL-BG] Erreur: {e}", "PoleAerien", Qgis.Critical
            )
            return False
    
    def _update_ft_sql(self, db, schema, table):
        """Exécute les UPDATE SQL pour les FT."""
        total = len(self.data_ft)
        count = 0
        gids_pot_ac = []
        
        for gid, row in self.data_ft.iterrows():
            if self.isCanceled():
                return count
            
            count += 1
            action = str(row.get("action", "")).upper()
            
            # Récupérer les valeurs actuelles (inf_num, commentair) pour concaténation
            current_inf_num = ""
            current_comment = ""
            select_sql = f'SELECT inf_num, commentair FROM "{schema}"."{table}" WHERE gid = {int(gid)}'
            select_query = QSqlQuery(db)
            if select_query.exec_(select_sql) and select_query.next():
                current_inf_num = select_query.value(0) or ""
                current_comment = select_query.value(1) or ""
            
            # Construire la requête UPDATE
            updates = []
            
            if action == "IMPLANTATION":
                updates.append("etat = 'FT KO'")
                updates.append("inf_propri = 'RAUV'")
                updates.append("inf_type = 'POT-AC'")
                updates.append("dce = 'O'")
                if row.get("inf_mat_replace"):
                    mat = self._escape_sql(row["inf_mat_replace"])
                    updates.append(f"inf_mat = '{mat}'")
                # Sauvegarder ancien inf_num dans nommage_fibees avant de le vider
                if current_inf_num and self._column_exists("nommage_fibees"):
                    ancien_num = self._escape_sql(current_inf_num)
                    updates.append(f"nommage_fibees = '{ancien_num}'")
                gids_pot_ac.append(int(gid))
            else:
                if row.get("etat"):
                    etat = self._escape_sql(row["etat"])
                    updates.append(f"etat = '{etat}'")
                if row.get("inf_mat_replace"):
                    mat = self._escape_sql(row["inf_mat_replace"])
                    updates.append(f"inf_mat = '{mat}'")
                updates.append("dce = 'O'")
            
            if row.get("etiquette_jaune") and self._column_exists("etiquette_jaune"):
                val = self._escape_sql(row["etiquette_jaune"])
                updates.append(f"etiquette_jaune = '{val}'")
            
            if row.get("etiquette_orange") and self._column_exists("etiquette_orange"):
                val = self._escape_sql(row["etiquette_orange"])
                updates.append(f"etiquette_orange = '{val}'")
            
            if row.get("transition_aerosout") and self._column_exists("transition_aerosout"):
                val = self._escape_sql(row["transition_aerosout"])
                updates.append(f"transition_aerosout = '{val}'")
            
            # Zone privée: concaténer /PRIVE au commentaire existant
            zone_privee = str(row.get("zone_privee", "")).strip().upper()
            if zone_privee == "X" and self._column_exists("commentair"):
                if "/PRIVE" not in str(current_comment).upper():
                    new_comment = f"{current_comment}/PRIVE" if current_comment else "/PRIVE"
                    new_comment_escaped = self._escape_sql(new_comment)
                    updates.append(f"commentair = '{new_comment_escaped}'")
            
            # Transition aérosout: concaténer /AEROSOUTRANSI au commentaire existant
            transition = str(row.get("transition_aerosout", "")).strip().upper()
            if transition == "OUI" and self._column_exists("commentair"):
                # Récupérer le commentaire actuel (peut avoir été modifié par zone_privee)
                comment_base = current_comment
                # Vérifier si on a déjà ajouté /PRIVE dans cette itération
                for upd in updates:
                    if "commentair = " in upd:
                        # Extraire la valeur déjà définie
                        comment_base = upd.split("'")[1] if "'" in upd else current_comment
                        updates.remove(upd)
                        break
                if "/AEROSOUTRANSI" not in str(comment_base).upper():
                    new_comment = f"{comment_base}/AEROSOUTRANSI" if comment_base else "/AEROSOUTRANSI"
                    new_comment_escaped = self._escape_sql(new_comment)
                    updates.append(f"commentair = '{new_comment_escaped}'")
            
            if updates:
                sql = f'UPDATE "{schema}"."{table}" SET {", ".join(updates)} WHERE gid = {int(gid)}'
                query = QSqlQuery(db)
                if not query.exec_(sql):
                    QgsMessageLog.logMessage(
                        f"[MAJ-SQL-BG] Erreur FT gid={gid}: {query.lastError().text()}",
                        "PoleAerien", Qgis.Warning
                    )
            
            # Progression
            if count % 5 == 0 or count == total:
                pct = 10 + int((count / total) * 40)  # 10% -> 50%
                self.signals.progress.emit(pct, f"MAJ FT: {count}/{total}")
        
        self.result['gids_pot_ac'] = gids_pot_ac
        return count
    
    def _update_bt_sql(self, db, schema, table):
        """Exécute les UPDATE SQL pour les BT."""
        total = len(self.data_bt)
        count = 0
        gids_pot_ac_bt = []
        
        for gid, row in self.data_bt.iterrows():
            if self.isCanceled():
                return count
            
            count += 1
            
            # Récupérer les valeurs actuelles (inf_num, commentair) pour concaténation
            current_inf_num = ""
            current_comment = ""
            select_sql = f'SELECT inf_num, commentair FROM "{schema}"."{table}" WHERE gid = {int(gid)}'
            select_query = QSqlQuery(db)
            if select_query.exec_(select_sql) and select_query.next():
                current_inf_num = select_query.value(0) or ""
                current_comment = select_query.value(1) or ""
            
            # Construire la requête UPDATE
            updates = []
            
            if str(row.get("Portée molle", "")).upper() == "X":
                updates.append("etat = 'PORTEE MOLLE'")
            else:
                if row.get("inf_type"):
                    val = self._escape_sql(row["inf_type"])
                    updates.append(f"inf_type = '{val}'")
                    # Si BT KO (IMPLANTATION), sauvegarder inf_num dans nommage_fibees
                    if val == "POT-AC" and current_inf_num and self._column_exists("nommage_fibees"):
                        ancien_num = self._escape_sql(current_inf_num)
                        updates.append(f"nommage_fibees = '{ancien_num}'")
                        gids_pot_ac_bt.append(int(gid))
                if row.get("inf_propri"):
                    val = self._escape_sql(row["inf_propri"])
                    updates.append(f"inf_propri = '{val}'")
                updates.append("noe_usage = 'DI'")
                if row.get("typ_po_mod"):
                    val = self._escape_sql(row["typ_po_mod"])
                    updates.append(f"inf_mat = '{val}'")
                if row.get("etat"):
                    val = self._escape_sql(row["etat"])
                    updates.append(f"etat = '{val}'")
                updates.append("dce = 'O'")
            
            if row.get("etiquette_orange") and self._column_exists("etiquette_orange"):
                val = self._escape_sql(row["etiquette_orange"])
                updates.append(f"etiquette_orange = '{val}'")
            
            # Zone privée BT: concaténer /PRIVE au commentaire existant
            zone_privee = str(row.get("zone_privee", "")).strip().upper()
            if zone_privee == "X" and self._column_exists("commentair"):
                if "/PRIVE" not in str(current_comment).upper():
                    new_comment = f"{current_comment}/PRIVE" if current_comment else "/PRIVE"
                    new_comment_escaped = self._escape_sql(new_comment)
                    updates.append(f"commentair = '{new_comment_escaped}'")
            
            if updates:
                sql = f'UPDATE "{schema}"."{table}" SET {", ".join(updates)} WHERE gid = {int(gid)}'
                query = QSqlQuery(db)
                if not query.exec_(sql):
                    QgsMessageLog.logMessage(
                        f"[MAJ-SQL-BG] Erreur BT gid={gid}: {query.lastError().text()}",
                        "PoleAerien", Qgis.Warning
                    )
            
            # Progression
            if count % 5 == 0 or count == total:
                pct = 50 + int((count / total) * 40)  # 50% -> 90%
                self.signals.progress.emit(pct, f"MAJ BT: {count}/{total}")
        
        return count
    
    def _escape_sql(self, value):
        """Échappe une valeur pour SQL (protection injection)."""
        if value is None:
            return ""
        return str(value).replace("'", "''")
    
    def _get_table_columns(self, db, schema, table):
        """Récupère la liste des colonnes existantes dans la table."""
        columns = set()
        # Nettoyer les doubles quotes pour la requête
        schema_clean = schema.replace('""', '"')
        table_clean = table.replace('""', '"')
        sql = f"""
            SELECT column_name 
            FROM information_schema.columns 
            WHERE table_schema = '{schema_clean}' AND table_name = '{table_clean}'
        """
        query = QSqlQuery(db)
        if query.exec_(sql):
            while query.next():
                columns.add(query.value(0))
        return columns
    
    def _column_exists(self, col_name):
        """Vérifie si une colonne existe dans la table."""
        return col_name in getattr(self, '_existing_columns', set())
    
    def finished(self, success):
        """Callback sur main thread après exécution."""
        if success:
            self.signals.finished.emit({
                'layer_name': self.layer_name,
                'ft_updated': self.result['ft_updated'],
                'bt_updated': self.result['bt_updated'],
                'gids_pot_ac': self.result.get('gids_pot_ac', [])
            })
        else:
            self.signals.error.emit(self.exception or "Annulé par l'utilisateur")


def get_layer_db_uri(layer_name):
    """Récupère l'URI de connexion PostgreSQL d'une couche."""
    lyrs = QgsProject.instance().mapLayersByName(layer_name)
    if not lyrs or not lyrs[0].isValid():
        return None
    layer = lyrs[0]
    if layer.dataProvider().name() != "postgres":
        return None
    return QgsDataSourceUri(layer.dataProvider().dataSourceUri())


def reload_layer(layer_name):
    """Recharge une couche QGIS après MAJ SQL directe."""
    lyrs = QgsProject.instance().mapLayersByName(layer_name)
    if lyrs and lyrs[0].isValid():
        layer = lyrs[0]
        layer.reload()
        layer.triggerRepaint()
        return True
    return False
