"""
Module de connexion PostgreSQL pour Police C6 v2.0
Connexion automatique via les credentials QGIS existants
"""

import psycopg2
from psycopg2 import sql
from dataclasses import dataclass
from typing import List, Optional, Tuple
from qgis.core import QgsSettings, QgsMessageLog, Qgis, QgsDataSourceUri


# Configuration cible
TARGET_HOST = "10.241.228.107"
TARGET_DATABASE = "auvergne"  # Nom en minuscule pour comparaison


@dataclass
class CableSegment:
    """Segment de câble découpé par l'infrastructure"""
    gid_dc2: int          # ID segment unique
    gid_dc: int           # ID câble original
    gid: int              # GID de référence
    sro: str
    nro: str
    length: float
    cab_type: str
    cab_capa: int         # Capacité FO
    cab_modulo: int
    isole: str
    date_modif: str
    modif_par: str
    cab_nature: str
    commentaire: str
    collecte: str
    cb_etiquet: str
    fon: str
    projet: str
    dce: str
    dist_type: str
    affectation: str
    posemode: int         # 0=souterrain, 1=aérien, 2=façade
    geom_wkt: str         # Géométrie en WKT


class DatabaseConnection:
    """
    Gère la connexion à la base PostgreSQL RIP AVG NGE.
    Récupère automatiquement les credentials depuis QGIS.
    """
    
    def __init__(self):
        self.connection = None
        self.host = None
        self.port = None
        self.database = None
        self.user = None
        self.password = None
        self._connection_name = None
    
    def find_auvergne_connection(self) -> Optional[str]:
        """
        Cherche la connexion PostgreSQL "Auvergne" dans les settings QGIS.
        Retourne le nom de la connexion si trouvée.
        """
        settings = QgsSettings()
        
        # Lister toutes les connexions PostgreSQL
        settings.beginGroup("PostgreSQL/connections")
        connections = settings.childGroups()
        settings.endGroup()
        
        for conn_name in connections:
            settings.beginGroup(f"PostgreSQL/connections/{conn_name}")
            host = settings.value("host", "")
            database = settings.value("database", "")
            settings.endGroup()
            
            # Vérifier si c'est la bonne connexion
            if host == TARGET_HOST or database.lower() == TARGET_DATABASE:
                QgsMessageLog.logMessage(
                    f"Connexion PostgreSQL trouvée: {conn_name} ({host}/{database})",
                    "PoleAerien", Qgis.Info
                )
                return conn_name
            
            # Chercher aussi par nom contenant "Auvergne"
            if "auvergne" in conn_name.lower():
                QgsMessageLog.logMessage(
                    f"Connexion PostgreSQL trouvée par nom: {conn_name}",
                    "PoleAerien", Qgis.Info
                )
                return conn_name
        
        return None
    
    def get_connection_params(self, connection_name: str) -> dict:
        """
        Récupère les paramètres de connexion depuis QGIS.
        """
        settings = QgsSettings()
        prefix = f"PostgreSQL/connections/{connection_name}"
        
        params = {
            'host': settings.value(f"{prefix}/host", TARGET_HOST),
            'port': int(settings.value(f"{prefix}/port", 5432)),
            'database': settings.value(f"{prefix}/database", ""),
            'user': settings.value(f"{prefix}/username", ""),
            'password': settings.value(f"{prefix}/password", ""),
        }
        
        # Si le mot de passe n'est pas stocké, essayer authcfg
        if not params['password']:
            authcfg = settings.value(f"{prefix}/authcfg", "")
            if authcfg:
                from qgis.core import QgsApplication
                auth_manager = QgsApplication.authManager()
                config = auth_manager.configAuthMethodConfig(authcfg)
                if config:
                    params['user'] = config.config('username', params['user'])
                    params['password'] = config.config('password', '')
        
        return params
    
    def connect(self) -> bool:
        """
        Établit la connexion à la base de données.
        Retourne True si succès, False sinon.
        """
        try:
            # Trouver la connexion Auvergne
            conn_name = self.find_auvergne_connection()
            if not conn_name:
                QgsMessageLog.logMessage(
                    f"Connexion PostgreSQL 'Auvergne' non trouvée. "
                    f"Veuillez configurer une connexion vers {TARGET_HOST}",
                    "PoleAerien", Qgis.Warning
                )
                return False
            
            self._connection_name = conn_name
            params = self.get_connection_params(conn_name)
            
            self.host = params['host']
            self.port = params['port']
            self.database = params['database']
            self.user = params['user']
            self.password = params['password']
            
            # Établir la connexion
            self.connection = psycopg2.connect(
                host=self.host,
                port=self.port,
                database=self.database,
                user=self.user,
                password=self.password
            )
            
            QgsMessageLog.logMessage(
                f"Connexion PostgreSQL établie: {self.host}/{self.database}",
                "PoleAerien", Qgis.Info
            )
            return True
            
        except Exception as e:
            QgsMessageLog.logMessage(
                f"Erreur connexion PostgreSQL: {e}",
                "PoleAerien", Qgis.Critical
            )
            return False
    
    def disconnect(self):
        """Ferme la connexion."""
        if self.connection:
            self.connection.close()
            self.connection = None
    
    def execute_fddcpi2(self, sro: str) -> List[CableSegment]:
        """
        Exécute la fonction fddcpi2 et retourne les câbles découpés.
        
        Args:
            sro: Code SRO (ex: '63041/B1I/PMZ/00003')
        
        Returns:
            Liste de CableSegment
        """
        if not self.connection:
            if not self.connect():
                return []
        
        try:
            cursor = self.connection.cursor()
            
            # Exécuter la fonction avec le SRO
            query = sql.SQL("SELECT *, ST_AsText(geom) as geom_wkt FROM rip_avg_nge.fddcpi2(%s)")
            cursor.execute(query, (sro,))
            
            rows = cursor.fetchall()
            segments = []
            
            for row in rows:
                # Mapper les colonnes vers CableSegment
                segment = CableSegment(
                    gid_dc2=row[0] or 0,
                    gid_dc=row[1] or 0,
                    gid=row[2] or 0,
                    sro=row[3] or '',
                    nro=row[4] or '',
                    length=row[5] or 0.0,
                    cab_type=row[6] or '',
                    cab_capa=row[7] or 0,
                    cab_modulo=row[8] or 0,
                    isole=row[9] or '',
                    date_modif=str(row[10]) if row[10] else '',
                    modif_par=row[11] or '',
                    cab_nature=row[13] or '',
                    commentaire=row[14] or '',
                    collecte=row[15] or '',
                    cb_etiquet=row[16] or '',
                    fon=row[17] or '',
                    projet=row[18] or '',
                    dce=row[19] or '',
                    dist_type=row[20] or '',
                    affectation=row[21] or '',
                    posemode=row[22] or 0,
                    geom_wkt=row[23] or ''  # ST_AsText(geom)
                )
                segments.append(segment)
            
            cursor.close()
            
            QgsMessageLog.logMessage(
                f"fddcpi2({sro}): {len(segments)} segments de câbles récupérés",
                "PoleAerien", Qgis.Info
            )
            
            return segments
            
        except Exception as e:
            QgsMessageLog.logMessage(
                f"Erreur exécution fddcpi2: {e}",
                "PoleAerien", Qgis.Warning
            )
            return []
    
    def get_cables_aeriens(self, sro: str) -> List[CableSegment]:
        """
        Récupère uniquement les câbles aériens (posemode=1).
        """
        all_segments = self.execute_fddcpi2(sro)
        return [s for s in all_segments if s.posemode == 1]

    def query_bpe_by_sro(self, sro: str) -> List[dict]:
        """
        Récupère les BPE d'un SRO avec leur géométrie et type.
        
        Args:
            sro: Code SRO
        
        Returns:
            Liste de dicts {gid, noe_type, noe_usage, inf_num, geom_wkt}
        """
        if not self.connection:
            if not self.connect():
                return []
        
        try:
            cursor = self.connection.cursor()
            query = sql.SQL(
                "SELECT gid, noe_type, noe_usage, inf_num, ST_AsText(geom) as geom_wkt "
                "FROM rip_avg_nge.bpe WHERE sro = %s"
            )
            cursor.execute(query, (sro,))
            rows = cursor.fetchall()
            
            results = []
            for row in rows:
                results.append({
                    'gid': row[0],
                    'noe_type': row[1] or '',
                    'noe_usage': row[2] or '',
                    'inf_num': row[3] or '',
                    'geom_wkt': row[4] or ''
                })
            
            cursor.close()
            
            QgsMessageLog.logMessage(
                f"BPE({sro}): {len(results)} boîtiers récupérés",
                "PoleAerien", Qgis.Info
            )
            
            return results
            
        except Exception as e:
            QgsMessageLog.logMessage(
                f"Erreur requête BPE: {e}",
                "PoleAerien", Qgis.Warning
            )
            return []


def extract_sro_from_layer(layer) -> Optional[str]:
    """
    Extrait le code SRO depuis une couche QGIS.
    Cherche dans les attributs 'sro', 'SRO', 'code_sro', etc.
    
    Args:
        layer: QgsVectorLayer
    
    Returns:
        Code SRO ou None
    """
    if not layer or not layer.isValid():
        return None
    
    # Noms de champs possibles pour le SRO
    sro_fields = ['sro', 'SRO', 'code_sro', 'CODE_SRO', 'pt_ad_sro', 'PT_AD_SRO']
    
    field_names = [f.name() for f in layer.fields()]
    
    for sro_field in sro_fields:
        if sro_field in field_names:
            # Récupérer la première valeur non vide
            for feature in layer.getFeatures():
                value = feature[sro_field]
                if value and str(value).strip():
                    return str(value).strip()
    
    return None


def extract_sro_from_layer_uri(layer) -> Optional[str]:
    """
    Extrait le code SRO depuis l'URI de la couche PostgreSQL.
    """
    if not layer or not layer.isValid():
        return None
    
    uri = QgsDataSourceUri(layer.source())
    
    # Le SRO peut être dans le filtre SQL
    sql_filter = uri.sql()
    if sql_filter:
        # Chercher un pattern comme sro = '...'
        import re
        match = re.search(r"sro\s*=\s*'([^']+)'", sql_filter, re.IGNORECASE)
        if match:
            return match.group(1)
    
    return None
