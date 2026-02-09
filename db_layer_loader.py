# -*- coding: utf-8 -*-
"""
DB Layer Loader - Creates filtered QGIS layers from PostgreSQL for Project Mode.

In Project Mode, layers are loaded directly from the database using the SRO
derived from the project directory name, instead of requiring the user to
manually load and filter layers in QGIS.

Supported layers:
- infra_pt_pot: filtered by sro + affectation != '3'
- etude_cap_ft: filtered by sro
- etude_comac: filtered by sro
"""

from typing import Optional, List, Tuple
from qgis.core import (
    QgsVectorLayer, QgsDataSourceUri, QgsProject,
    QgsMessageLog, Qgis
)

from .db_connection import DatabaseConnection


# Default schema for all RIP AVG NGE tables
_DEFAULT_SCHEMA = 'rip_avg_nge'

# Layer definitions: (table_name, geometry_column, key_column)
_LAYER_DEFS = {
    'infra_pt_pot': ('infra_pt_pot', 'geom', 'gid'),
    'etude_cap_ft': ('etude_cap_ft', 'geom', 'gid'),
    'etude_comac':  ('etude_comac',  'geom', 'gid'),
}


def _build_pg_uri(conn_params: dict, schema: str, table: str,
                  geom_col: str, key_col: str,
                  sql_filter: str = '') -> QgsDataSourceUri:
    """Build a QgsDataSourceUri for a PostgreSQL table with optional filter."""
    uri = QgsDataSourceUri()
    uri.setConnection(
        conn_params['host'],
        str(conn_params['port']),
        conn_params['database'],
        conn_params['user'],
        conn_params['password']
    )
    uri.setDataSource(schema, table, geom_col, sql_filter, key_col)
    return uri


class DbLayerLoader:
    """Creates QGIS vector layers from PostgreSQL filtered by SRO.

    Usage:
        loader = DbLayerLoader()
        if loader.connect():
            lyr_pot = loader.load_infra_pt_pot(sro)
            lyr_cap = loader.load_etude_cap_ft(sro)
            lyr_com = loader.load_etude_comac(sro)
    """

    def __init__(self, schema: str = _DEFAULT_SCHEMA):
        self._schema = schema
        self._conn_params = None
        self._loaded_layers = []  # Track layers for cleanup

    def connect(self) -> bool:
        """Find and validate PostgreSQL connection from QGIS settings.

        Returns:
            True if connection parameters were found successfully.
        """
        db = DatabaseConnection()
        conn_name = db.find_auvergne_connection()
        if not conn_name:
            QgsMessageLog.logMessage(
                "Mode Projet: connexion PostgreSQL 'Auvergne' non trouvee",
                "PoleAerien", Qgis.Warning
            )
            return False

        self._conn_params = db.get_connection_params(conn_name)
        return True

    def _create_layer(self, table_key: str, sql_filter: str,
                      layer_name: str) -> Optional[QgsVectorLayer]:
        """Create a QgsVectorLayer from a PostgreSQL table with filter.

        Args:
            table_key: Key in _LAYER_DEFS.
            sql_filter: SQL WHERE clause for filtering.
            layer_name: Display name for the layer in QGIS.

        Returns:
            Valid QgsVectorLayer or None on failure.
        """
        if not self._conn_params:
            return None

        table_name, geom_col, key_col = _LAYER_DEFS[table_key]
        uri = _build_pg_uri(
            self._conn_params, self._schema,
            table_name, geom_col, key_col, sql_filter
        )

        layer = QgsVectorLayer(uri.uri(False), layer_name, 'postgres')
        if not layer.isValid():
            QgsMessageLog.logMessage(
                f"Mode Projet: couche '{layer_name}' invalide "
                f"(table={self._schema}.{table_name}, filtre={sql_filter})",
                "PoleAerien", Qgis.Warning
            )
            return None

        feat_count = layer.featureCount()
        QgsMessageLog.logMessage(
            f"Mode Projet: {layer_name} = {feat_count} entites "
            f"(filtre: {sql_filter})",
            "PoleAerien", Qgis.Info
        )
        return layer

    def load_infra_pt_pot(self, sro: str,
                          add_to_project: bool = False) -> Optional[QgsVectorLayer]:
        """Load infra_pt_pot filtered by SRO (excluding affectation='3').

        Filter: sro ilike '{sro}' AND affectation NOT LIKE '3'

        Args:
            sro: SRO code (e.g. '63041/B1I/PMZ/00003')
            add_to_project: If True, add layer to QGIS project.

        Returns:
            Filtered QgsVectorLayer or None.
        """
        sql_filter = f"sro ILIKE '{sro}' AND affectation NOT LIKE '3'"
        layer_name = f"infra_pt_pot [{sro}]"
        layer = self._create_layer('infra_pt_pot', sql_filter, layer_name)

        if layer and add_to_project:
            QgsProject.instance().addMapLayer(layer)
            self._loaded_layers.append(layer.id())

        return layer

    def load_etude_cap_ft(self, sro: str,
                          add_to_project: bool = False) -> Optional[QgsVectorLayer]:
        """Load etude_cap_ft filtered by SRO.

        Args:
            sro: SRO code
            add_to_project: If True, add layer to QGIS project.

        Returns:
            Filtered QgsVectorLayer or None.
        """
        sql_filter = f"sro ILIKE '{sro}'"
        layer_name = f"etude_cap_ft [{sro}]"
        layer = self._create_layer('etude_cap_ft', sql_filter, layer_name)

        if layer and add_to_project:
            QgsProject.instance().addMapLayer(layer)
            self._loaded_layers.append(layer.id())

        return layer

    def load_etude_comac(self, sro: str,
                         add_to_project: bool = False) -> Optional[QgsVectorLayer]:
        """Load etude_comac filtered by SRO.

        Args:
            sro: SRO code
            add_to_project: If True, add layer to QGIS project.

        Returns:
            Filtered QgsVectorLayer or None.
        """
        sql_filter = f"sro ILIKE '{sro}'"
        layer_name = f"etude_comac [{sro}]"
        layer = self._create_layer('etude_comac', sql_filter, layer_name)

        if layer and add_to_project:
            QgsProject.instance().addMapLayer(layer)
            self._loaded_layers.append(layer.id())

        return layer

    def load_all(self, sro: str,
                 add_to_project: bool = False
                 ) -> Tuple[Optional[QgsVectorLayer],
                            Optional[QgsVectorLayer],
                            Optional[QgsVectorLayer]]:
        """Load all three layers filtered by SRO.

        Args:
            sro: SRO code
            add_to_project: If True, add layers to QGIS project.

        Returns:
            Tuple (infra_pt_pot, etude_cap_ft, etude_comac). Any may be None.
        """
        lyr_pot = self.load_infra_pt_pot(sro, add_to_project)
        lyr_cap = self.load_etude_cap_ft(sro, add_to_project)
        lyr_com = self.load_etude_comac(sro, add_to_project)
        return lyr_pot, lyr_cap, lyr_com

    def cleanup_layers(self):
        """Remove all layers loaded by this loader from the QGIS project."""
        project = QgsProject.instance()
        for layer_id in self._loaded_layers:
            if project.mapLayer(layer_id):
                project.removeMapLayer(layer_id)
        self._loaded_layers.clear()

    @property
    def loaded_layer_ids(self) -> List[str]:
        """IDs of layers added to the QGIS project by this loader."""
        return list(self._loaded_layers)
