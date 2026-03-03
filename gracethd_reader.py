# -*- coding: utf-8 -*-
"""
GraceTHD Reader - Chargement et jointure des fichiers GraceTHD (SHP + CSV).

Responsabilite unique: transformer un repertoire GraceTHD en structures
de donnees compatibles avec le pipeline existant (CableSegment, BPE dicts,
poteaux dicts).

Tables utilisees (sur ~30 dans le format GraceTHD):
  SHP: t_noeud, t_cableline, t_cheminement
  CSV: t_cable, t_ptech, t_ebp

Jointures:
  Cables:  t_cable.csv JOIN t_cableline.shp ON cl_cb_code = cb_code
  BPE:     t_ebp.csv JOIN t_ptech.csv ON bp_pt_code = pt_code
           JOIN t_noeud.shp ON pt_nd_code = nd_code
  Poteaux: t_ptech.csv (pt_typephy='A') JOIN t_noeud.shp ON pt_nd_code = nd_code
"""

import csv
import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from qgis.core import (
    QgsMessageLog,
    Qgis,
    QgsVectorLayer,
    QgsFeature,
    QgsGeometry,
)

from .db_connection import CableSegment


# ---------------------------------------------------------------------------
#  Constants
# ---------------------------------------------------------------------------

_REQUIRED_FILES = ['t_noeud.shp', 't_cableline.shp']
_OPTIONAL_FILES = ['t_cable.csv', 't_ptech.csv', 't_ebp.csv', 't_cheminement.shp']

_CSV_DELIMITER = ';'
_CSV_ENCODING = 'utf-8'

_LOG_TAG = 'GraceTHD'


# ---------------------------------------------------------------------------
#  Validation result
# ---------------------------------------------------------------------------

@dataclass
class GraceTHDValidation:
    """Result of GraceTHD directory validation."""
    valid: bool = True
    missing_required: List[str] = field(default_factory=list)
    missing_optional: List[str] = field(default_factory=list)
    file_counts: Dict[str, int] = field(default_factory=dict)

    @property
    def error_message(self) -> str:
        if self.valid:
            return ''
        return (
            f"GraceTHD: fichiers requis manquants: "
            f"{', '.join(self.missing_required)}"
        )


# ---------------------------------------------------------------------------
#  CSV loader (pure Python, no QGIS dependency)
# ---------------------------------------------------------------------------

def _load_csv(filepath: str) -> List[Dict[str, str]]:
    """Load a semicolon-delimited CSV into a list of dicts."""
    if not os.path.isfile(filepath):
        return []
    with open(filepath, 'r', encoding=_CSV_ENCODING, errors='replace') as f:
        reader = csv.DictReader(f, delimiter=_CSV_DELIMITER)
        return list(reader)


# ---------------------------------------------------------------------------
#  SHP loader (via QgsVectorLayer OGR provider)
# ---------------------------------------------------------------------------

def _load_shp_features(filepath: str) -> List[QgsFeature]:
    """Load all features from a shapefile. Returns empty list on failure."""
    if not os.path.isfile(filepath):
        return []
    layer = QgsVectorLayer(filepath, os.path.basename(filepath), 'ogr')
    if not layer.isValid():
        QgsMessageLog.logMessage(
            f"GraceTHD: impossible de charger {filepath}",
            _LOG_TAG, Qgis.Warning
        )
        return []
    return list(layer.getFeatures())


def _load_shp_as_dict(filepath: str, key_field: str) -> Dict[str, QgsFeature]:
    """Load SHP and index features by a unique key field."""
    features = _load_shp_features(filepath)
    result = {}
    for feat in features:
        key = str(feat[key_field]).strip() if feat[key_field] else ''
        if key:
            result[key] = feat
    return result


# ---------------------------------------------------------------------------
#  GraceTHDReader
# ---------------------------------------------------------------------------

class GraceTHDReader:
    """Reads a GraceTHD directory and produces plugin-compatible data.

    Usage:
        reader = GraceTHDReader('/path/to/GRACE_APD_DTR_03L02_00010')
        validation = reader.validate()
        if not validation.valid:
            raise ValueError(validation.error_message)
        cables = reader.load_cables_as_segments()
        bpe = reader.load_bpe()
        poteaux = reader.load_poteaux()
    """

    def __init__(self, gracethd_dir: str):
        self._dir = gracethd_dir
        self._noeuds: Optional[Dict[str, QgsFeature]] = None

    def _path(self, filename: str) -> str:
        return os.path.join(self._dir, filename)

    # ------------------------------------------------------------------
    #  Validation
    # ------------------------------------------------------------------

    def validate(self) -> GraceTHDValidation:
        """Check that required GraceTHD files exist."""
        result = GraceTHDValidation()

        for f in _REQUIRED_FILES:
            path = self._path(f)
            if os.path.isfile(path):
                result.file_counts[f] = 1
            else:
                result.missing_required.append(f)
                result.valid = False

        for f in _OPTIONAL_FILES:
            path = self._path(f)
            if os.path.isfile(path):
                result.file_counts[f] = 1
            else:
                result.missing_optional.append(f)

        return result

    # ------------------------------------------------------------------
    #  Noeuds (lazy-loaded, shared across load methods)
    # ------------------------------------------------------------------

    def _ensure_noeuds(self) -> Dict[str, QgsFeature]:
        """Load t_noeud.shp once, index by nd_code."""
        if self._noeuds is None:
            self._noeuds = _load_shp_as_dict(
                self._path('t_noeud.shp'), 'nd_code'
            )
            QgsMessageLog.logMessage(
                f"t_noeud: {len(self._noeuds)} noeuds charges",
                _LOG_TAG, Qgis.Info
            )
        return self._noeuds

    def _noeud_geom_wkt(self, nd_code: str) -> str:
        """Get WKT geometry for a noeud code, or '' if not found."""
        noeuds = self._ensure_noeuds()
        feat = noeuds.get(nd_code)
        if feat and feat.hasGeometry():
            return feat.geometry().asWkt()
        return ''

    def _noeud_geom(self, nd_code: str) -> Optional[QgsGeometry]:
        """Get QgsGeometry for a noeud code."""
        noeuds = self._ensure_noeuds()
        feat = noeuds.get(nd_code)
        if feat and feat.hasGeometry():
            return QgsGeometry(feat.geometry())
        return None

    # ------------------------------------------------------------------
    #  Cables  (P1-2)
    # ------------------------------------------------------------------

    def load_cables_as_segments(self, typelog_filter: str = 'DI') -> List[CableSegment]:
        """Load cables from GraceTHD and produce CableSegment objects.

        Join: t_cable.csv JOIN t_cableline.shp ON cl_cb_code = cb_code
        Filter: cb_typelog = typelog_filter (default 'DI' = distribution)

        Returns:
            List of CableSegment compatible with cable_analyzer pipeline.
        """
        # 1. Load t_cableline.shp indexed by cl_cb_code
        cableline_feats = _load_shp_as_dict(
            self._path('t_cableline.shp'), 'cl_cb_code'
        )

        # 2. Load t_cable.csv
        cables_csv = _load_csv(self._path('t_cable.csv'))

        # 3. Filter by typelog
        if typelog_filter:
            cables_csv = [
                r for r in cables_csv
                if r.get('cb_typelog', '').upper() == typelog_filter.upper()
            ]

        # 4. Join and produce CableSegment
        segments = []
        missing_geom = 0

        for idx, row in enumerate(cables_csv):
            cb_code = row.get('cb_code', '')
            cl_feat = cableline_feats.get(cb_code)

            if not cl_feat or not cl_feat.hasGeometry():
                missing_geom += 1
                continue

            geom = cl_feat.geometry()
            geom_wkt = geom.asWkt()
            length = geom.length() if geom else 0.0

            # cl_long from cableline (may be more accurate)
            cl_long = cl_feat['cl_long']
            if cl_long and float(cl_long) > 0:
                length = float(cl_long)

            cab_capa = 0
            raw_capa = row.get('cb_capafo', '0')
            try:
                cab_capa = int(raw_capa)
            except (ValueError, TypeError):
                pass

            segment = CableSegment(
                gid_dc2=idx + 1,
                gid_dc=idx + 1,
                gid=idx + 1,
                sro='',
                nro='',
                length=length,
                cab_type='CDI' if typelog_filter == 'DI' else row.get('cb_typelog', ''),
                cab_capa=cab_capa,
                cab_modulo=int(row.get('cb_modulo', '0') or '0'),
                isole='',
                date_modif='',
                modif_par='',
                cab_nature='',
                commentaire='',
                collecte='',
                cb_etiquet=row.get('cb_etiquet', ''),
                fon='',
                projet='',
                dce='',
                dist_type='',
                affectation='',
                posemode=1,  # default aerien, refined below
                geom_wkt=geom_wkt,
            )
            segments.append(segment)

        if missing_geom:
            QgsMessageLog.logMessage(
                f"t_cable: {missing_geom} cables sans geometrie (jointure t_cableline echouee)",
                _LOG_TAG, Qgis.Warning
            )

        # Resume detaille
        capas = {}
        for s in segments:
            capas[s.cab_capa] = capas.get(s.cab_capa, 0) + 1
        capas_str = ', '.join(f'{c}FO x{n}' for c, n in sorted(capas.items(), reverse=True))
        QgsMessageLog.logMessage(
            f"GraceTHD cables: {len(segments)} segments {typelog_filter} charges "
            f"({len(cables_csv)} filtres, {missing_geom} sans geom) | "
            f"Capacites: [{capas_str}] | "
            f"cab_type={'CDI' if typelog_filter == 'DI' else '?'}, posemode=1 (aerien par defaut)",
            _LOG_TAG, Qgis.Info
        )

        return segments

    # ------------------------------------------------------------------
    #  BPE  (P1-3)
    # ------------------------------------------------------------------

    def load_bpe(self, typelog_filter: Tuple[str, ...] = ('PBO', 'BPE')) -> List[Dict]:
        """Load BPE from GraceTHD.

        Join: t_ebp.csv JOIN t_ptech.csv ON bp_pt_code = pt_code
              JOIN t_noeud.shp ON pt_nd_code = nd_code
        Filter: bp_typelog in typelog_filter

        Returns:
            List of dicts {gid, noe_type, noe_usage, inf_num, geom_wkt}
            compatible with db_connection.query_bpe_by_sro() output.
        """
        # 1. Load t_ptech.csv indexed by pt_code
        ptech_csv = _load_csv(self._path('t_ptech.csv'))
        ptech_by_code = {r['pt_code']: r for r in ptech_csv if r.get('pt_code')}

        # 2. Load t_ebp.csv
        ebp_csv = _load_csv(self._path('t_ebp.csv'))

        # 3. Filter by typelog
        if typelog_filter:
            ebp_csv = [
                r for r in ebp_csv
                if r.get('bp_typelog', '').upper() in
                   tuple(t.upper() for t in typelog_filter)
            ]

        # 4. Ensure noeuds loaded
        self._ensure_noeuds()

        # 5. Join and produce BPE dicts
        results = []
        no_geom = 0

        for idx, row in enumerate(ebp_csv):
            bp_pt_code = row.get('bp_pt_code', '').strip()
            bp_typelog = row.get('bp_typelog', '')
            bp_typephy = row.get('bp_typephy', '')

            # Find geometry via ptech -> noeud
            geom_wkt = ''
            inf_num = ''

            if bp_pt_code and bp_pt_code in ptech_by_code:
                pt = ptech_by_code[bp_pt_code]
                nd_code = pt.get('pt_nd_code', '')
                inf_num = pt.get('pt_etiquet', '')
                geom_wkt = self._noeud_geom_wkt(nd_code)

            if not geom_wkt:
                no_geom += 1
                continue

            results.append({
                'gid': idx + 1,
                'noe_type': bp_typelog,
                'noe_usage': bp_typephy,
                'inf_num': inf_num,
                'geom_wkt': geom_wkt,
            })

        if no_geom:
            QgsMessageLog.logMessage(
                f"t_ebp: {no_geom} BPE sans geometrie (jointure ptech/noeud echouee)",
                _LOG_TAG, Qgis.Warning
            )

        QgsMessageLog.logMessage(
            f"GraceTHD BPE: {len(results)} boitiers charges "
            f"({no_geom} sans geom exclus)",
            _LOG_TAG, Qgis.Info
        )

        return results

    # ------------------------------------------------------------------
    #  Poteaux  (P1-4)
    # ------------------------------------------------------------------

    def load_poteaux(self) -> List[Dict]:
        """Load poteaux (appuis) from GraceTHD.

        Source: t_ptech.csv WHERE pt_typephy = 'A'
        Geometry: t_noeud.shp via pt_nd_code = nd_code

        Returns:
            List of dicts {pt_code, inf_num, nd_code, nature, geom_wkt, geom}
        """
        ptech_csv = _load_csv(self._path('t_ptech.csv'))
        self._ensure_noeuds()

        poteaux = []
        no_geom = 0

        for row in ptech_csv:
            if row.get('pt_typephy', '').upper() != 'A':
                continue

            nd_code = row.get('pt_nd_code', '')
            geom = self._noeud_geom(nd_code)

            if not geom:
                no_geom += 1
                continue

            poteaux.append({
                'pt_code': row.get('pt_code', ''),
                'inf_num': row.get('pt_etiquet', ''),
                'nd_code': nd_code,
                'nature': row.get('pt_nature', ''),
                'geom_wkt': geom.asWkt(),
                'geom': geom,
            })

        if no_geom:
            QgsMessageLog.logMessage(
                f"t_ptech: {no_geom} poteaux sans geometrie (noeud manquant)",
                _LOG_TAG, Qgis.Warning
            )

        QgsMessageLog.logMessage(
            f"GraceTHD poteaux: {len(poteaux)} appuis charges "
            f"({no_geom} sans geom exclus)",
            _LOG_TAG, Qgis.Info
        )

        return poteaux

    # ------------------------------------------------------------------
    #  Cheminements (bonus - pour deduire posemode)
    # ------------------------------------------------------------------

    def load_cheminements(self) -> List[Dict]:
        """Load cheminements from GraceTHD.

        Returns:
            List of dicts {cm_code, cm_ndcode1, cm_ndcode2, cm_typ_imp, cm_long, geom_wkt}
        """
        features = _load_shp_features(self._path('t_cheminement.shp'))
        results = []
        for feat in features:
            results.append({
                'cm_code': str(feat['cm_code']) if feat['cm_code'] else '',
                'cm_ndcode1': str(feat['cm_ndcode1']) if feat['cm_ndcode1'] else '',
                'cm_ndcode2': str(feat['cm_ndcode2']) if feat['cm_ndcode2'] else '',
                'cm_typ_imp': str(feat['cm_typ_imp']) if feat['cm_typ_imp'] else '',
                'cm_long': float(feat['cm_long']) if feat['cm_long'] else 0.0,
                'geom_wkt': feat.geometry().asWkt() if feat.hasGeometry() else '',
            })

        QgsMessageLog.logMessage(
            f"GraceTHD cheminements: {len(results)} troncons charges",
            _LOG_TAG, Qgis.Info
        )
        return results

    # ------------------------------------------------------------------
    #  Summary
    # ------------------------------------------------------------------

    def summary(self) -> Dict[str, int]:
        """Return file presence summary for diagnostics."""
        counts = {}
        for f in _REQUIRED_FILES + _OPTIONAL_FILES:
            path = self._path(f)
            if os.path.isfile(path):
                counts[f] = 1
            else:
                counts[f] = 0
        return counts

    def inventory(self) -> Dict:
        """Scan the GraceTHD directory and produce a structured inventory.

        Returns:
            Dict with keys:
                dir_name: basename of gracethd directory
                total_files: total file count
                shp_files: list of .shp basenames
                csv_files: list of .csv basenames
                shp_count: number of shapefiles
                csv_count: number of CSV files
                other_count: number of support files (.dbf, .prj, .cpg, .shx, .sbn, .sbx)
                total_size_mb: total size in MB
                key_tables: dict of key table -> row count (from CSV) or 'present'/'absent'
                missing_required: list of missing required files
                missing_optional: list of missing optional files
        """
        result = {
            'dir_name': os.path.basename(self._dir),
            'dir_path': self._dir,
            'total_files': 0,
            'shp_files': [],
            'csv_files': [],
            'shp_count': 0,
            'csv_count': 0,
            'other_count': 0,
            'total_size_mb': 0.0,
            'key_tables': {},
            'missing_required': [],
            'missing_optional': [],
        }

        if not os.path.isdir(self._dir):
            return result

        total_size = 0
        shp_ext = {'.shp'}
        csv_ext = {'.csv'}
        support_ext = {'.dbf', '.prj', '.cpg', '.shx', '.sbn', '.sbx'}

        for entry in os.listdir(self._dir):
            fpath = os.path.join(self._dir, entry)
            if not os.path.isfile(fpath):
                continue
            result['total_files'] += 1
            total_size += os.path.getsize(fpath)

            ext = os.path.splitext(entry)[1].lower()
            if ext in shp_ext:
                result['shp_files'].append(entry)
                result['shp_count'] += 1
            elif ext in csv_ext:
                result['csv_files'].append(entry)
                result['csv_count'] += 1
            elif ext in support_ext:
                result['other_count'] += 1

        result['total_size_mb'] = round(total_size / (1024 * 1024), 1)

        # Key tables: count rows for CSV, check presence for SHP
        key_csv_tables = {
            't_cable.csv': 'cb_code',
            't_ptech.csv': 'pt_code',
            't_ebp.csv': 'bp_code',
        }
        for fname, _id_col in key_csv_tables.items():
            fpath = self._path(fname)
            if os.path.isfile(fpath):
                rows = _load_csv(fpath)
                result['key_tables'][fname] = len(rows)
            else:
                result['key_tables'][fname] = -1  # absent

        key_shp_tables = ['t_noeud.shp', 't_cableline.shp', 't_cheminement.shp']
        for fname in key_shp_tables:
            fpath = self._path(fname)
            result['key_tables'][fname] = 'present' if os.path.isfile(fpath) else 'absent'

        # Missing files
        for f in _REQUIRED_FILES:
            if not os.path.isfile(self._path(f)):
                result['missing_required'].append(f)
        for f in _OPTIONAL_FILES:
            if not os.path.isfile(self._path(f)):
                result['missing_optional'].append(f)

        return result
