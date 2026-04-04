# -*- coding: utf-8 -*-
"""
Batch Data Extractor - centralises all QGIS and PostgreSQL extraction.

Single extraction pass on the main thread before any task is launched.
All modules share the same pre-extracted data via ExtractedData.

Principles:
- ExtractedData is read-only after creation (no mutation in workers)
- fddcpi2 queried once here; workers receive list[CableSegment] directly
- BPE and attaches queried once here for both COMAC and Police C6
- Each workflow skips its own extraction when ExtractedData is provided
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional
from .compat import MSG_INFO, MSG_WARNING, MSG_CRITICAL


@dataclass
class ExtractedData:
    """Immutable snapshot of all batch data extracted on the main thread."""

    # CAP_FT extraction
    doublons_capft: List = field(default_factory=list)
    hors_etude_capft: List = field(default_factory=list)
    dico_qgis_capft: Dict = field(default_factory=dict)
    dico_prives_capft: Dict = field(default_factory=dict)
    all_inf_nums_capft: Any = field(default_factory=set)
    coords_qgis_capft: Dict = field(default_factory=dict)

    # COMAC extraction
    doublons_comac: List = field(default_factory=list)
    hors_etude_comac: List = field(default_factory=list)
    dico_qgis_comac: Dict = field(default_factory=dict)
    dico_prives_comac: Dict = field(default_factory=dict)
    all_inf_nums_comac: Any = field(default_factory=set)
    coords_qgis_comac: Dict = field(default_factory=dict)

    # Appuis WKB: two variants (COMAC needs /commune, Police C6 does not)
    appuis_wkb_commune: List = field(default_factory=list)
    appuis_wkb_plain: List = field(default_factory=list)

    # C6 vs BD extraction (DataFrames)
    df_c6bd_in: Any = None
    df_c6bd_out: Any = None

    # PostgreSQL data (queried once, shared)
    cables: Optional[List] = None       # list[CableSegment] from fddcpi2 or GraceTHD
    cables_source: str = ''             # 'fddcpi2' | 'gracethd' | ''
    bpe_list: List = field(default_factory=list)
    attaches_raw: List = field(default_factory=list)

    # Metadata
    sro: str = ''
    be_type: str = 'nge'


class BatchDataExtractor:
    """Centralised QGIS + PostgreSQL extraction for batch runs.

    Called once on the main thread before runner.start().
    Delegates to existing module functions (no logic duplication).
    """

    def extract_all(
        self,
        module_keys,
        lyr_pot,
        lyr_cap,
        lyr_com,
        sro: str,
        be_type: str = 'nge',
        gracethd_dir: str = '',
        capft_col: str = '',
        comac_col: str = '',
        c6bd_col: str = '',
    ) -> ExtractedData:
        """Extract all data required by module_keys.

        Args:
            module_keys: modules that will run in this batch
            lyr_pot: QgsVectorLayer infra_pt_pot
            lyr_cap: QgsVectorLayer etude_cap_ft
            lyr_com: QgsVectorLayer etude_comac
            sro: SRO code string
            be_type: 'nge' | 'axione'
            gracethd_dir: GraceTHD directory path (Axione only)
            capft_col: field name for CAP_FT study column
            comac_col: field name for COMAC study column
            c6bd_col: field name for C6BD study column

        Returns:
            ExtractedData (read-only after return)
        """
        keys = set(module_keys)
        data = ExtractedData(sro=sro, be_type=be_type)

        needs_capft = 'capft' in keys
        needs_comac = 'comac' in keys
        needs_c6bd = bool({'c6bd', 'c6c3a', 'police_c6'} & keys)
        needs_appuis = bool({'comac', 'police_c6'} & keys)
        needs_pg = bool({'comac', 'police_c6'} & keys)

        if needs_capft and lyr_pot and lyr_cap:
            self._extract_capft(data, lyr_pot.name(), lyr_cap.name(), capft_col)

        if needs_comac and lyr_pot and lyr_com:
            self._extract_comac(data, lyr_pot.name(), lyr_com.name(), comac_col, be_type)

        if needs_c6bd and lyr_pot and lyr_cap:
            self._extract_c6bd(data, lyr_pot.name(), lyr_cap.name(), c6bd_col)

        if needs_appuis and lyr_pot and sro:
            self._extract_appuis(data, lyr_pot)

        if needs_pg and sro:
            self._extract_pg(data, sro, be_type, gracethd_dir)

        return data

    # ------------------------------------------------------------------
    #  Phase A: QGIS extractions (main thread)
    # ------------------------------------------------------------------

    def _extract_capft(self, data, pot_name, cap_name, col):
        from .CapFt import CapFt
        try:
            cap = CapFt()
            (data.doublons_capft, data.hors_etude_capft,
             data.dico_qgis_capft, data.dico_prives_capft,
             data.all_inf_nums_capft, data.coords_qgis_capft) = (
                cap.extraire_donnees_capft(pot_name, cap_name, col)
            )
        except Exception as e:
            from qgis.core import QgsMessageLog, Qgis
            QgsMessageLog.logMessage(
                f"BatchExtractor._extract_capft: {e}", "PoleAerien", MSG_WARNING
            )

    def _extract_comac(self, data, pot_name, com_name, col, be_type):
        from .Comac import Comac
        try:
            com = Comac()
            (data.doublons_comac, data.hors_etude_comac,
             data.dico_qgis_comac, data.dico_prives_comac,
             data.all_inf_nums_comac, data.coords_qgis_comac) = (
                com.extraire_donnees_comac(pot_name, com_name, col, be_type=be_type)
            )
        except Exception as e:
            from qgis.core import QgsMessageLog, Qgis
            QgsMessageLog.logMessage(
                f"BatchExtractor._extract_comac: {e}", "PoleAerien", MSG_WARNING
            )

    def _extract_c6bd(self, data, pot_name, cap_name, col):
        from .C6_vs_Bd import C6_vs_Bd
        try:
            c6bd = C6_vs_Bd()
            data.df_c6bd_in, data.df_c6bd_out = c6bd.extraire_poteaux_in_out(
                pot_name, cap_name, col
            )
        except Exception as e:
            from qgis.core import QgsMessageLog, Qgis
            QgsMessageLog.logMessage(
                f"BatchExtractor._extract_c6bd: {e}", "PoleAerien", MSG_WARNING
            )

    def _extract_appuis(self, data, lyr_pot):
        from .cable_analyzer import extraire_appuis_wkb
        try:
            data.appuis_wkb_commune = extraire_appuis_wkb(lyr_pot, keep_commune=True)
            data.appuis_wkb_plain = extraire_appuis_wkb(lyr_pot, keep_commune=False)
        except Exception as e:
            from qgis.core import QgsMessageLog, Qgis
            QgsMessageLog.logMessage(
                f"BatchExtractor._extract_appuis: {e}", "PoleAerien", MSG_WARNING
            )

    # ------------------------------------------------------------------
    #  Phase B: PostgreSQL / GraceTHD extraction (still main thread)
    # ------------------------------------------------------------------

    def _extract_pg(self, data, sro, be_type, gracethd_dir):
        if be_type == 'axione' and gracethd_dir:
            self._extract_gracethd(data, gracethd_dir)
        else:
            self._extract_fddcpi(data, sro)
            self._extract_bpe_attaches(data, sro)

    def _extract_fddcpi(self, data, sro):
        from .db_connection import get_shared_connection
        try:
            db = get_shared_connection()
            if db.connect():
                data.cables = db.execute_fddcpi2(sro)
                data.cables_source = 'fddcpi2'
                from qgis.core import QgsMessageLog, Qgis
                QgsMessageLog.logMessage(
                    f"BatchExtractor: fddcpi2({sro}) -> {len(data.cables)} segments",
                    "PoleAerien", MSG_INFO
                )
        except Exception as e:
            from qgis.core import QgsMessageLog, Qgis
            QgsMessageLog.logMessage(
                f"BatchExtractor._extract_fddcpi: {e}", "PoleAerien", MSG_WARNING
            )

    def _extract_bpe_attaches(self, data, sro):
        from .db_connection import get_shared_connection
        try:
            db = get_shared_connection()
            if db.connect():
                data.bpe_list = db.query_bpe_by_sro(sro)
                data.attaches_raw = db.query_attaches_by_sro(sro)
        except Exception as e:
            from qgis.core import QgsMessageLog, Qgis
            QgsMessageLog.logMessage(
                f"BatchExtractor._extract_bpe_attaches: {e}", "PoleAerien", MSG_WARNING
            )

    def _extract_gracethd(self, data, gracethd_dir):
        from .gracethd_reader import GraceTHDReader
        try:
            reader = GraceTHDReader(gracethd_dir)
            data.cables = reader.load_cables_as_segments('DI')
            data.cables_source = 'gracethd'
            data.bpe_list = reader.load_bpe()
        except Exception as e:
            from qgis.core import QgsMessageLog, Qgis
            QgsMessageLog.logMessage(
                f"BatchExtractor._extract_gracethd: {e}", "PoleAerien", MSG_WARNING
            )
