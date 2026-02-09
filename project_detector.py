# -*- coding: utf-8 -*-
"""
Project folder auto-detection for PoleAerien plugin.

Given a project root directory (e.g. '63041-B1I-PMZ-00003'), detects:
- FT-BT KO Excel file     → MAJ module
- CAP FT sub-folder        → CAP_FT, C6 vs BD, Police C6 modules
- COMAC sub-folder          → COMAC module
- BASE DE DONNEES/NGE      → GraceTHD shapefiles (Police C6)

Real project convention:
    <project_root>/
        FT-BT KO *.xlsx
        CAP FT/                   ← studies + C6 annexes inside
            CMD 1/
                FTTH-NGE-ETUDE-*/
            CMD 2/
            ...
        COMAC/
        BASE DE DONNEES/
            NGE/                  ← GraceTHD shapefiles

Note: There is NO separate C6 folder. C6 annexes are Excel files
within the study folders under CAP FT/. For Police C6 and C6 vs BD,
the CAP FT directory IS the C6 browse directory.
"""

import os
import re
from dataclasses import dataclass, field
from typing import Optional, List


def extract_sro_from_project_name(project_name: str) -> Optional[str]:
    """Derive SRO code from project directory name.

    Convention: directory uses hyphens, SRO uses slashes.
    Example: '63041-B1I-PMZ-00003' -> '63041/B1I/PMZ/00003'

    Args:
        project_name: Basename of the project directory.

    Returns:
        SRO string or None if the name does not match the expected pattern.
    """
    if not project_name:
        return None
    sro = project_name.replace('-', '/')
    if re.match(r'^\d{5}/[A-Z]\d[A-Z]/[A-Z]{3}/\d{5}$', sro):
        return sro
    return None


@dataclass
class DetectionResult:
    """Result of project folder auto-detection."""
    project_root: str = ''
    project_name: str = ''

    # SRO derived from project name
    sro: str = ''

    # MAJ FT/BT
    ftbt_excel: str = ''

    # CAP FT
    capft_dir: str = ''
    capft_studies: List[str] = field(default_factory=list)

    # COMAC
    comac_dir: str = ''
    comac_studies: List[str] = field(default_factory=list)

    # C6 browse dir (= CAP FT dir, or dedicated C6 folder)
    c6_dir: str = ''

    # GraceTHD shapefiles directory
    gracethd_dir: str = ''

    # C7
    c7_file: str = ''

    # C3A
    c3a_file: str = ''

    # Export (defaults to project root)
    export_dir: str = ''

    @property
    def has_ftbt(self) -> bool:
        return bool(self.ftbt_excel) and os.path.isfile(self.ftbt_excel)

    @property
    def has_capft(self) -> bool:
        return bool(self.capft_dir) and os.path.isdir(self.capft_dir)

    @property
    def has_comac(self) -> bool:
        return bool(self.comac_dir) and os.path.isdir(self.comac_dir)

    @property
    def has_c6(self) -> bool:
        return bool(self.c6_dir) and os.path.isdir(self.c6_dir)

    @property
    def has_gracethd(self) -> bool:
        return bool(self.gracethd_dir) and os.path.isdir(self.gracethd_dir)

    @property
    def has_c7(self) -> bool:
        return bool(self.c7_file) and os.path.isfile(self.c7_file)

    @property
    def has_c3a(self) -> bool:
        return bool(self.c3a_file) and os.path.isfile(self.c3a_file)

    def detected_modules(self) -> List[str]:
        """Return list of detected module keys."""
        modules = []
        if self.has_ftbt:
            modules.append('maj')
        if self.has_capft:
            modules.append('capft')
        if self.has_comac:
            modules.append('comac')
        if self.has_c6:
            modules.append('c6bd')
            modules.append('police_c6')
        if self.has_c6 and (self.has_c7 or self.has_c3a):
            modules.append('c6c3a')
        return modules

    def summary_lines(self) -> List[tuple]:
        """Return list of (label, path, found) for display."""
        return [
            ('FT-BT KO', self.ftbt_excel, self.has_ftbt),
            ('CAP FT', self.capft_dir, self.has_capft),
            ('COMAC', self.comac_dir, self.has_comac),
            ('C6 (via CAP FT)', self.c6_dir, self.has_c6),
            ('GraceTHD', self.gracethd_dir, self.has_gracethd),
            ('C7', self.c7_file, self.has_c7),
            ('C3A', self.c3a_file, self.has_c3a),
        ]


# --- Detection patterns (case-insensitive) ---

_FTBT_PATTERNS = [
    r'FT[\s_-]*BT[\s_-]*KO',
    r'FTBTKO',
]

_CAPFT_DIR_PATTERNS = [
    r'^CAP[\s_-]*FT$',
    r'^CAPFT$',
    r'^CAP_FT$',
]

_COMAC_DIR_PATTERNS = [
    r'^COMAC$',
    r'^Export[\s_-]*COMAC$',
    r'^ExportComac$',
]

_GRACETHD_DIR_PATTERNS = [
    r'^NGE$',
    r'^GraceTHD$',
    r'^GRACETHD$',
    r'^Grace[\s_-]*THD$',
]

_BASE_DE_DONNEES_PATTERNS = [
    r'^BASE[\s_-]*DE[\s_-]*DONN[EÉ]ES$',
    r'^BDD$',
    r'^BASE[\s_-]*DONNEES$',
]


def _match_dir(name: str, patterns: list) -> bool:
    """Check if directory name matches any pattern."""
    for pat in patterns:
        if re.match(pat, name, re.IGNORECASE):
            return True
    return False


def _find_dir(root: str, patterns: list) -> Optional[str]:
    """Find first sub-directory matching patterns."""
    try:
        entries = os.listdir(root)
    except OSError:
        return None
    for entry in sorted(entries):
        full = os.path.join(root, entry)
        if os.path.isdir(full) and _match_dir(entry, patterns):
            return full
    return None


def _find_excel(root: str, name_patterns: list) -> Optional[str]:
    """Find first Excel file matching name patterns in root."""
    try:
        entries = os.listdir(root)
    except OSError:
        return None
    xlsx_files = [e for e in entries if e.lower().endswith(('.xlsx', '.xls'))]
    for pat in name_patterns:
        regex = re.compile(pat, re.IGNORECASE)
        for f in sorted(xlsx_files):
            if regex.search(f):
                return os.path.join(root, f)
    return None


def _find_excel_in_dir(directory: str) -> Optional[str]:
    """Find first .xlsx file in a directory."""
    if not directory or not os.path.isdir(directory):
        return None
    for f in sorted(os.listdir(directory)):
        if f.lower().endswith(('.xlsx', '.xls')):
            return os.path.join(directory, f)
    return None


def _find_gracethd(project_root: str) -> Optional[str]:
    """Find GraceTHD directory (NGE shapefiles).

    Searches in:
    1. <root>/BASE DE DONNEES/NGE/
    2. <root>/NGE/
    3. <root>/GraceTHD/
    4. Any subdir containing bpe.shp and t_cheminement.shp
    """
    # Check inside BASE DE DONNEES first
    bdd = _find_dir(project_root, _BASE_DE_DONNEES_PATTERNS)
    if bdd:
        nge = _find_dir(bdd, _GRACETHD_DIR_PATTERNS)
        if nge and _is_gracethd_dir(nge):
            return nge

    # Direct subdir
    nge = _find_dir(project_root, _GRACETHD_DIR_PATTERNS)
    if nge and _is_gracethd_dir(nge):
        return nge

    return None


def _is_gracethd_dir(path: str) -> bool:
    """Verify directory contains GraceTHD shapefiles."""
    try:
        files = {f.lower() for f in os.listdir(path)}
    except OSError:
        return False
    # Must contain at least bpe.shp or t_cheminement.shp
    return 'bpe.shp' in files or 't_cheminement.shp' in files


def _list_studies(directory: str) -> List[str]:
    """List sub-directories (study names) in a directory."""
    if not directory or not os.path.isdir(directory):
        return []
    result = []
    for entry in sorted(os.listdir(directory)):
        full = os.path.join(directory, entry)
        if os.path.isdir(full):
            result.append(entry)
    return result


def _count_studies_recursive(capft_dir: str) -> int:
    """Count study Excel files (FTTH-NGE-ETUDE-*.xlsx) recursively in CAP FT."""
    count = 0
    for dirpath, _, filenames in os.walk(capft_dir):
        for fn in filenames:
            if fn.lower().endswith('.xlsx') and 'ETUDE' in fn.upper():
                count += 1
    return count


def detect_project(project_root: str) -> DetectionResult:
    """Detect project structure from root directory.

    Args:
        project_root: Absolute path to project root folder.

    Returns:
        DetectionResult with all detected paths.
    """
    result = DetectionResult()

    if not project_root or not os.path.isdir(project_root):
        return result

    result.project_root = os.path.normpath(project_root)
    result.project_name = os.path.basename(result.project_root)
    result.export_dir = result.project_root

    # Derive SRO from project name (e.g. 63041-B1I-PMZ-00003 -> 63041/B1I/PMZ/00003)
    result.sro = extract_sro_from_project_name(result.project_name) or ''

    # FT-BT KO Excel (in root)
    ftbt = _find_excel(project_root, _FTBT_PATTERNS)
    if ftbt:
        result.ftbt_excel = ftbt

    # CAP FT directory
    capft = _find_dir(project_root, _CAPFT_DIR_PATTERNS)
    if capft:
        result.capft_dir = capft
        result.capft_studies = _list_studies(capft)
        # CAP FT IS the C6 browse directory for Police C6 and C6 vs BD
        result.c6_dir = capft

    # COMAC directory
    comac = _find_dir(project_root, _COMAC_DIR_PATTERNS)
    if comac:
        result.comac_dir = comac
        result.comac_studies = _list_studies(comac)

    # GraceTHD
    gracethd = _find_gracethd(project_root)
    if gracethd:
        result.gracethd_dir = gracethd

    # C7 file (file or directory)
    c7_file = _find_excel(project_root, [r'C7', r'Annexe[\s_-]*C7'])
    if c7_file:
        result.c7_file = c7_file
    else:
        c7_dir = _find_dir(project_root, [r'^C7$', r'^Annexe[\s_-]*C7$'])
        if c7_dir:
            c7_f = _find_excel_in_dir(c7_dir)
            if c7_f:
                result.c7_file = c7_f

    # C3A file (file or directory)
    c3a_file = _find_excel(project_root, [r'C3A', r'Annexe[\s_-]*C3A'])
    if c3a_file:
        result.c3a_file = c3a_file
    else:
        c3a_dir = _find_dir(project_root, [r'^C3A$', r'^Annexe[\s_-]*C3A$'])
        if c3a_dir:
            c3a_f = _find_excel_in_dir(c3a_dir)
            if c3a_f:
                result.c3a_file = c3a_f

    return result
