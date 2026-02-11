# -*- coding: utf-8 -*-
"""
Project folder auto-detection for PoleAerien plugin.

Given a project root directory (e.g. '63041-B1I-PMZ-00003'), detects:
- FT-BT KO Excel file     → MAJ module
- CAP FT sub-folder        → CAP_FT, C6 vs BD, Police C6 modules
- COMAC sub-folder          → COMAC module
- C7 / C3A Excel files     → C6-C3A-C7-BD module

Real project convention:
    <project_root>/
        FT-BT KO *.xlsx
        CAP FT/                   ← studies + C6 annexes inside
            CMD 1/                  (also: CAPFT, ETUDE CAPFT,
                FTTH-NGE-ETUDE-*/    ETUDE CAP FT, KPFT, etc.)
            CMD 2/
            ...
        COMAC/                    (also: ETUDE COMAC, Export COMAC)

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
    Examples:
        '63041-B1I-PMZ-00003'      -> '63041/B1I/PMZ/00003'
        '63471-S05-PMZ-49785_CDC'  -> '63471/S05/PMZ/49785'

    Handles:
        - Middle segments with digits (S05, B1I, A2B, etc.)
        - Suffixes after the 5-digit code (_CDC, _v2, etc.)
        - Case-insensitive matching with uppercase normalization

    Args:
        project_name: Basename of the project directory.

    Returns:
        SRO string or None if the name does not match the expected pattern.
    """
    if not project_name:
        return None
    # Search for SRO pattern anywhere in the name (ignores suffixes like _CDC)
    # Accepts both hyphens and underscores as separators
    match = re.search(
        r'(\d{5})[-_]([A-Z0-9]{3})[-_]([A-Z]{3})[-_](\d{5})',
        project_name,
        re.IGNORECASE
    )
    if match:
        return f"{match.group(1)}/{match.group(2).upper()}/{match.group(3).upper()}/{match.group(4)}"
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

    # C7
    c7_file: str = ''

    # C3A
    c3a_file: str = ''

    # Export (defaults to project root)
    export_dir: str = ''

    # Diagnostics: hints for missing resources
    # List of (resource_label, hint_message) for resources NOT found
    diagnostics: List[tuple] = field(default_factory=list)

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
            ('C7', self.c7_file, self.has_c7),
            ('C3A', self.c3a_file, self.has_c3a),
        ]

    def get_diagnostic(self, resource_label: str) -> str:
        """Get diagnostic hint for a specific missing resource."""
        for label, hint in self.diagnostics:
            if label == resource_label:
                return hint
        return ''


# --- Detection patterns (case-insensitive) ---

_FTBT_PATTERNS = [
    r'FT[\s_-]*BT[\s_-]*KO',
    r'FTBTKO',
]

_CAPFT_DIR_PATTERNS = [
    r'^CAP[\s_-]*FT$',
    r'^CAPFT$',
    r'^CAP_FT$',
    r'^ETUDE[\s_-]*CAP[\s_-]*FT$',
    r'^ETUDE[\s_-]*CAPFT$',
    r'^KPFT$',
    r'^ETUDE[\s_-]*KPFT$',
]

_COMAC_DIR_PATTERNS = [
    r'^COMAC$',
    r'^Export[\s_-]*COMAC$',
    r'^ExportComac$',
    r'^ETUDE[\s_-]*COMAC$',
]

# Human-readable expected names for diagnostics
_EXPECTED_NAMES = {
    'FT-BT KO': 'Fichier Excel contenant "FT-BT KO" ou "FTBTKO" dans le nom (a la racine du projet)',
    'CAP FT': 'Dossier nomme: CAP FT, CAPFT, CAP_FT, ETUDE CAP FT, ETUDE CAPFT, KPFT, ETUDE KPFT',
    'COMAC': 'Dossier nomme: COMAC, Export COMAC, ExportComac, ETUDE COMAC',
    'C7': 'Fichier Excel contenant "C7" ou dossier "C7" / "Annexe C7" avec un .xlsx dedans',
    'C3A': 'Fichier Excel contenant "C3A" ou dossier "C3A" / "Annexe C3A" avec un .xlsx dedans',
}


def _list_folder_contents(root: str, max_items: int = 12) -> tuple:
    """List directories and Excel files in a folder for diagnostics.

    Returns:
        (dirs: List[str], excels: List[str]) basenames only, sorted.
    """
    try:
        entries = os.listdir(root)
    except OSError:
        return [], []
    dirs = sorted(e for e in entries if os.path.isdir(os.path.join(root, e)))
    excels = sorted(e for e in entries if e.lower().endswith(('.xlsx', '.xls'))
                    and os.path.isfile(os.path.join(root, e)))
    return dirs[:max_items], excels[:max_items]


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


def _build_diagnostics(result: DetectionResult, project_root: str) -> List[tuple]:
    """Generate actionable diagnostic hints for each missing resource."""
    diags = []
    dirs, excels = _list_folder_contents(project_root)

    if not result.has_ftbt:
        hint = _EXPECTED_NAMES['FT-BT KO']
        if excels:
            hint += f"\nFichiers Excel trouves a la racine: {', '.join(excels)}"
        else:
            hint += "\nAucun fichier Excel trouve a la racine du projet"
        diags.append(('FT-BT KO', hint))

    if not result.has_capft:
        hint = _EXPECTED_NAMES['CAP FT']
        if dirs:
            hint += f"\nSous-dossiers trouves: {', '.join(dirs)}"
        else:
            hint += "\nAucun sous-dossier trouve dans le projet"
        diags.append(('CAP FT', hint))

    if not result.has_comac:
        hint = _EXPECTED_NAMES['COMAC']
        if dirs:
            hint += f"\nSous-dossiers trouves: {', '.join(dirs)}"
        else:
            hint += "\nAucun sous-dossier trouve dans le projet"
        diags.append(('COMAC', hint))

    if not result.has_c6:
        if result.has_capft:
            diags.append(('C6 (via CAP FT)', 'C6 = dossier CAP FT (deja detecte)'))
        else:
            diags.append(('C6 (via CAP FT)',
                          'Les annexes C6 sont dans le dossier CAP FT. '
                          'Detecter CAP FT pour activer C6 vs BD et Police C6.'))

    if not result.has_c7:
        hint = _EXPECTED_NAMES['C7']
        if excels:
            hint += f"\nFichiers Excel trouves: {', '.join(excels)}"
        if dirs:
            c7_like = [d for d in dirs if 'c7' in d.lower()]
            if c7_like:
                hint += f"\nDossier(s) C7-like trouves (sans Excel dedans?): {', '.join(c7_like)}"
        diags.append(('C7', hint))

    if not result.has_c3a:
        hint = _EXPECTED_NAMES['C3A']
        if excels:
            hint += f"\nFichiers Excel trouves: {', '.join(excels)}"
        if dirs:
            c3a_like = [d for d in dirs if 'c3a' in d.lower()]
            if c3a_like:
                hint += f"\nDossier(s) C3A-like trouves (sans Excel dedans?): {', '.join(c3a_like)}"
        diags.append(('C3A', hint))

    return diags


def _count_studies_recursive(capft_dir: str) -> int:
    """Count study Excel files (FTTH-NGE-ETUDE-*.xlsx) recursively in CAP FT."""
    count = 0
    for _, _, filenames in os.walk(capft_dir):
        for fn in filenames:
            if fn.lower().endswith('.xlsx') and 'ETUDE' in fn.upper():
                count += 1
    return count


def _count_excels_recursive(directory: str) -> int:
    """Count all .xlsx/.xls files recursively."""
    count = 0
    for _, _, filenames in os.walk(directory):
        count += sum(1 for f in filenames if f.lower().endswith(('.xlsx', '.xls')))
    return count


def _count_cmd_folders(capft_dir: str) -> int:
    """Count CMD sub-folders in CAP FT directory."""
    if not capft_dir or not os.path.isdir(capft_dir):
        return 0
    return sum(
        1 for e in os.listdir(capft_dir)
        if os.path.isdir(os.path.join(capft_dir, e))
        and re.match(r'^CMD', e, re.IGNORECASE)
    )


def _count_etude_folders(directory: str) -> int:
    """Count FTTH-NGE-ETUDE-* study folders recursively."""
    count = 0
    for _, dirnames, _ in os.walk(directory):
        count += sum(1 for d in dirnames if 'ETUDE' in d.upper())
    return count


def analyse_livrable(result: 'DetectionResult') -> dict:
    """Analyse approfondie du contenu livrable pour le diagnostic.

    Returns:
        dict with keys per resource, each containing detailed counts.
    """
    analysis = {}

    if result.has_capft:
        cmd = _count_cmd_folders(result.capft_dir)
        etudes = _count_etude_folders(result.capft_dir)
        excels = _count_excels_recursive(result.capft_dir)
        analysis['capft'] = {'cmd': cmd, 'etudes': etudes, 'excels': excels}

    if result.has_comac:
        excels = _count_excels_recursive(result.comac_dir)
        analysis['comac'] = {
            'sous_dossiers': len(result.comac_studies),
            'excels': excels,
        }

    if result.has_ftbt:
        analysis['ftbt'] = {'fichier': os.path.basename(result.ftbt_excel)}

    if result.has_c7:
        analysis['c7'] = {'fichier': os.path.basename(result.c7_file)}

    if result.has_c3a:
        analysis['c3a'] = {'fichier': os.path.basename(result.c3a_file)}

    return analysis


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

    # --- Generate diagnostics for missing resources ---
    result.diagnostics = _build_diagnostics(result, project_root)

    return result
