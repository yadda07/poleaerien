# -*- coding: utf-8 -*-
"""
PoleAerien Dialog V2 - Streamlined batch interface.

Pure Python, no .ui file. Native QGIS 3.28+ compatible.
Uses QgsApplication.getThemeIcon() for all icons.
Single project folder -> auto-detect -> select modules -> execute all.
"""

import os

from qgis.PyQt.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QPushButton, QLineEdit, QCheckBox, QFrame,
    QTextBrowser, QProgressBar, QFileDialog,
    QWidget, QScrollArea, QSplitter,
    QRadioButton, QButtonGroup,
)
from qgis.PyQt.QtCore import Qt, QSize, pyqtSignal
from qgis.PyQt.QtGui import QTextCursor, QFont
from qgis.core import QgsApplication, QgsProject, QgsMapLayerProxyModel
from qgis.gui import QgsMapLayerComboBox, QgsCollapsibleGroupBox

from .project_detector import detect_project, DetectionResult, analyse_livrable
from .batch_runner import MODULE_REGISTRY
from .async_tasks import SmoothProgressController


PLUGIN_DIR = os.path.dirname(__file__)


def _qgs_icon(name):
    """Get a native QGIS theme icon by name."""
    return QgsApplication.getThemeIcon(name)


def _plugin_icon(name):
    """Load SVG icon from plugin images folder (fallback)."""
    path = os.path.join(PLUGIN_DIR, 'images', name)
    if os.path.exists(path):
        return QIcon(path)
    return QIcon()


# ---------------------------------------------------------------------------
#  Stylesheet — professional, compact, native-feeling
# ---------------------------------------------------------------------------

_GLOBAL_SS = """
QDialog#PoleAerienBatch {
    background-color: palette(window);
}

/* --- Groupboxes: native look with subtle refinement --- */
QGroupBox {
    font-weight: 600;
    border: 1px solid palette(mid);
    border-radius: 3px;
    margin-top: 1.2em;
    padding-top: 0.6em;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 8px;
    padding: 0 4px;
}

/* --- Buttons --- */
QPushButton#btnStart {
    background-color: #2e7d32;
    color: white;
    border: none;
    border-radius: 3px;
    padding: 7px 22px;
    font-weight: bold;
    font-size: 10pt;
}
QPushButton#btnStart:hover {
    background-color: #1b5e20;
}
QPushButton#btnStart:pressed {
    background-color: #174f1a;
}
QPushButton#btnStart:disabled {
    background-color: palette(mid);
    color: palette(midlight);
}

QPushButton#btnCancel {
    background-color: #c62828;
    color: white;
    border: none;
    border-radius: 3px;
    padding: 7px 22px;
    font-weight: bold;
    font-size: 10pt;
}
QPushButton#btnCancel:hover {
    background-color: #b71c1c;
}

QPushButton#btnSecondary {
    border: 1px solid palette(mid);
    border-radius: 3px;
    padding: 4px 12px;
    background: palette(button);
}
QPushButton#btnSecondary:hover {
    background: palette(midlight);
    border-color: palette(dark);
}

QPushButton#btnBrowse {
    border: 1px solid palette(mid);
    border-radius: 2px;
    padding: 3px;
    background: palette(button);
    min-width: 28px;
    max-width: 28px;
}
QPushButton#btnBrowse:hover {
    background: palette(midlight);
    border-color: palette(highlight);
}

/* --- Path inputs --- */
QLineEdit#pathEdit {
    border: 1px solid palette(mid);
    border-radius: 2px;
    padding: 4px 6px;
}
QLineEdit#pathEdit:focus {
    border-color: palette(highlight);
}
QLineEdit#pathEdit:read-only {
    background: palette(window);
    color: palette(mid);
}

/* --- Progress bar --- */
QProgressBar#mainProgress {
    border: 1px solid palette(mid);
    border-radius: 2px;
    text-align: center;
    height: 16px;
    font-size: 8pt;
}
QProgressBar#mainProgress::chunk {
    background-color: #2e7d32;
    border-radius: 1px;
}

/* --- Log area --- */
QTextBrowser#logBrowser {
    border: 1px solid palette(mid);
    border-radius: 2px;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 8.5pt;
}

/* --- Detection tree --- */
QTreeWidget#detectionTree {
    border: 1px solid palette(mid);
    border-radius: 2px;
    alternate-background-color: palette(alternateBase);
    font-size: 9pt;
}
QTreeWidget#detectionTree::item {
    padding: 3px 0;
}
QTreeWidget#detectionTree::item:hover {
    background: palette(midlight);
}

/* --- Module checkboxes --- */
QCheckBox#moduleCheck {
    spacing: 6px;
    font-size: 9pt;
}
QCheckBox#moduleCheck:disabled {
    color: palette(mid);
}
"""


# ---------------------------------------------------------------------------
#  Module row widget
# ---------------------------------------------------------------------------
class _ModuleRow(QWidget):
    """Compact module row: [checkbox] [status_icon] [detected_path]."""

    toggled = pyqtSignal(str, bool)

    def __init__(self, key, label, parent=None, tooltip=''):
        super().__init__(parent)
        self.key = key
        self._found = False
        self._prereq_tooltip = tooltip

        lay = QHBoxLayout(self)
        lay.setContentsMargins(4, 2, 4, 2)
        lay.setSpacing(6)

        self.checkbox = QCheckBox(label)
        self.checkbox.setObjectName("moduleCheck")
        self.checkbox.setEnabled(False)
        if tooltip:
            self.checkbox.setToolTip(tooltip)
        self.checkbox.toggled.connect(lambda c: self.toggled.emit(self.key, c))
        lay.addWidget(self.checkbox, 1)

        self.status_icon = QLabel()
        self.status_icon.setFixedWidth(28)
        self.status_icon.setToolTip("Non detecte")
        lay.addWidget(self.status_icon)

        self.path_label = QLabel()
        self.path_label.setStyleSheet("color: palette(mid); font-size: 8pt;")
        self.path_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.path_label.setMinimumWidth(80)
        self.path_label.setMaximumWidth(260)
        lay.addWidget(self.path_label)

        self._update_status_icon(False)

    def set_detected(self, found, path='', diagnostic=''):
        self._found = found
        self.checkbox.setEnabled(found)
        self.checkbox.setChecked(found)
        self._update_status_icon(found, diagnostic)

        if found:
            display = os.path.basename(path) if path else ''
            self.path_label.setText(display)
            self.path_label.setToolTip(path)
        else:
            self.path_label.setText("")
            self.path_label.setToolTip("")

    def _update_status_icon(self, found, diagnostic=''):
        self.status_icon.clear()
        if found:
            self.status_icon.setText("[OK]")
            self.status_icon.setStyleSheet("color: #2e7d32; font-size: 8pt; font-weight: bold;")
            self.status_icon.setToolTip("Detecte")
        else:
            self.status_icon.setText("[--]")
            self.status_icon.setStyleSheet("color: #c62828; font-size: 8pt; font-weight: bold;")
            tip = "Non detecte"
            if diagnostic:
                tip += "\n" + diagnostic
            elif self._prereq_tooltip:
                tip += "\n" + self._prereq_tooltip
            self.status_icon.setToolTip(tip)

    @property
    def is_selected(self):
        return self.checkbox.isChecked() and self.checkbox.isEnabled()


# ---------------------------------------------------------------------------
#  Main Dialog
# ---------------------------------------------------------------------------
class PoleAerienDialogV2(QDialog):
    """Streamlined batch analysis dialog — professional native QGIS look."""

    start_requested = pyqtSignal(list)
    cancel_requested = pyqtSignal()
    help_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("PoleAerienBatch")
        self._detection = DetectionResult()
        self._is_running = False
        self._module_rows = {}
        self._project_mode = False
        self.smooth_progress = SmoothProgressController(interval_ms=25)

        self._build_ui()
        self._connect_signals()

    # ==================================================================
    #  UI CONSTRUCTION
    # ==================================================================

    def _build_ui(self):
        self.setWindowTitle("Pole Aerien - Analyse par lot")
        self.setWindowIcon(_qgs_icon('mActionStart.svg'))
        self.setMinimumSize(700, 750)
        self.resize(740, 850)
        self.setStyleSheet(_GLOBAL_SS)

        root = QVBoxLayout(self)
        root.setSpacing(6)
        root.setContentsMargins(8, 8, 8, 8)

        # --- Header bar ---
        root.addLayout(self._build_header())

        # --- Main splitter: config (top) / log (bottom) ---
        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(2)

        # Top: config (no scroll)
        top_widget = QWidget()
        top_lay = QVBoxLayout(top_widget)
        top_lay.setSpacing(4)
        top_lay.setContentsMargins(0, 0, 0, 0)

        self._build_project_group(top_lay)
        self._build_mode_group(top_lay)
        self._build_layers_group(top_lay)
        self._build_modules_group(top_lay)
        splitter.addWidget(top_widget)

        # Bottom: log
        log_widget = self._build_log()
        splitter.addWidget(log_widget)

        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        root.addWidget(splitter, 1)

        # --- Footer ---
        root.addLayout(self._build_footer())

    # ---- Header ----
    def _build_header(self):
        lay = QHBoxLayout()
        lay.setSpacing(8)

        icon_lbl = QLabel()
        icon_lbl.setPixmap(_qgs_icon('mActionStart.svg').pixmap(22, 22))
        lay.addWidget(icon_lbl)

        title = QLabel("Pôle Aérien")
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        title.setFont(title_font)
        lay.addWidget(title)

        self._subtitle = QLabel("  -  Selectionnez un dossier projet")
        self._subtitle.setStyleSheet("color: palette(mid); font-size: 9pt;")
        lay.addWidget(self._subtitle)

        lay.addStretch()

        self._project_tag = QLabel()
        self._project_tag.setVisible(False)
        self._project_tag.setStyleSheet("""
            QLabel {
                background: #37474f;
                color: #eceff1;
                padding: 2px 10px;
                border-radius: 2px;
                font-weight: bold;
                font-size: 9pt;
            }
        """)
        lay.addWidget(self._project_tag)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)

        outer = QVBoxLayout()
        outer.setSpacing(4)
        outer.addLayout(lay)
        outer.addWidget(sep)
        return outer

    # ---- Project group ----
    def _build_project_group(self, parent):
        grp = QgsCollapsibleGroupBox("Dossier projet")
        lay = QVBoxLayout()
        lay.setSpacing(4)
        lay.setContentsMargins(6, 4, 6, 4)

        # Project path
        row1 = QHBoxLayout()
        row1.setSpacing(4)

        lbl = QLabel("Projet :")
        lbl.setFixedWidth(52)
        lbl.setStyleSheet("font-weight: 600; font-size: 9pt;")
        row1.addWidget(lbl)

        self.projectPathEdit = QLineEdit()
        self.projectPathEdit.setObjectName("pathEdit")
        self.projectPathEdit.setPlaceholderText("Chemin du dossier projet...")
        row1.addWidget(self.projectPathEdit)

        self.browseProjectBtn = QPushButton()
        self.browseProjectBtn.setObjectName("btnBrowse")
        self.browseProjectBtn.setIcon(_qgs_icon('mActionFileOpen.svg'))
        self.browseProjectBtn.setIconSize(QSize(16, 16))
        self.browseProjectBtn.setToolTip("Parcourir...")
        self.browseProjectBtn.setCursor(Qt.PointingHandCursor)
        row1.addWidget(self.browseProjectBtn)

        self.refreshBtn = QPushButton()
        self.refreshBtn.setObjectName("btnBrowse")
        self.refreshBtn.setIcon(_qgs_icon('mActionRefresh.svg'))
        self.refreshBtn.setIconSize(QSize(16, 16))
        self.refreshBtn.setToolTip("Re-detecter les dossiers (synchroniser)")
        self.refreshBtn.setCursor(Qt.PointingHandCursor)
        row1.addWidget(self.refreshBtn)

        lay.addLayout(row1)

        # Export path
        row2 = QHBoxLayout()
        row2.setSpacing(4)

        lbl2 = QLabel("Export :")
        lbl2.setFixedWidth(52)
        lbl2.setStyleSheet("font-weight: 600; font-size: 9pt;")
        row2.addWidget(lbl2)

        self.exportPathEdit = QLineEdit()
        self.exportPathEdit.setObjectName("pathEdit")
        self.exportPathEdit.setPlaceholderText("Dossier d'export (défaut = projet)")
        row2.addWidget(self.exportPathEdit)

        self.browseExportBtn = QPushButton()
        self.browseExportBtn.setObjectName("btnBrowse")
        self.browseExportBtn.setIcon(_qgs_icon('mActionFileSaveAs.svg'))
        self.browseExportBtn.setIconSize(QSize(16, 16))
        self.browseExportBtn.setToolTip("Changer le dossier d'export")
        self.browseExportBtn.setCursor(Qt.PointingHandCursor)
        row2.addWidget(self.browseExportBtn)

        lay.addLayout(row2)

        # Detection summary label
        self._det_summary = QLabel()
        self._det_summary.setWordWrap(True)
        self._det_summary.setStyleSheet("color: palette(mid); font-size: 8pt; padding: 2px 0;")
        self._det_summary.setVisible(False)
        lay.addWidget(self._det_summary)

        grp.setLayout(lay)
        parent.addWidget(grp)

    # ---- Mode group ----
    def _build_mode_group(self, parent):
        grp = QgsCollapsibleGroupBox("Mode d'execution")
        lay = QVBoxLayout()
        lay.setSpacing(4)
        lay.setContentsMargins(6, 4, 6, 4)

        # Radio buttons
        radio_row = QHBoxLayout()
        radio_row.setSpacing(12)

        self._mode_group = QButtonGroup(self)
        self._radio_qgis = QRadioButton("Couches QGIS")
        self._radio_qgis.setToolTip(
            "Utiliser les couches deja chargees dans le projet QGIS"
        )
        self._radio_project = QRadioButton("Mode Projet (BDD)")
        self._radio_project.setToolTip(
            "Charger automatiquement les couches depuis PostgreSQL "
            "en filtrant par le SRO du projet"
        )
        self._mode_group.addButton(self._radio_qgis, 0)
        self._mode_group.addButton(self._radio_project, 1)
        self._radio_qgis.setChecked(True)

        radio_row.addWidget(self._radio_qgis)
        radio_row.addWidget(self._radio_project)
        radio_row.addStretch()
        lay.addLayout(radio_row)

        # SRO info row (visible only in project mode)
        self._sro_row = QWidget()
        sro_lay = QHBoxLayout(self._sro_row)
        sro_lay.setContentsMargins(0, 2, 0, 0)
        sro_lay.setSpacing(6)

        sro_lbl = QLabel("SRO :")
        sro_lbl.setFixedWidth(52)
        sro_lbl.setStyleSheet("font-weight: 600; font-size: 9pt;")
        sro_lay.addWidget(sro_lbl)

        self._sro_value = QLabel("-")
        self._sro_value.setStyleSheet(
            "font-family: 'Consolas', monospace; font-size: 9pt; "
            "color: #1565c0; font-weight: bold;"
        )
        sro_lay.addWidget(self._sro_value)
        sro_lay.addStretch()

        self._sro_status = QLabel()
        self._sro_status.setStyleSheet("font-size: 8pt;")
        sro_lay.addWidget(self._sro_status)

        self._sro_row.setVisible(False)
        lay.addWidget(self._sro_row)

        # Load layers checkbox (visible only in project mode)
        self._chk_load_layers = QCheckBox(
            "Charger les couches dans le projet QGIS"
        )
        self._chk_load_layers.setToolTip(
            "Si coche, les couches filtrees par SRO seront ajoutees au projet QGIS.\n"
            "Sinon, elles seront utilisees temporairement pour l'analyse uniquement."
        )
        self._chk_load_layers.setChecked(False)
        self._chk_load_layers.setVisible(False)
        lay.addWidget(self._chk_load_layers)

        grp.setLayout(lay)
        parent.addWidget(grp)

    # ---- Layers group ----
    def _build_layers_group(self, parent):
        self._layers_grp = QgsCollapsibleGroupBox("Couches QGIS")
        grp = self._layers_grp
        form = QFormLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(4)
        form.setContentsMargins(6, 4, 6, 4)

        self.comboInfraPtPot = QgsMapLayerComboBox()
        self.comboInfraPtPot.setFilters(QgsMapLayerProxyModel.PointLayer)
        form.addRow("infra_pt_pot :", self.comboInfraPtPot)

        self.comboEtudeCapFt = QgsMapLayerComboBox()
        self.comboEtudeCapFt.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        form.addRow("etude_cap_ft :", self.comboEtudeCapFt)

        self.comboEtudeComac = QgsMapLayerComboBox()
        self.comboEtudeComac.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        form.addRow("etude_comac :", self.comboEtudeComac)

        grp.setLayout(form)
        parent.addWidget(grp)

    # ---- Modules group ----
    def _build_modules_group(self, parent):
        grp = QgsCollapsibleGroupBox("Modules d'analyse")
        lay = QVBoxLayout()
        lay.setSpacing(1)
        lay.setContentsMargins(6, 4, 6, 4)

        # Actions row
        act_row = QHBoxLayout()
        act_row.setSpacing(6)

        self.selectAllBtn = QPushButton("Tout cocher")
        self.selectAllBtn.setObjectName("btnSecondary")
        self.selectAllBtn.setIcon(_qgs_icon('mActionSelectAll.svg'))
        self.selectAllBtn.setIconSize(QSize(14, 14))
        self.selectAllBtn.setCursor(Qt.PointingHandCursor)
        act_row.addWidget(self.selectAllBtn)

        self.deselectAllBtn = QPushButton("Tout décocher")
        self.deselectAllBtn.setObjectName("btnSecondary")
        self.deselectAllBtn.setIcon(_qgs_icon('mActionDeselectAll.svg'))
        self.deselectAllBtn.setIconSize(QSize(14, 14))
        self.deselectAllBtn.setCursor(Qt.PointingHandCursor)
        act_row.addWidget(self.deselectAllBtn)

        act_row.addStretch()

        self._mod_count = QLabel("0/6")
        self._mod_count.setStyleSheet("color: palette(mid); font-size: 8.5pt;")
        act_row.addWidget(self._mod_count)

        lay.addLayout(act_row)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)
        lay.addWidget(sep)

        # Module rows with prerequisite tooltips
        modules_def = [
            ('maj',       'MAJ BD - Import FT-BT KO',
             "Prerequis: fichier FT-BT KO (.xlsx) a la racine\n"
             "Couches: infra_pt_pot, etude_cap_ft, etude_comac"),
            ('capft',     'VERIF CAP_FT - Poteaux vs fiches appuis',
             "Prerequis: dossier CAP FT avec fiches appuis\n"
             "Couches: infra_pt_pot, etude_cap_ft"),
            ('comac',     'VERIF COMAC - Poteaux BT vs ExportComac',
             "Prerequis: dossier COMAC (ExportComac .xlsx)\n"
             "Couches: infra_pt_pot, etude_comac\n"
             "Connexion PostgreSQL (verif cables)"),
            ('c6bd',      'C6 vs BD - Annexe C6 vs base de donnees',
             "Prerequis: dossier CAP FT (annexes C6 = etudes .xlsx)\n"
             "Couches: infra_pt_pot, etude_cap_ft"),
            ('police_c6', 'POLICE C6 - Analyse complete',
             "Prerequis: dossier CAP FT (annexes C6)\n"
             "Couches: infra_pt_pot\n"
             "Connexion PostgreSQL (fddcpi2)"),
            ('c6c3a',     'C6-C3A-C7-BD - Croisement annexes',
             "Prerequis: dossier CAP FT + fichier C7 et/ou C3A\n"
             "Couches: infra_pt_pot, etude_cap_ft"),
        ]
        for key, label, tooltip in modules_def:
            row = _ModuleRow(key, label, tooltip=tooltip)
            row.toggled.connect(self._on_module_toggled)
            self._module_rows[key] = row
            lay.addWidget(row)

        grp.setLayout(lay)
        parent.addWidget(grp)

    # ---- Progress ----
    def _build_progress(self, parent):
        self.progressBar = QProgressBar()
        self.progressBar.setObjectName("mainProgress")
        self.progressBar.setRange(0, 100)
        self.progressBar.setValue(0)
        self.progressBar.setTextVisible(True)
        self.progressBar.setFormat("%p%")
        self.progressBar.setMinimumHeight(20)
        parent.addWidget(self.progressBar)
        self.smooth_progress.set_progress_bar(self.progressBar)

    # ---- Log ----
    def _build_log(self):
        widget = QWidget()
        lay = QVBoxLayout(widget)
        lay.setSpacing(3)
        lay.setContentsMargins(0, 0, 0, 0)

        # Header
        hdr = QHBoxLayout()
        hdr.setSpacing(6)

        lbl = QLabel("Journal")
        lbl.setStyleSheet("font-weight: 600; font-size: 9pt;")
        hdr.addWidget(lbl)
        hdr.addStretch()

        self.clearLogBtn = QPushButton("Effacer")
        self.clearLogBtn.setObjectName("btnSecondary")
        self.clearLogBtn.setIcon(_qgs_icon('mActionDeleteSelected.svg'))
        self.clearLogBtn.setIconSize(QSize(12, 12))
        self.clearLogBtn.setCursor(Qt.PointingHandCursor)
        hdr.addWidget(self.clearLogBtn)

        self.exportLogBtn = QPushButton("Exporter")
        self.exportLogBtn.setObjectName("btnSecondary")
        self.exportLogBtn.setIcon(_qgs_icon('mActionFileSaveAs.svg'))
        self.exportLogBtn.setIconSize(QSize(12, 12))
        self.exportLogBtn.setCursor(Qt.PointingHandCursor)
        hdr.addWidget(self.exportLogBtn)

        lay.addLayout(hdr)

        self.textBrowser = QTextBrowser()
        self.textBrowser.setObjectName("logBrowser")
        self.textBrowser.setOpenExternalLinks(False)
        self.textBrowser.setOpenLinks(False)
        lay.addWidget(self.textBrowser, 1)

        self._build_progress(lay)

        return widget

    # ---- Footer ----
    def _build_footer(self):
        lay = QHBoxLayout()
        lay.setSpacing(8)

        self.helpButton = QPushButton("Aide")
        self.helpButton.setObjectName("btnSecondary")
        self.helpButton.setIcon(_qgs_icon('mActionHelpContents.svg'))
        self.helpButton.setIconSize(QSize(16, 16))
        self.helpButton.setCursor(Qt.PointingHandCursor)
        lay.addWidget(self.helpButton)

        self.diagButton = QPushButton("Diagnostic")
        self.diagButton.setObjectName("btnSecondary")
        self.diagButton.setIcon(_qgs_icon('mActionPropertiesWidget.svg'))
        self.diagButton.setIconSize(QSize(16, 16))
        self.diagButton.setCursor(Qt.PointingHandCursor)
        self.diagButton.setToolTip(
            "Scanner le projet et afficher un rapport complet\n"
            "des ressources detectees, manquantes et couches QGIS."
        )
        lay.addWidget(self.diagButton)

        lay.addStretch()

        self.cancelBtn = QPushButton("  Annuler")
        self.cancelBtn.setObjectName("btnCancel")
        self.cancelBtn.setIcon(_qgs_icon('mTaskCancel.svg'))
        self.cancelBtn.setIconSize(QSize(16, 16))
        self.cancelBtn.setCursor(Qt.PointingHandCursor)
        self.cancelBtn.setVisible(False)
        lay.addWidget(self.cancelBtn)

        self.startBtn = QPushButton("  Démarrer l'analyse")
        self.startBtn.setObjectName("btnStart")
        self.startBtn.setIcon(_qgs_icon('mActionStart.svg'))
        self.startBtn.setIconSize(QSize(16, 16))
        self.startBtn.setCursor(Qt.PointingHandCursor)
        self.startBtn.setEnabled(False)
        self.startBtn.setMinimumWidth(180)
        lay.addWidget(self.startBtn)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)

        outer = QVBoxLayout()
        outer.setSpacing(4)
        outer.addWidget(sep)
        outer.addLayout(lay)
        return outer

    # ==================================================================
    #  SIGNAL CONNECTIONS
    # ==================================================================

    def _connect_signals(self):
        self.browseProjectBtn.clicked.connect(self._browse_project)
        self.refreshBtn.clicked.connect(self._refresh_detection)
        self.browseExportBtn.clicked.connect(self._browse_export)
        self.projectPathEdit.textChanged.connect(self._on_project_path_changed)
        self.selectAllBtn.clicked.connect(self._select_all)
        self.deselectAllBtn.clicked.connect(self._deselect_all)
        self.clearLogBtn.clicked.connect(self.textBrowser.clear)
        self.exportLogBtn.clicked.connect(self._export_log)
        self.startBtn.clicked.connect(self._on_start_clicked)
        self.cancelBtn.clicked.connect(self._on_cancel_clicked)
        self.helpButton.clicked.connect(lambda: self.help_requested.emit())
        self.diagButton.clicked.connect(self._run_diagnostic)
        self._mode_group.buttonClicked.connect(self._on_mode_changed)
        self.textBrowser.anchorClicked.connect(self._on_link_clicked)

    # ==================================================================
    #  PROJECT DETECTION
    # ==================================================================

    def _browse_project(self):
        path = QFileDialog.getExistingDirectory(
            self, "Sélectionner le dossier projet",
            self.projectPathEdit.text() or ""
        )
        if path:
            self.projectPathEdit.setText(path)

    def _browse_export(self):
        path = QFileDialog.getExistingDirectory(
            self, "Sélectionner le dossier d'export",
            self.exportPathEdit.text() or ""
        )
        if path:
            self.exportPathEdit.setText(path)

    def _refresh_detection(self):
        """Re-run detection on the current project path (resync after folder rename)."""
        path = self.projectPathEdit.text()
        if path and os.path.isdir(path):
            self._run_detection(path)
            self._log_info("Detection resynchronisee.")

    def _on_mode_changed(self, button):
        is_project = (button == self._radio_project)
        self._project_mode = is_project

        # Toggle visibility
        self._layers_grp.setVisible(not is_project)
        self._sro_row.setVisible(is_project)
        self._chk_load_layers.setVisible(is_project)

        # Enable project mode radio only if SRO is available
        if is_project and not self._detection.sro:
            self._sro_value.setText("-")
            self._sro_status.setText("[--] SRO non derivable du nom de projet")
            self._sro_status.setStyleSheet("color: #c62828; font-size: 8pt;")

        self._validate_start()

    def _on_project_path_changed(self, path):
        if not path or not os.path.isdir(path):
            self._reset_detection()
            return
        self._run_detection(path)

    def _run_detection(self, path):
        self._detection = detect_project(path)
        d = self._detection

        # Header
        self._project_tag.setText(d.project_name)
        self._project_tag.setVisible(bool(d.project_name))
        self._subtitle.setText(f"  -  {d.project_name}")

        # SRO display
        if d.sro:
            self._sro_value.setText(d.sro)
            self._sro_status.setText("[OK] SRO detecte")
            self._sro_status.setStyleSheet("color: #2e7d32; font-size: 8pt;")
            self._radio_project.setEnabled(True)
        else:
            self._sro_value.setText("-")
            self._sro_status.setText("[--] Format nom projet non reconnu")
            self._sro_status.setStyleSheet("color: #c62828; font-size: 8pt;")
            # Disable project mode if no SRO
            if self._project_mode:
                self._radio_qgis.setChecked(True)
                self._on_mode_changed(self._radio_qgis)
            self._radio_project.setEnabled(False)

        # Export default
        if not self.exportPathEdit.text():
            self.exportPathEdit.setText(d.export_dir)

        # Summary
        parts = []
        for label, p, found in d.summary_lines():
            parts.append(f"{'[OK]' if found else '[--]'} {label}")
        self._det_summary.setText("   ".join(parts))
        self._det_summary.setVisible(True)

        # Module rows with diagnostic hints
        module_map = {
            'maj':       (d.has_ftbt,  d.ftbt_excel,  'FT-BT KO'),
            'capft':     (d.has_capft, d.capft_dir,   'CAP FT'),
            'comac':     (d.has_comac, d.comac_dir,   'COMAC'),
            'c6bd':      (d.has_c6,    d.c6_dir,      'C6 (via CAP FT)'),
            'police_c6': (d.has_c6,    d.c6_dir,      'C6 (via CAP FT)'),
            'c6c3a':     (d.has_c6 and (d.has_c7 or d.has_c3a), d.c6_dir, 'C6 (via CAP FT)'),
        }
        for key, row in self._module_rows.items():
            found, mp, res_label = module_map.get(key, (False, '', ''))
            diag = '' if found else d.get_diagnostic(res_label)
            row.set_detected(found, mp, diagnostic=diag)

        self._update_mod_count()
        self._validate_start()

        # Log with diagnostics for missing resources
        self._log_info(f"Projet : {d.project_name}")
        if d.sro:
            self._log_ok(f"  SRO : {d.sro}")
        has_missing = False
        for label, p, found in d.summary_lines():
            if found:
                self._log_ok(f"  {label} : {os.path.basename(p)}")
            else:
                self._log_dim(f"  {label} : non detecte")
                has_missing = True
                diag = d.get_diagnostic(label)
                if diag:
                    for line in diag.split('\n'):
                        self._log_dim(f"      {line}")
        if has_missing:
            self._log_warn(
                "Conseil : verifiez le nommage des dossiers/fichiers "
                "ou utilisez le bouton Rafraichir apres renommage."
            )

    def _reset_detection(self):
        self._detection = DetectionResult()
        self._project_tag.setVisible(False)
        self._subtitle.setText("  -  Selectionnez un dossier projet")
        self._det_summary.setVisible(False)
        self._sro_value.setText("-")
        self._sro_status.setText("")
        self._radio_project.setEnabled(False)
        for row in self._module_rows.values():
            row.set_detected(False)
        self._update_mod_count()
        self._validate_start()

    # ==================================================================
    #  MODULE SELECTION
    # ==================================================================

    def _on_module_toggled(self, key, checked):
        self._update_mod_count()
        self._validate_start()

    def _select_all(self):
        for row in self._module_rows.values():
            if row.checkbox.isEnabled():
                row.checkbox.setChecked(True)

    def _deselect_all(self):
        for row in self._module_rows.values():
            row.checkbox.setChecked(False)

    def _update_mod_count(self):
        sel = sum(1 for r in self._module_rows.values() if r.is_selected)
        self._mod_count.setText(f"{sel}/{len(self._module_rows)} sélectionné(s)")

    def _validate_start(self):
        has_proj = bool(self.projectPathEdit.text()) and os.path.isdir(
            self.projectPathEdit.text() or ''
        )
        has_mod = any(r.is_selected for r in self._module_rows.values())
        # In project mode, SRO must be available
        if self._project_mode and not self._detection.sro:
            self.startBtn.setEnabled(False)
            return
        self.startBtn.setEnabled(has_proj and has_mod and not self._is_running)

    def selected_modules(self):
        order = ['maj', 'capft', 'comac', 'c6bd', 'police_c6', 'c6c3a']
        return [k for k in order if k in self._module_rows and self._module_rows[k].is_selected]

    # ==================================================================
    #  EXECUTION STATE
    # ==================================================================

    def _on_start_clicked(self):
        mods = self.selected_modules()
        if mods:
            self.start_requested.emit(mods)

    def _on_cancel_clicked(self):
        self.cancel_requested.emit()

    def set_running(self, running):
        self._is_running = running
        self.startBtn.setVisible(not running)
        self.cancelBtn.setVisible(running)
        self.browseProjectBtn.setEnabled(not running)
        self.browseExportBtn.setEnabled(not running)
        self.projectPathEdit.setReadOnly(running)
        self.selectAllBtn.setEnabled(not running)
        self.deselectAllBtn.setEnabled(not running)

        for row in self._module_rows.values():
            row.checkbox.setEnabled(not running and row._found)

        if running:
            self.smooth_progress.reset()
            self.smooth_progress.set_target(2)
        self._validate_start()

    def set_progress(self, percent):
        self.smooth_progress.set_target(percent)

    def reset_after_batch(self):
        self._is_running = False
        self.smooth_progress.set_target(100)
        self.startBtn.setVisible(True)
        self.cancelBtn.setVisible(False)
        self.browseProjectBtn.setEnabled(True)
        self.browseExportBtn.setEnabled(True)
        self.projectPathEdit.setReadOnly(False)
        self.selectAllBtn.setEnabled(True)
        self.deselectAllBtn.setEnabled(True)
        for row in self._module_rows.values():
            row.checkbox.setEnabled(row._found)
        self._validate_start()

    # ==================================================================
    #  DIAGNOSTIC
    # ==================================================================

    def _run_diagnostic(self):
        """Pre-diagnostic intelligent du livrable et de l'environnement."""
        self.textBrowser.clear()

        path = self.projectPathEdit.text()
        if not path or not os.path.isdir(path):
            self._log_err("Aucun dossier projet selectionne.")
            return

        # Ensure detection is up-to-date
        d = self._detection
        if not d.project_name:
            self._run_detection(path)
            d = self._detection

        # Deep analysis
        analysis = analyse_livrable(d)

        # ── Header ──
        self._log_info(f"<b>=== PRE-DIAGNOSTIC : {d.project_name or os.path.basename(path)} ===</b>")
        if d.sro:
            self._log_ok(f"  SRO : {d.sro}")
        else:
            self._log_warn(
                "  SRO : non reconnu (format attendu : XXXXX-YYY-ZZZ-NNNNN)"
            )

        # ── 1. Contenu du livrable ──
        self._log_info("")
        self._log_info("<b>--- Contenu du livrable ---</b>")

        # CAP FT
        if d.has_capft:
            cap = analysis.get('capft', {})
            parts = []
            if cap.get('cmd'):
                parts.append(f"{cap['cmd']} CMD")
            if cap.get('etudes'):
                parts.append(f"{cap['etudes']} etude(s)")
            if cap.get('excels'):
                parts.append(f"{cap['excels']} fichiers Excel")
            self._log_ok(f"  CAP FT ........... {' | '.join(parts) if parts else 'detecte'}")
        else:
            self._log_warn("  CAP FT ........... Non livre")

        # COMAC
        if d.has_comac:
            com = analysis.get('comac', {})
            parts = []
            if com.get('sous_dossiers'):
                parts.append(f"{com['sous_dossiers']} sous-dossier(s)")
            if com.get('excels'):
                parts.append(f"{com['excels']} fichiers Excel")
            self._log_ok(f"  COMAC ............ {' | '.join(parts) if parts else 'detecte'}")
        else:
            self._log_warn("  COMAC ............ Non livre")

        # FT-BT KO
        if d.has_ftbt:
            nom = analysis.get('ftbt', {}).get('fichier', '?')
            self._log_ok(f"  FT-BT KO ......... {nom}")
        else:
            self._log_warn("  FT-BT KO ......... Non livre")

        # C7 / C3A
        if d.has_c7:
            self._log_ok(f"  C7 ............... {analysis.get('c7', {}).get('fichier', '?')}")
        else:
            self._log_dim("  C7 ............... Non livre")
        if d.has_c3a:
            self._log_ok(f"  C3A .............. {analysis.get('c3a', {}).get('fichier', '?')}")
        else:
            self._log_dim("  C3A .............. Non livre")

        # ── 2. Environnement QGIS ──
        self._log_info("")
        self._log_info("<b>--- Environnement QGIS ---</b>")

        mode_label = "Projet (PostgreSQL)" if self._project_mode else "Couches QGIS (manuel)"
        self._log_dim(f"  Mode : {mode_label}")

        # PostgreSQL
        pg_ok = False
        if self._project_mode:
            try:
                from .db_connection import DatabaseConnection
                db = DatabaseConnection()
                conn_name = db.find_auvergne_connection()
                if conn_name and db.connect():
                    self._log_ok(f"  PostgreSQL : OK ({conn_name})")
                    pg_ok = True
                    db.disconnect()
                elif conn_name:
                    self._log_warn(f"  PostgreSQL : connexion '{conn_name}' trouvee mais echec")
                else:
                    self._log_warn("  PostgreSQL : aucune connexion 'Auvergne' configuree")
            except Exception as exc:
                self._log_warn(f"  PostgreSQL : erreur ({exc})")

            if pg_ok:
                self._log_dim("  Couches : chargement automatique via SRO")
            else:
                self._log_warn("  Couches : chargement impossible sans connexion PG")
        else:
            # QGIS layers mode - check selected layers
            layer_checks = [
                (self.comboInfraPtPot.currentLayer(), 'infra_pt_pot',
                 'Poteaux (Point)'),
                (self.comboEtudeCapFt.currentLayer(), 'etude_cap_ft',
                 'Zonage CAP FT (Polygone)'),
                (self.comboEtudeComac.currentLayer(), 'etude_comac',
                 'Zonage COMAC (Polygone)'),
            ]
            crs_set = set()
            for lyr, key, desc in layer_checks:
                if lyr:
                    fc = lyr.featureCount()
                    crs = lyr.crs().authid() if lyr.crs().isValid() else '?'
                    crs_set.add(crs)
                    if fc > 0:
                        self._log_ok(
                            f"  {key} : {lyr.name()} ({fc:,} entites, {crs})")
                    else:
                        self._log_warn(
                            f"  {key} : {lyr.name()} (VIDE - 0 entite)")
                else:
                    self._log_warn(f"  {key} : non selectionnee ({desc})")

            if len(crs_set) > 1:
                self._log_warn(
                    f"  CRS : ATTENTION, CRS heterogenes ({', '.join(crs_set)})")
            elif len(crs_set) == 1:
                self._log_ok(f"  CRS : coherent ({crs_set.pop()})")

        # Export
        export_dir = self.exportPathEdit.text() or d.export_dir
        if export_dir and os.path.isdir(export_dir):
            try:
                test_path = os.path.join(export_dir, '.poleaerien_diag_test')
                with open(test_path, 'w') as _f:
                    pass
                os.remove(test_path)
                self._log_ok(f"  Export : {export_dir}")
            except OSError:
                self._log_err(f"  Export : {export_dir} (NON inscriptible)")
        elif export_dir:
            self._log_err(f"  Export : {export_dir} (dossier introuvable)")
        else:
            self._log_dim("  Export : racine projet (par defaut)")

        # ── 3. Modules ──
        self._log_info("")
        self._log_info("<b>--- Modules ---</b>")

        detected = d.detected_modules()

        module_defs = [
            ('maj',       'MAJ BD',       self._diag_prereq_maj),
            ('capft',     'CAP_FT',       self._diag_prereq_capft),
            ('comac',     'COMAC',        self._diag_prereq_comac),
            ('c6bd',      'C6 vs BD',     self._diag_prereq_c6),
            ('police_c6', 'POLICE C6',    self._diag_prereq_c6),
            ('c6c3a',     'C6-C3A-C7-BD', self._diag_prereq_c6c3a),
        ]

        n_ready = 0
        for key, name, prereq_fn in module_defs:
            if key in detected:
                detail = prereq_fn(analysis)
                self._log_ok(f"  PRET    {name:.<16s} {detail}")
                n_ready += 1
            else:
                raison = prereq_fn(analysis, missing=True)
                self._log_warn(f"  BLOQUE  {name:.<16s} {raison}")

        # ── 4. Recommandations ──
        recommandations = self._build_recommandations(d, detected, pg_ok)
        if recommandations:
            self._log_info("")
            self._log_info("<b>--- Recommandations ---</b>")
            for reco in recommandations:
                self._log_dim(f"  > {reco}")

        # ── Verdict ──
        self._log_info("")
        total = len(module_defs)
        if n_ready == total:
            self._log_ok(f"<b>Verdict : {n_ready}/{total} modules prets. Livrable complet.</b>")
        elif n_ready > 0:
            self._log_warn(
                f"<b>Verdict : {n_ready}/{total} modules prets. "
                f"Livrable partiel.</b>")
        else:
            self._log_err(f"<b>Verdict : 0/{total} modules prets.</b>")

    # --- Diagnostic helpers ---

    def _diag_prereq_maj(self, analysis, missing=False):
        if missing:
            return "fichier FT-BT KO manquant a la racine"
        return analysis.get('ftbt', {}).get('fichier', '')

    def _diag_prereq_capft(self, analysis, missing=False):
        if missing:
            return "dossier CAP FT non detecte"
        cap = analysis.get('capft', {})
        parts = []
        if cap.get('cmd'):
            parts.append(f"{cap['cmd']} CMD")
        if cap.get('etudes'):
            parts.append(f"{cap['etudes']} etudes")
        return ', '.join(parts) if parts else 'detecte'

    def _diag_prereq_comac(self, analysis, missing=False):
        if missing:
            return "dossier COMAC non detecte"
        com = analysis.get('comac', {})
        n = com.get('sous_dossiers', 0)
        return f"{n} etude(s)" if n else 'detecte'

    def _diag_prereq_c6(self, analysis, missing=False):
        if missing:
            return "dossier CAP FT requis (C6 = etudes dans CAP FT)"
        cap = analysis.get('capft', {})
        n = cap.get('etudes', 0)
        return f"via CAP FT ({n} etudes)" if n else 'via CAP FT'

    def _diag_prereq_c6c3a(self, analysis, missing=False):
        if missing:
            return "C7 et/ou C3A requis en complement de C6"
        parts = []
        if 'c7' in analysis:
            parts.append('C7')
        if 'c3a' in analysis:
            parts.append('C3A')
        return f"C6 + {' + '.join(parts)}" if parts else 'detecte'

    def _build_recommandations(self, d, detected, pg_ok):
        recos = []
        if 'maj' not in detected:
            recos.append(
                'MAJ BD : placer le fichier Excel FT-BT KO '
                'a la racine du projet')
        if 'capft' not in detected:
            recos.append(
                'CAP FT / C6 / Police : creer le dossier CAP FT '
                '(ou ETUDE_CAPFT) avec les CMD a l\'interieur')
        if 'comac' not in detected:
            recos.append(
                'COMAC : creer le dossier COMAC '
                '(ou ETUDE_COMAC) avec les sous-dossiers d\'etude')
        if 'c6c3a' not in detected and 'c6bd' in detected:
            missing = []
            if not d.has_c7:
                missing.append('C7')
            if not d.has_c3a:
                missing.append('C3A')
            if missing:
                recos.append(
                    f"C6-C3A-C7-BD : ajouter les annexes "
                    f"{' et '.join(missing)}")
        if self._project_mode and not pg_ok:
            recos.append(
                'Mode Projet : configurer la connexion PostgreSQL '
                'Auvergne dans QGIS > Parametres > PostgreSQL')
        if not self._project_mode:
            any_missing = not all([
                self.comboInfraPtPot.currentLayer(),
                self.comboEtudeCapFt.currentLayer(),
                self.comboEtudeComac.currentLayer(),
            ])
            if any_missing:
                recos.append(
                    'Mode Couches : selectionner les 3 couches QGIS '
                    '(poteaux, zonage CAP FT, zonage COMAC)')
        return recos

    # ==================================================================
    #  LOGGING
    # ==================================================================

    def _log_html(self, html):
        self.textBrowser.append(html)
        c = self.textBrowser.textCursor()
        c.movePosition(QTextCursor.End)
        self.textBrowser.setTextCursor(c)

    def _log_info(self, msg):
        self._log_html(f'<span style="color:#1565c0">{msg}</span>')

    def _log_ok(self, msg):
        self._log_html(f'<span style="color:#2e7d32">{msg}</span>')

    def _log_warn(self, msg):
        self._log_html(f'<span style="color:#e65100">{msg}</span>')

    def _log_err(self, msg):
        self._log_html(f'<span style="color:#c62828"><b>{msg}</b></span>')

    def _log_dim(self, msg):
        self._log_html(f'<span style="color:gray">{msg}</span>')

    def log_result_link(self, filepath):
        """Log a clickable link to open a result file."""
        fname = os.path.basename(filepath)
        url = filepath.replace('\\', '/')
        self._log_html(
            f'<span style="color:#2e7d32">Rapport : '
            f'<a href="file:///{url}" style="color:#1565c0; '
            f'text-decoration:underline;">{fname}</a></span>'
        )

    def _on_link_clicked(self, url):
        """Open a file link clicked in the log browser."""
        path = url.toLocalFile()
        if path and os.path.isfile(path):
            os.startfile(path)
        else:
            self._log_warn(f"Fichier introuvable : {path}")

    def log_message(self, msg, level='info'):
        dispatch = {
            'info': self._log_info,
            'success': self._log_ok,
            'warning': self._log_warn,
            'error': self._log_err,
        }
        dispatch.get(level, self._log_info)(msg)

    def _export_log(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter le journal", "", "Fichier texte (*.txt)"
        )
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(self.textBrowser.toPlainText())
                self._log_ok(f"Journal exporté : {path}")
            except OSError as exc:
                self._log_err(f"Erreur export : {exc}")

    # ==================================================================
    #  PUBLIC API
    # ==================================================================

    @property
    def detection(self):
        return self._detection

    @property
    def is_project_mode(self):
        """True when 'Mode Projet (BDD)' is selected."""
        return self._project_mode

    @property
    def load_layers_in_qgis(self):
        """True when user wants DB layers added to the QGIS project."""
        return self._project_mode and self._chk_load_layers.isChecked()

    @property
    def sro(self):
        """SRO derived from project name, or empty string."""
        return self._detection.sro

    def get_export_dir(self):
        return self.exportPathEdit.text() or self._detection.export_dir

    def init_default_layers(self):
        # Collect candidates with priority (higher = better match)
        best_pot = (None, 0)
        best_cap = (None, 0)
        best_com = (None, 0)

        for layer in QgsProject.instance().mapLayers().values():
            name = layer.name().lower()
            geom = layer.geometryType() if hasattr(layer, 'geometryType') else -1

            # infra_pt_pot: Point layers only
            if 'infra_pt_pot' in name and geom == 0:  # QgsWkbTypes.PointGeometry
                best_pot = (layer, 1)

            # etude_cap_ft: Polygon layers, prefer exact match
            if geom == 2:  # QgsWkbTypes.PolygonGeometry
                if 'etude_cap_ft' in name:
                    best_cap = (layer, 3)
                elif ('cap_ft' in name or 'capft' in name) and best_cap[1] < 2:
                    best_cap = (layer, 2)

            # etude_comac: Polygon layers, prefer exact match
            if geom == 2:
                if 'etude_comac' in name:
                    best_com = (layer, 3)
                elif 'comac' in name and best_com[1] < 2:
                    best_com = (layer, 2)

        if best_pot[0]:
            self.comboInfraPtPot.setLayer(best_pot[0])
        if best_cap[0]:
            self.comboEtudeCapFt.setLayer(best_cap[0])
        if best_com[0]:
            self.comboEtudeComac.setLayer(best_com[0])

    def closeEvent(self, event):
        if self._is_running:
            self.cancel_requested.emit()
        if self.smooth_progress:
            self.smooth_progress.reset()
        super().closeEvent(event)
