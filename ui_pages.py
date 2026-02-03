# -*- coding: utf-8 -*-
"""
Page builder for modern QGIS plugin UI.
Creates standardized pages with header, content, footer structure.
"""

import os
from qgis.PyQt.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QFrame,
    QScrollArea, QProgressBar, QSpacerItem, QSizePolicy,
    QRadioButton
)
from qgis.PyQt.QtCore import Qt, QSize
from qgis.PyQt.QtGui import QFont, QIcon
from qgis.gui import QgsMapLayerComboBox, QgsFileWidget, QgsCollapsibleGroupBox
from qgis.core import QgsMapLayerProxyModel, QgsProject

# Plugin dir for icons
PLUGIN_DIR = os.path.dirname(__file__)

# Style constants
STYLES = {
    'page_header': 'background:#fff;border-bottom:1px solid #e2e8f0;',
    'page_content': 'background:#f8fafc;border:none;',
    'page_footer': 'background:#fff;border-top:1px solid #e2e8f0;',
    'title': 'font-size:14pt;font-weight:bold;color:#1e293b;',
    'subtitle': 'font-size:9pt;color:#64748b;',
    'groupbox': '''QgsCollapsibleGroupBox{font-weight:bold;background:#fff;
        border:1px solid #e2e8f0;border-radius:6px;padding-top:16px;}''',
    'kpi_default': '''QLabel{background:#f1f5f9;color:#475569;padding:8px 14px;
        border-radius:8px;font-weight:600;font-size:10pt;min-width:60px;}''',
    'kpi_ok': '''QLabel{background:qlineargradient(x1:0,y1:0,x2:0,y2:1,stop:0 #dcfce7,stop:1 #bbf7d0);
        color:#166534;padding:8px 14px;border-radius:8px;font-weight:600;font-size:10pt;}''',
    'kpi_warn': '''QLabel{background:qlineargradient(x1:0,y1:0,x2:0,y2:1,stop:0 #fef3c7,stop:1 #fde68a);
        color:#92400e;padding:8px 14px;border-radius:8px;font-weight:600;font-size:10pt;}''',
    'kpi_error': '''QLabel{background:qlineargradient(x1:0,y1:0,x2:0,y2:1,stop:0 #fee2e2,stop:1 #fecaca);
        color:#991b1b;padding:8px 14px;border-radius:8px;font-weight:600;font-size:10pt;}''',
    'kpi_info': '''QLabel{background:qlineargradient(x1:0,y1:0,x2:0,y2:1,stop:0 #dbeafe,stop:1 #bfdbfe);
        color:#1d4ed8;padding:8px 14px;border-radius:8px;font-weight:600;font-size:10pt;}''',
    'status_empty': 'background:#f1f5f9;color:#64748b;padding:6px 12px;border-radius:4px;',
    'status_ready': 'background:#dcfce7;color:#166534;padding:6px 12px;border-radius:4px;',
    'status_running': 'background:#dbeafe;color:#1d4ed8;padding:6px 12px;border-radius:4px;',
    'status_error': 'background:#fee2e2;color:#991b1b;padding:6px 12px;border-radius:4px;',
    'progress': '''QProgressBar{background:#e2e8f0;border-radius:4px;}
        QProgressBar::chunk{background:#2563eb;border-radius:4px;}''',
    'btn_primary': '''QPushButton{background:#2563eb;color:#fff;border:none;
        border-radius:6px;padding:10px 24px;font-weight:bold;}
        QPushButton:hover{background:#1d4ed8;}
        QPushButton:disabled{background:#cbd5e0;color:#a0aec0;}''',
}


class PageBuilder:
    """Builder for standardized module pages."""

    @staticmethod
    def create_header(title, subtitle, kpis=None):
        """Create page header with title, subtitle, and KPI badges."""
        frame = QFrame()
        frame.setMinimumHeight(80)
        frame.setMaximumHeight(80)
        frame.setStyleSheet(STYLES['page_header'])

        layout = QHBoxLayout(frame)
        layout.setContentsMargins(20, 12, 20, 12)

        # Title block
        title_layout = QVBoxLayout()
        title_layout.setSpacing(2)

        lbl_title = QLabel(title)
        lbl_title.setStyleSheet(STYLES['title'])
        title_layout.addWidget(lbl_title)

        lbl_subtitle = QLabel(subtitle)
        lbl_subtitle.setStyleSheet(STYLES['subtitle'])
        title_layout.addWidget(lbl_subtitle)

        layout.addLayout(title_layout)
        layout.addStretch()

        # KPI badges
        kpi_widgets = {}
        if kpis:
            for name, (text, style_key) in kpis.items():
                lbl = QLabel(text)
                lbl.setStyleSheet(STYLES.get(style_key, STYLES['kpi_default']))
                layout.addWidget(lbl)
                kpi_widgets[name] = lbl

        # Status badge
        status_lbl = QLabel('Vide')
        status_lbl.setStyleSheet(STYLES['status_empty'])
        layout.addWidget(status_lbl)
        kpi_widgets['status'] = status_lbl

        return frame, kpi_widgets

    @staticmethod
    def create_scroll_content():
        """Create scrollable content area."""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(STYLES['page_content'])

        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)

        scroll.setWidget(content)
        return scroll, layout

    @staticmethod
    def create_footer(btn_name='Exécuter'):
        """Create page footer with progress bar and action button."""
        frame = QFrame()
        frame.setMinimumHeight(64)
        frame.setMaximumHeight(64)
        frame.setStyleSheet(STYLES['page_footer'])

        layout = QHBoxLayout(frame)
        layout.setContentsMargins(20, 12, 20, 12)

        # Progress bar
        progress = QProgressBar()
        progress.setMinimumSize(200, 8)
        progress.setMaximumSize(300, 8)
        progress.setValue(0)
        progress.setTextVisible(False)
        progress.setStyleSheet(STYLES['progress'])
        layout.addWidget(progress)

        # Progress text
        progress_text = QLabel('')
        progress_text.setStyleSheet('color:#64748b;font-size:9pt;')
        layout.addWidget(progress_text)

        layout.addStretch()

        # Action button with icon
        btn = QPushButton(btn_name)
        btn.setEnabled(False)
        btn.setMinimumSize(140, 36)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setStyleSheet(STYLES['btn_primary'])
        icon_path = os.path.join(PLUGIN_DIR, 'images', 'icon-play.svg')
        if os.path.exists(icon_path):
            btn.setIcon(QIcon(icon_path))
            btn.setIconSize(QSize(18, 18))
        layout.addWidget(btn)

        return frame, {'progress': progress, 'progress_text': progress_text, 'btn': btn}

    @staticmethod
    def create_groupbox(title):
        """Create styled collapsible group box."""
        gb = QgsCollapsibleGroupBox(title)
        gb.setStyleSheet(STYLES['groupbox'])

        layout = QFormLayout()
        layout.setHorizontalSpacing(16)
        layout.setVerticalSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)
        gb.setLayout(layout)

        return gb, layout

    @staticmethod
    def create_layer_combo(min_width=280):
        """Create QgsMapLayerComboBox."""
        cb = QgsMapLayerComboBox()
        cb.setMinimumWidth(min_width)
        cb.setProject(QgsProject.instance())
        cb.setFilters(QgsMapLayerProxyModel.HasGeometry)
        return cb

    @staticmethod
    def create_field_combo(min_width=100):
        """Create QComboBox for field selection."""
        cb = QComboBox()
        cb.setMinimumWidth(min_width)
        return cb

    @staticmethod
    def create_path_row(placeholder='', icon_type='folder'):
        """Create path input with browse button.
        
        Args:
            placeholder: Placeholder text
            icon_type: 'folder', 'export', 'excel', 'import'
        """
        layout = QHBoxLayout()
        layout.setSpacing(8)

        line = QLineEdit()
        line.setPlaceholderText(placeholder)
        line.setMinimumHeight(28)
        layout.addWidget(line)

        btn = QPushButton()
        btn.setMinimumSize(32, 28)
        btn.setMaximumSize(32, 28)
        btn.setCursor(Qt.PointingHandCursor)
        
        # Icon mapping
        icon_map = {
            'folder': 'icon-folder-open.svg',
            'export': 'icon-export.svg',
            'excel': 'icon-excel.svg',
            'import': 'icon-import.svg',
        }
        icon_file = icon_map.get(icon_type, 'icon-folder-open.svg')
        icon_path = os.path.join(PLUGIN_DIR, 'images', icon_file)
        if os.path.exists(icon_path):
            btn.setIcon(QIcon(icon_path))
            btn.setIconSize(QSize(18, 18))
        else:
            btn.setText('...')
        
        layout.addWidget(btn)

        return layout, line, btn


def build_page_maj(parent=None):
    """Build MAJ BD page."""
    page = QWidget(parent)
    main_layout = QVBoxLayout(page)
    main_layout.setSpacing(0)
    main_layout.setContentsMargins(0, 0, 0, 0)

    # Header
    kpis = {
        'total': ('—', 'kpi_default'),
        'ft': ('FT: —', 'kpi_info'),
        'bt': ('BT: —', 'kpi_warn'),
    }
    header, kpi_widgets = PageBuilder.create_header(
        'Mise à jour BD',
        'Import des poteaux FT/BT KO depuis fichier Excel',
        kpis
    )
    main_layout.addWidget(header)

    # Content
    scroll, content_layout = PageBuilder.create_scroll_content()

    # Group: Couches
    gb_couches, form_couches = PageBuilder.create_groupbox('Données sources')

    cb_infra = PageBuilder.create_layer_combo()
    form_couches.addRow('Couche Poteaux', cb_infra)

    cb_etude_ft = PageBuilder.create_layer_combo()
    form_couches.addRow('Études CAP FT', cb_etude_ft)

    cb_etude_comac = PageBuilder.create_layer_combo()
    form_couches.addRow('Études COMAC', cb_etude_comac)

    content_layout.addWidget(gb_couches)

    # Group: Fichier - use VBoxLayout for better FileWidget display
    gb_file = QgsCollapsibleGroupBox('Fichier FT-BT KO')
    gb_file.setStyleSheet(STYLES['groupbox'])
    file_layout = QVBoxLayout()
    file_layout.setSpacing(12)
    file_layout.setContentsMargins(16, 20, 16, 16)
    
    file_label = QLabel('Fichier Excel FT-BT KO :')
    file_label.setStyleSheet('font-weight:normal;')
    file_layout.addWidget(file_label)
    
    file_widget = QgsFileWidget()
    file_widget.setStorageMode(QgsFileWidget.GetFile)
    file_widget.setFilter('Fichiers Excel (*.xlsx *.xls)')
    file_widget.setMinimumHeight(28)
    file_layout.addWidget(file_widget)

    file_status = QLabel('Aucun fichier sélectionné')
    file_status.setStyleSheet('color:#64748b;font-size:9pt;font-weight:normal;')
    file_layout.addWidget(file_status)
    
    gb_file.setLayout(file_layout)
    content_layout.addWidget(gb_file)
    content_layout.addStretch()

    main_layout.addWidget(scroll)

    # Footer
    footer, footer_widgets = PageBuilder.create_footer('Exécuter')
    main_layout.addWidget(footer)

    # Store refs
    page.widgets = {
        'MajcomboBox_infra_pt_pot': cb_infra,
        'MajcomboBox_etude_cap_ft': cb_etude_ft,
        'MajcomboBox_etude_comac': cb_etude_comac,
        'MajFileWidget': file_widget,
        'majFileStatus': file_status,
        'majBdLanceur': footer_widgets['btn'],
        'progressMaj': footer_widgets['progress'],
        'progressMajText': footer_widgets['progress_text'],
        'kpis': kpi_widgets,
    }

    return page


def build_page_c6bd(parent=None):
    """Build C6 vs BD page."""
    page = QWidget(parent)
    main_layout = QVBoxLayout(page)
    main_layout.setSpacing(0)
    main_layout.setContentsMargins(0, 0, 0, 0)

    # Header
    kpis = {
        'total': ('—', 'kpi_default'),
        'ok': ('OK: —', 'kpi_ok'),
        'diff': ('Diff: —', 'kpi_error'),
    }
    header, kpi_widgets = PageBuilder.create_header(
        'C6 vs BD',
        'Comparaison annexe C6 avec la base de données',
        kpis
    )
    main_layout.addWidget(header)

    # Content
    scroll, content_layout = PageBuilder.create_scroll_content()

    # Group: Couches
    gb_couches, form_couches = PageBuilder.create_groupbox('Couches sources')

    cb_infra = PageBuilder.create_layer_combo()
    form_couches.addRow('Couche Poteaux', cb_infra)

    cb_etude = PageBuilder.create_layer_combo(250)
    form_couches.addRow('Zone d\'étude CAP FT', cb_etude)

    content_layout.addWidget(gb_couches)

    # Group: Répertoires
    gb_rep, form_rep = PageBuilder.create_groupbox('Répertoires')

    row_c6, line_c6, btn_c6 = PageBuilder.create_path_row('Dossier contenant les C6 (.xlsx)', 'folder')
    form_rep.addRow('Répertoire C6', row_c6)

    row_exp, line_exp, btn_exp = PageBuilder.create_path_row('Dossier d\'export', 'export')
    form_rep.addRow('Répertoire Export', row_exp)

    content_layout.addWidget(gb_rep)
    content_layout.addStretch()

    main_layout.addWidget(scroll)

    # Footer
    footer, footer_widgets = PageBuilder.create_footer('Exécuter')
    main_layout.addWidget(footer)

    page.widgets = {
        'C6BdcomboBox_infra_pt_pot': cb_infra,
        'C6BdcomboBox_etude_cap_ft': cb_etude,
        'lienCheminFichiersC6': line_c6,
        'C6BdboutonCheminFichiersC6': btn_c6,
        'lienCheminExportDonnees': line_exp,
        'C6BdboutonCheminExportDonnees': btn_exp,
        'C6BdLanceur': footer_widgets['btn'],
        'progressC6Bd': footer_widgets['progress'],
        'kpis': kpi_widgets,
    }

    return page


def build_page_capft(parent=None):
    """Build CAP_FT page."""
    page = QWidget(parent)
    main_layout = QVBoxLayout(page)
    main_layout.setSpacing(0)
    main_layout.setContentsMargins(0, 0, 0, 0)

    # Header
    kpis = {
        'total': ('—', 'kpi_default'),
        'ok': ('OK: —', 'kpi_ok'),
        'warn': ('Alerte: —', 'kpi_warn'),
    }
    header, kpi_widgets = PageBuilder.create_header(
        'Vérification CAP_FT',
        'Contrôle des poteaux vs fiches appuis FT',
        kpis
    )
    main_layout.addWidget(header)

    # Content
    scroll, content_layout = PageBuilder.create_scroll_content()

    # Group: Couches
    gb_couches, form_couches = PageBuilder.create_groupbox('Couches sources')

    cb_infra = PageBuilder.create_layer_combo()
    form_couches.addRow('Couche Poteaux', cb_infra)

    row_etude = QHBoxLayout()
    cb_etude = PageBuilder.create_layer_combo(180)
    row_etude.addWidget(cb_etude)
    row_etude.addWidget(QLabel('Champ'))
    cb_champs = PageBuilder.create_field_combo()
    row_etude.addWidget(cb_champs)
    form_couches.addRow('Zone d\'étude', row_etude)

    content_layout.addWidget(gb_couches)

    # Group: Répertoires
    gb_rep, form_rep = PageBuilder.create_groupbox('Répertoires')

    row_capft, line_capft, btn_capft = PageBuilder.create_path_row('Chemin contenant les études CAP_FT', 'folder')
    form_rep.addRow('Répertoire CAP_FT', row_capft)

    row_exp, line_exp, btn_exp = PageBuilder.create_path_row('Chemin d\'export', 'export')
    form_rep.addRow('Répertoire Export', row_exp)

    content_layout.addWidget(gb_rep)
    content_layout.addStretch()

    main_layout.addWidget(scroll)

    # Footer
    footer, footer_widgets = PageBuilder.create_footer('Exécuter')
    main_layout.addWidget(footer)

    page.widgets = {
        'capFtComboBoxCoucheInfra_pt_pot': cb_infra,
        'capFtComboBox_etude_cap_ft': cb_etude,
        'capFtComboBoxChampsCapFt': cb_champs,
        'lienRepertoireCapFt': line_capft,
        'boutonCheminEtudeCapFt': btn_capft,
        'lienCheminExportCapFt': line_exp,
        'boutonCheminExportComac': btn_exp,
        'cap_ftLanceur': footer_widgets['btn'],
        'progressCapFt': footer_widgets['progress'],
        'kpis': kpi_widgets,
    }

    return page


def build_page_comac(parent=None):
    """Build COMAC page."""
    page = QWidget(parent)
    main_layout = QVBoxLayout(page)
    main_layout.setSpacing(0)
    main_layout.setContentsMargins(0, 0, 0, 0)

    # Header
    kpis = {
        'total': ('—', 'kpi_default'),
        'ok': ('OK: —', 'kpi_ok'),
        'crit': ('Critique: —', 'kpi_error'),
    }
    header, kpi_widgets = PageBuilder.create_header(
        'Vérification COMAC',
        'Contrôle BT vs ExportComac + règles sécurité NFC 11201',
        kpis
    )
    main_layout.addWidget(header)

    # Content
    scroll, content_layout = PageBuilder.create_scroll_content()

    # Group: Couches
    gb_couches, form_couches = PageBuilder.create_groupbox('Couches sources')

    cb_infra = PageBuilder.create_layer_combo()
    form_couches.addRow('Couche Poteaux', cb_infra)

    row_etude = QHBoxLayout()
    cb_etude = PageBuilder.create_layer_combo(180)
    row_etude.addWidget(cb_etude)
    row_etude.addWidget(QLabel('Champ'))
    cb_champs = PageBuilder.create_field_combo()
    row_etude.addWidget(cb_champs)
    form_couches.addRow('Zone d\'étude COMAC', row_etude)

    content_layout.addWidget(gb_couches)

    # Group: Répertoires
    gb_rep, form_rep = PageBuilder.create_groupbox('Répertoires')

    row_comac, line_comac, btn_comac = PageBuilder.create_path_row('Chemin contenant les études COMAC', 'folder')
    form_rep.addRow('Répertoire COMAC', row_comac)

    row_exp, line_exp, btn_exp = PageBuilder.create_path_row('Chemin d\'export', 'export')
    form_rep.addRow('Répertoire Export', row_exp)

    content_layout.addWidget(gb_rep)

    # Group: Zone climatique (hidden by default)
    gb_zone, form_zone = PageBuilder.create_groupbox('Zone climatique')
    cb_zone = QComboBox()
    cb_zone.addItems(['ZVN', 'ZVF'])
    cb_zone.setToolTip('Zone climatique pour calcul des portées max (NFC 11201-A1)')
    form_zone.addRow('Zone', cb_zone)
    gb_zone.setVisible(False)
    content_layout.addWidget(gb_zone)

    content_layout.addStretch()

    main_layout.addWidget(scroll)

    # Footer
    footer, footer_widgets = PageBuilder.create_footer('Exécuter')
    main_layout.addWidget(footer)

    page.widgets = {
        'comacComboBoxCoucheInfra_pt_pot': cb_infra,
        'comboBoxCoucheComac': cb_etude,
        'comboBoxChampsComac': cb_champs,
        'lienRepertoireComac': line_comac,
        'boutonCheminEtudeComac': btn_comac,
        'lienCheminExportComac': line_exp,
        'boutonCheminExportCapFt': btn_exp,
        'comboBoxZoneClimatique': cb_zone,
        'cap_comacLanceur': footer_widgets['btn'],
        'progressComac': footer_widgets['progress'],
        'kpis': kpi_widgets,
    }

    return page


def build_page_police_c6(parent=None):
    """Build Police C6 / COMAC page."""
    page = QWidget(parent)
    main_layout = QVBoxLayout(page)
    main_layout.setSpacing(0)
    main_layout.setContentsMargins(0, 0, 0, 0)

    # Header
    kpis = {
        'total': ('—', 'kpi_default'),
        'ok': ('OK: —', 'kpi_ok'),
        'anomalie': ('Anomalie: —', 'kpi_error'),
    }
    header, kpi_widgets = PageBuilder.create_header(
        'Police C6 / COMAC',
        'Analyse Police C6 + GraceTHD + études COMAC',
        kpis
    )
    main_layout.addWidget(header)

    # Content
    scroll, content_layout = PageBuilder.create_scroll_content()

    # Group: Fichier C6
    gb_file, form_file = PageBuilder.create_groupbox('Fichier C6')

    row_c6, line_c6, btn_c6 = PageBuilder.create_path_row('Annexe C6 (.xlsx)', 'import')
    form_file.addRow('Fichier C6', row_c6)

    content_layout.addWidget(gb_file)

    # Group: GraceTHD
    gb_gthd, form_gthd = PageBuilder.create_groupbox('GraceTHD')

    # Répertoire GraceTHD (shapefiles)
    row_gthd, line_gthd, btn_gthd = PageBuilder.create_path_row('Répertoire GraceTHD (shapefiles)', 'folder')
    form_gthd.addRow('Répertoire', row_gthd)

    content_layout.addWidget(gb_gthd)

    # Group: Études COMAC - REQ-PLC6-005
    gb_comac, form_comac = PageBuilder.create_groupbox('Études COMAC')

    # REQ-PLC6-005: Chemin études COMAC
    row_comac, line_comac, btn_comac = PageBuilder.create_path_row('Répertoire études COMAC (ExportComac.xlsx + *.pcm)', 'folder')
    form_comac.addRow('Répertoire COMAC', row_comac)

    content_layout.addWidget(gb_comac)

    # Group: Couches
    gb_couches, form_couches = PageBuilder.create_groupbox('Couches')

    cb_bpe = QComboBox()
    cb_bpe.setMinimumWidth(250)
    form_couches.addRow('Couche BPE', cb_bpe)

    cb_attaches = QComboBox()
    cb_attaches.setMinimumWidth(250)
    form_couches.addRow('Couche Attaches', cb_attaches)

    content_layout.addWidget(gb_couches)

    # Group: Découpage - REQ-PLC6-004
    gb_dec, form_dec = PageBuilder.create_groupbox('Zone de découpage')

    # REQ-PLC6-004: QgsMapLayerComboBox pour zone découpage
    cb_zone_decoupage = PageBuilder.create_layer_combo()
    cb_zone_decoupage.setFilters(QgsMapLayerProxyModel.PolygonLayer)
    form_dec.addRow('Couche zone', cb_zone_decoupage)

    cb_etudes = QComboBox()
    cb_etudes.setMinimumWidth(250)
    form_dec.addRow('Couche études', cb_etudes)

    cb_col = QComboBox()
    cb_col.setMinimumWidth(250)
    form_dec.addRow('Filtre Colonne', cb_col)

    cb_val = QComboBox()
    cb_val.setMinimumWidth(250)
    form_dec.addRow('Filtre Valeur', cb_val)

    content_layout.addWidget(gb_dec)
    content_layout.addStretch()

    main_layout.addWidget(scroll)

    # Footer
    footer, footer_widgets = PageBuilder.create_footer('Exécuter')
    main_layout.addWidget(footer)

    page.widgets = {
        'C6LienCheminImportFichier': line_c6,
        'c6BoutonCheminImport': btn_c6,
        'c6LienCheminGraceThd': line_gthd,
        'boutonCheminGraceThd': btn_gthd,
        # REQ-PLC6-005: Études COMAC
        'c6LienCheminComac': line_comac,
        'c6BoutonCheminComac': btn_comac,
        # Couches
        'c6ComboBoxCoucheBpe': cb_bpe,
        'c6ComboBoxCoucheAttaches': cb_attaches,
        # REQ-PLC6-004: Zone découpage
        'c6ComboBoxZoneDecoupage': cb_zone_decoupage,
        'c6ComboBoxCoucheEtudes': cb_etudes,
        'c6comboBoxColonneDecoupage': cb_col,
        'c6ComboBoxValeur': cb_val,
        'c6Lanceur': footer_widgets['btn'],
        'progressC6': footer_widgets['progress'],
        'kpis': kpi_widgets,
    }

    return page


def build_page_c6_c3a_bd(parent=None):
    """Build C6-C3A-C7-BD page."""
    page = QWidget(parent)
    main_layout = QVBoxLayout(page)
    main_layout.setSpacing(0)
    main_layout.setContentsMargins(0, 0, 0, 0)

    # Header
    kpis = {
        'total': ('—', 'kpi_default'),
        'ok': ('OK: —', 'kpi_ok'),
        'manquant': ('Manquant: —', 'kpi_warn'),
        'diff': ('Différent: —', 'kpi_error'),
    }
    header, kpi_widgets = PageBuilder.create_header(
        'C6 vs C3A vs C7 vs BD',
        'Analyse croisée des annexes et base de données',
        kpis
    )
    main_layout.addWidget(header)

    # Content
    scroll, content_layout = PageBuilder.create_scroll_content()

    # Group: Couches
    gb_couches, form_couches = PageBuilder.create_groupbox('Couches sources')

    cb_infra = QComboBox()
    cb_infra.setMinimumWidth(250)
    form_couches.addRow('Couche Poteaux', cb_infra)

    row_dec = QHBoxLayout()
    cb_dec = QComboBox()
    cb_dec.setMinimumWidth(120)
    row_dec.addWidget(cb_dec)
    cb_dec_champs = QComboBox()
    cb_dec_champs.setMinimumWidth(120)
    row_dec.addWidget(cb_dec_champs)
    cb_dec_val = QComboBox()
    cb_dec_val.setMinimumWidth(120)
    row_dec.addWidget(cb_dec_val)
    form_couches.addRow('Zone découpage', row_dec)

    content_layout.addWidget(gb_couches)

    # Group: Fichiers
    gb_files, form_files = PageBuilder.create_groupbox('Fichiers annexes')

    row_c6, line_c6, btn_c6 = PageBuilder.create_path_row('Fichier C6 (.xlsx)', 'excel')
    form_files.addRow('Fichier C6', row_c6)

    row_c7, line_c7, btn_c7 = PageBuilder.create_path_row('Fichier C7 (.xlsx)', 'excel')
    form_files.addRow('Fichier C7', row_c7)

    # Mode radio
    mode_layout = QHBoxLayout()
    rb_qgis = QRadioButton('QGIS')
    rb_qgis.setChecked(True)
    mode_layout.addWidget(rb_qgis)
    rb_excel = QRadioButton('EXCEL')
    mode_layout.addWidget(rb_excel)
    mode_layout.addStretch()
    form_files.addRow('Source C3A', mode_layout)

    cb_c3a = QComboBox()
    cb_c3a.setMinimumWidth(250)
    form_files.addRow('Couche C3A', cb_c3a)

    row_c3a, line_c3a, btn_c3a = PageBuilder.create_path_row('Fichier C3A (.xlsx)', 'excel')
    form_files.addRow('Fichier C3A', row_c3a)

    content_layout.addWidget(gb_files)

    # Group: Export
    gb_exp, form_exp = PageBuilder.create_groupbox('Export')

    row_exp, line_exp, btn_exp = PageBuilder.create_path_row('Dossier d\'export', 'export')
    form_exp.addRow('Répertoire Export', row_exp)

    content_layout.addWidget(gb_exp)
    content_layout.addStretch()

    main_layout.addWidget(scroll)

    # Footer
    footer, footer_widgets = PageBuilder.create_footer('Exécuter')
    main_layout.addWidget(footer)

    page.widgets = {
        'comboBox_infra_pt_pot_c6_c3a_bd': cb_infra,
        'comboBox_Decoupage': cb_dec,
        'comboBox_Dcp_champs': cb_dec_champs,
        'comboBox_Dcp_Valeur_champs': cb_dec_val,
        'lienCheminC6_c6_c3a_bd': line_c6,
        'boutonCheminC6_c6_c3a_bd': btn_c6,
        'lienCheminC7_c6_c3a_bd': line_c7,
        'boutonCheminC7_c6_c3a_bd': btn_c7,
        'radioButtonQgis': rb_qgis,
        'radioButtonExcel': rb_excel,
        'comboBox_Cmd_c6_c3a_bd': cb_c3a,
        'lienCheminC3A_c6_c3a_bd': line_c3a,
        'boutonCheminC3A_c6_c3a_bd': btn_c3a,
        'lienCheminExportDonnees_c6_c3a_bd': line_exp,
        'boutonCheminExportDonnees_c6_c3a_bd': btn_exp,
        'c6_c3a_bdLanceur': footer_widgets['btn'],
        'progressC6C3A': footer_widgets['progress'],
        'kpis': kpi_widgets,
    }

    return page


# Page builders list (order matches sidebar)
PAGE_BUILDERS = [
    build_page_maj,
    build_page_c6bd,
    build_page_capft,
    build_page_comac,
    build_page_police_c6,
    build_page_c6_c3a_bd,
]
