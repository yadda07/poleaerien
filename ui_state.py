# -*- coding: utf-8 -*-
"""
UI State Manager for PoleAerien plugin.
Centralized state management for all module pages.
"""

import os
from enum import Enum
from qgis.PyQt.QtCore import QObject, pyqtSignal


class PageState(Enum):
    """Page states matching the design system."""
    EMPTY = 'empty'
    READY = 'ready'
    RUNNING = 'running'
    DONE = 'done'
    ERROR = 'error'


class ModuleStateManager(QObject):
    """Manages state for a single module page."""
    
    state_changed = pyqtSignal(str)  # Emits state name
    kpi_updated = pyqtSignal(str, str)  # Emits (kpi_name, value)
    
    def __init__(self, page_index, dialog, parent=None):
        super().__init__(parent)
        self.page_index = page_index
        self.dialog = dialog
        self._state = PageState.EMPTY
        self._prerequisites_met = False
    
    @property
    def state(self):
        return self._state
    
    def set_state(self, state):
        """Set page state and update UI."""
        if isinstance(state, str):
            state = PageState(state)
        self._state = state
        self.dialog.set_page_status(self.page_index, state.value)
        self.state_changed.emit(state.value)
    
    def set_empty(self):
        self.set_state(PageState.EMPTY)
    
    def set_ready(self):
        self.set_state(PageState.READY)
    
    def set_running(self):
        self.set_state(PageState.RUNNING)
    
    def set_done(self):
        self.set_state(PageState.DONE)
    
    def set_error(self):
        self.set_state(PageState.ERROR)
    
    def update_kpi(self, name, value):
        """Update a KPI badge."""
        self.dialog.update_kpi(self.page_index, name, value)
        self.kpi_updated.emit(name, str(value))
    
    def check_prerequisites(self, conditions):
        """Check if all prerequisites are met.
        
        Args:
            conditions: dict of {name: bool} prerequisites
            
        Returns:
            True if all conditions are met
        """
        self._prerequisites_met = all(conditions.values())
        if self._prerequisites_met:
            self.set_ready()
        else:
            self.set_empty()
        return self._prerequisites_met
    
    def enable_action(self, btn_name):
        """Enable the action button."""
        btn = getattr(self.dialog, btn_name, None)
        if btn:
            btn.setEnabled(True)
    
    def disable_action(self, btn_name):
        """Disable the action button."""
        btn = getattr(self.dialog, btn_name, None)
        if btn:
            btn.setEnabled(False)


class UIStateController:
    """Central controller for all module states."""
    
    # Module indices
    MAJ = 0
    C6BD = 1
    CAPFT = 2
    COMAC = 3
    POLICE_C6 = 4
    C6_C3A_BD = 5
    
    # Action button names per module
    ACTION_BUTTONS = {
        0: 'majBdLanceur',
        1: 'C6BdLanceur',
        2: 'cap_ftLanceur',
        3: 'cap_comacLanceur',
        4: 'c6Lanceur',
        5: 'c6_c3a_bdLanceur',
    }
    
    def __init__(self, dialog):
        self.dialog = dialog
        self.modules = {}
        self._init_modules()
    
    def _init_modules(self):
        """Initialize state managers for all modules."""
        for idx in range(6):
            self.modules[idx] = ModuleStateManager(idx, self.dialog)
    
    def get_module(self, index):
        """Get state manager for a module."""
        return self.modules.get(index)
    
    def validate_prerequisites_maj(self):
        """Validate MAJ BD prerequisites."""
        dlg = self.dialog
        file_path = dlg.MajFileWidget.filePath() if hasattr(dlg, 'MajFileWidget') else ''
        conditions = {
            'layer_infra': dlg.MajcomboBox_infra_pt_pot.currentLayer() is not None,
            'layer_etude_ft': dlg.MajcomboBox_etude_cap_ft.currentLayer() is not None,
            'layer_etude_comac': dlg.MajcomboBox_etude_comac.currentLayer() is not None,
            'file': bool(file_path),
            'file_exists': os.path.isfile(file_path or ''),
        }
        
        mgr = self.modules[self.MAJ]
        if mgr.check_prerequisites(conditions):
            mgr.enable_action(self.ACTION_BUTTONS[self.MAJ])
        else:
            mgr.disable_action(self.ACTION_BUTTONS[self.MAJ])
        
        return conditions
    
    def validate_prerequisites_c6bd(self):
        """Validate C6 vs BD prerequisites."""
        dlg = self.dialog
        path_c6 = dlg.lienCheminFichiersC6.text()
        path_export = dlg.lienCheminExportDonnees.text()
        layer_infra = dlg.C6BdcomboBox_infra_pt_pot.currentLayer()
        layer_etude = dlg.C6BdcomboBox_etude_cap_ft.currentLayer()
        layer_diff = True
        if layer_infra is not None and layer_etude is not None:
            layer_diff = layer_infra.id() != layer_etude.id()
        conditions = {
            'layer_infra': layer_infra is not None,
            'layer_etude': layer_etude is not None,
            'layer_diff': layer_diff,
            'path_c6': bool(path_c6),
            'path_c6_exists': os.path.isdir(path_c6 or ''),
            'path_export': bool(path_export),
            'path_export_exists': os.path.isdir(path_export or ''),
        }
        
        mgr = self.modules[self.C6BD]
        if mgr.check_prerequisites(conditions):
            mgr.enable_action(self.ACTION_BUTTONS[self.C6BD])
        else:
            mgr.disable_action(self.ACTION_BUTTONS[self.C6BD])
        
        return conditions
    
    def validate_prerequisites_capft(self):
        """Validate CAP_FT prerequisites."""
        dlg = self.dialog
        path_capft = dlg.lienRepertoireCapFt.text()
        path_export = dlg.lienCheminExportCapFt.text()
        conditions = {
            'layer_infra': dlg.capFtComboBoxCoucheInfra_pt_pot.currentLayer() is not None,
            'layer_etude': dlg.capFtComboBox_etude_cap_ft.currentLayer() is not None,
            'field_etude': bool(dlg.capFtComboBoxChampsCapFt.currentText()),
            'path_capft': bool(path_capft),
            'path_capft_exists': os.path.isdir(path_capft or ''),
            'path_export': bool(path_export),
            'path_export_exists': os.path.isdir(path_export or ''),
        }
        
        mgr = self.modules[self.CAPFT]
        if mgr.check_prerequisites(conditions):
            mgr.enable_action(self.ACTION_BUTTONS[self.CAPFT])
        else:
            mgr.disable_action(self.ACTION_BUTTONS[self.CAPFT])
        
        return conditions
    
    def validate_prerequisites_comac(self):
        """Validate COMAC prerequisites."""
        dlg = self.dialog
        path_comac = dlg.lienRepertoireComac.text()
        path_export = dlg.lienCheminExportComac.text()
        conditions = {
            'layer_infra': dlg.comacComboBoxCoucheInfra_pt_pot.currentLayer() is not None,
            'layer_etude': dlg.comboBoxCoucheComac.currentLayer() is not None,
            'field_etude': bool(dlg.comboBoxChampsComac.currentText()),
            'path_comac': bool(path_comac),
            'path_comac_exists': os.path.isdir(path_comac or ''),
            'path_export': bool(path_export),
            'path_export_exists': os.path.isdir(path_export or ''),
        }
        
        mgr = self.modules[self.COMAC]
        if mgr.check_prerequisites(conditions):
            mgr.enable_action(self.ACTION_BUTTONS[self.COMAC])
        else:
            mgr.disable_action(self.ACTION_BUTTONS[self.COMAC])
        
        return conditions
    
    def validate_prerequisites_police_c6(self):
        """Validate Police C6 prerequisites."""
        dlg = self.dialog
        conditions = {
            'file_c6': bool(dlg.C6LienCheminImportFichier.text()),
        }
        
        mgr = self.modules[self.POLICE_C6]
        if mgr.check_prerequisites(conditions):
            mgr.enable_action(self.ACTION_BUTTONS[self.POLICE_C6])
        else:
            mgr.disable_action(self.ACTION_BUTTONS[self.POLICE_C6])
        
        return conditions
    
    def validate_prerequisites_c6_c3a_bd(self):
        """Validate C6-C3A-C7-BD prerequisites."""
        dlg = self.dialog
        path_c6 = dlg.lienCheminC6_c6_c3a_bd.text()
        path_c7 = dlg.lienCheminC7_c6_c3a_bd.text()
        path_c3a = dlg.lienCheminC3A_c6_c3a_bd.text()
        path_export = dlg.lienCheminExportDonnees_c6_c3a_bd.text()
        extra_ok = True
        if dlg.comboBox_infra_pt_pot_c6_c3a_bd.currentText() == dlg.comboBox_Decoupage.currentText():
            extra_ok = False
        if dlg.radioButtonQgis.isChecked():
            extra_ok = extra_ok and bool(dlg.comboBox_Cmd_c6_c3a_bd.currentText())
        if dlg.radioButtonExcel.isChecked():
            extra_ok = extra_ok and os.path.exists(path_c3a or '')
        conditions = {
            'layer_infra': bool(dlg.comboBox_infra_pt_pot_c6_c3a_bd.currentText()),
            'layer_decoupage': bool(dlg.comboBox_Decoupage.currentText()),
            'field_decoupage': bool(dlg.comboBox_Dcp_champs.currentText()),
            'value_decoupage': bool(dlg.comboBox_Dcp_Valeur_champs.currentText()),
            'file_c6': bool(path_c6),
            'file_c6_exists': os.path.exists(path_c6 or ''),
            'file_c7': bool(path_c7),
            'file_c7_exists': os.path.exists(path_c7 or ''),
            'path_export': bool(path_export),
            'path_export_exists': os.path.isdir(path_export or ''),
            'extra_ok': extra_ok,
        }
        
        mgr = self.modules[self.C6_C3A_BD]
        if mgr.check_prerequisites(conditions):
            mgr.enable_action(self.ACTION_BUTTONS[self.C6_C3A_BD])
        else:
            mgr.disable_action(self.ACTION_BUTTONS[self.C6_C3A_BD])
        
        return conditions
    
    def connect_validation_signals(self):
        """Connect all widgets to validation methods."""
        dlg = self.dialog
        
        # MAJ
        dlg.MajcomboBox_infra_pt_pot.layerChanged.connect(self.validate_prerequisites_maj)
        dlg.MajcomboBox_etude_cap_ft.layerChanged.connect(self.validate_prerequisites_maj)
        dlg.MajcomboBox_etude_comac.layerChanged.connect(self.validate_prerequisites_maj)
        dlg.MajFileWidget.fileChanged.connect(self.validate_prerequisites_maj)
        
        # C6 vs BD
        dlg.C6BdcomboBox_infra_pt_pot.layerChanged.connect(self.validate_prerequisites_c6bd)
        dlg.C6BdcomboBox_etude_cap_ft.layerChanged.connect(self.validate_prerequisites_c6bd)
        dlg.lienCheminFichiersC6.textChanged.connect(self.validate_prerequisites_c6bd)
        dlg.lienCheminExportDonnees.textChanged.connect(self.validate_prerequisites_c6bd)
        
        # CAP_FT
        dlg.capFtComboBoxCoucheInfra_pt_pot.layerChanged.connect(self.validate_prerequisites_capft)
        dlg.capFtComboBox_etude_cap_ft.layerChanged.connect(self.validate_prerequisites_capft)
        dlg.capFtComboBoxChampsCapFt.currentTextChanged.connect(self.validate_prerequisites_capft)
        dlg.lienRepertoireCapFt.textChanged.connect(self.validate_prerequisites_capft)
        dlg.lienCheminExportCapFt.textChanged.connect(self.validate_prerequisites_capft)
        
        # COMAC
        dlg.comacComboBoxCoucheInfra_pt_pot.layerChanged.connect(self.validate_prerequisites_comac)
        dlg.comboBoxCoucheComac.layerChanged.connect(self.validate_prerequisites_comac)
        dlg.comboBoxChampsComac.currentTextChanged.connect(self.validate_prerequisites_comac)
        dlg.lienRepertoireComac.textChanged.connect(self.validate_prerequisites_comac)
        dlg.lienCheminExportComac.textChanged.connect(self.validate_prerequisites_comac)
        
        # C6-C3A-BD
        dlg.comboBox_infra_pt_pot_c6_c3a_bd.currentTextChanged.connect(self.validate_prerequisites_c6_c3a_bd)
        dlg.comboBox_Decoupage.currentTextChanged.connect(self.validate_prerequisites_c6_c3a_bd)
        dlg.comboBox_Dcp_champs.currentTextChanged.connect(self.validate_prerequisites_c6_c3a_bd)
        dlg.comboBox_Dcp_Valeur_champs.currentTextChanged.connect(self.validate_prerequisites_c6_c3a_bd)
        dlg.comboBox_Cmd_c6_c3a_bd.currentTextChanged.connect(self.validate_prerequisites_c6_c3a_bd)
        dlg.radioButtonQgis.clicked.connect(self.validate_prerequisites_c6_c3a_bd)
        dlg.radioButtonExcel.clicked.connect(self.validate_prerequisites_c6_c3a_bd)
        dlg.lienCheminC6_c6_c3a_bd.textChanged.connect(self.validate_prerequisites_c6_c3a_bd)
        dlg.lienCheminC7_c6_c3a_bd.textChanged.connect(self.validate_prerequisites_c6_c3a_bd)
        dlg.lienCheminC3A_c6_c3a_bd.textChanged.connect(self.validate_prerequisites_c6_c3a_bd)
        dlg.lienCheminExportDonnees_c6_c3a_bd.textChanged.connect(self.validate_prerequisites_c6_c3a_bd)
