# -*- coding: utf-8 -*-
"""Package workflows - Orchestrateurs pour les modules métier."""

from .maj_workflow import MajWorkflow
from .comac_workflow import ComacWorkflow
from .capft_workflow import CapFtWorkflow
from .c6bd_workflow import C6BdWorkflow
from .c6c3a_workflow import C6C3AWorkflow
from .police_workflow import PoliceWorkflow

__all__ = [
    'MajWorkflow',
    'ComacWorkflow',
    'CapFtWorkflow',
    'C6BdWorkflow',
    'C6C3AWorkflow',
    'PoliceWorkflow'
]
