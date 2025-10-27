"""
Компоненты для работы с КОМПАС-3D
"""
from .base_component import BaseKompasComponent
from .project_copier import ProjectCopier
from .cascading_variables_updater import CascadingVariablesUpdater

__all__ = [
    'BaseKompasComponent',
    'ProjectCopier',
    'CascadingVariablesUpdater',
]

