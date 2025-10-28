"""
Компоненты для работы с КОМПАС-3D
"""
from .base_component import BaseKompasComponent
from .project_copier import ProjectCopier
from .cascading_variables_updater import CascadingVariablesUpdater
from .bmp_organizer import BmpOrganizer
from .template_manager import TemplateManager

__all__ = [
    'BaseKompasComponent',
    'ProjectCopier',
    'CascadingVariablesUpdater',
    'BmpOrganizer',
    'TemplateManager',
]

