"""Модуль с низкоуровневыми функциями для работы с объектом Document."""
from .color import color_run, highlight_run
from .extensions import SpecifiedTextToHighlight, TextToHighlight
from .run_collector import RunCollector
from .search import RunFinder

__all__ = (
    'color_run', 'highlight_run',
    'TextToHighlight', 'SpecifiedTextToHighlight',
    'RunCollector',
    'RunFinder',
)
