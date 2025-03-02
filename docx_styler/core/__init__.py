"""Модуль с низкоуровневыми функциями для работы с объектом Document."""
from .color import color_run, highlight_run
from .extensions import HighlightPattern
from .runs_map import RunsMap
from .search import RunFinder

__all__ = (
    'color_run', 'highlight_run',
    'HighlightPattern',
    'RunsMap',
    'RunFinder',
)
