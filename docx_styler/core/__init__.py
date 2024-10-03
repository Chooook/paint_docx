"""Пакет с низкоуровневыми функциями для работы с объектом Document."""
from .extensions import SpecifiedTextHighlight, TextHighlight
from .color import Color
from .search import (check_text_in_element, get_paragraphs_with_text,
                     get_runs_with_text)
from .structure import allocate_run_with_text

__all__ = (
    'TextHighlight', 'SpecifiedTextHighlight',
    'Color',
    'check_text_in_element', 'get_paragraphs_with_text', 'get_runs_with_text',
    'allocate_run_with_text',
)
