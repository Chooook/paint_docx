"""Пакет с низкоуровневыми функциями для работы с объектом Document."""
from .structure import allocate_run_with_text
from .color import color_run, fill_run_with_color
from .search import (check_text_in_element, get_paragraphs_with_text,
                     get_runs_with_text)
from extensions import TextHighlight, SpecifiedTextHighlight

__all__ = (
    'allocate_run_with_text',
    'color_run', 'fill_run_with_color',
    'check_text_in_element', 'get_paragraphs_with_text', 'get_runs_with_text',
    'TextHighlight', 'SpecifiedTextHighlight'
)
