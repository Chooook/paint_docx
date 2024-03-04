"""Пакет с низкоуровневыми функциями для работы с объектом Document."""
from .structure import add_runs, allocate_run_with_text
from .color import color_run
from .search import (check_text_in_element, get_paragraphs_with_text,
                     get_runs_with_text)

__all__ = (
    'add_runs', 'allocate_run_with_text',
    'color_run',
    'check_text_in_element', 'get_paragraphs_with_text', 'get_runs_with_text',
)
