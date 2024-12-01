"""Модуль с низкоуровневыми функциями для работы с объектом Document."""
from .extensions import SpecifiedTextToHighlight, TextToHighlight
from .color import color_run, highlight_run
from .search import (check_text_in_element, get_paragraphs_with_text,
                     get_runs_with_text_from_paragraph,
                     get_runs_with_text_from_document)

__all__ = (
    'TextToHighlight', 'SpecifiedTextToHighlight',
    'color_run', 'highlight_run',
    'check_text_in_element', 'get_paragraphs_with_text',
    'get_runs_with_text_from_paragraph',
    'get_runs_with_text_from_document',
)
