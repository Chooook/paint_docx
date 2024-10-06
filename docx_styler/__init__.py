"""Реализация стайлера для .docx файлов в виде ящика с инструментами.

(пока только одна отвёртка и несколько редких бит)
TODO:
    * Покраска в таблицах
    * Добавление комментариев по тексту
    * Расширенная работа со стилями (шрифт, размер, написание, ...)
"""

from .core import color_run, highlight_run
from .core import allocate_run_with_text
from .core import (check_text_in_element, get_paragraphs_with_text,
                   get_runs_with_text)
from .main import color_text, highlight_text

__all__ = (
    'color_run', 'highlight_run',
    'allocate_run_with_text',
    'check_text_in_element', 'get_paragraphs_with_text', 'get_runs_with_text',
    'color_text', 'highlight_text',
)
