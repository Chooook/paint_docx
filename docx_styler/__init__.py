"""Реализация стайлера для .docx файлов в виде ящика с инструментами.

(пока только одна отвёртка и та шлицевая)
TODO:
    * Покраска в таблицах
    * Не красить слово по частям, когда оно в нескольких run,
        а вытаскивать все его части в один run
    * Добавление комментариев по тексту
    * Расширенная работа со стилями (шрифт, размер, написание, ...)
"""

from .change import add_runs, allocate_run_with_text
from .color import color_run
from .main import color_text
from .search import (check_text_in_element, get_paragraphs_with_text,
                     get_runs_with_text)

__all__ = (
    'add_runs', 'allocate_run_with_text',
    'color_run',
    'color_text',
    'check_text_in_element', 'get_paragraphs_with_text', 'get_runs_with_text',
)
