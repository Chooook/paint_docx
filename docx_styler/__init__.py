"""Реализация стайлера для .docx файлов в виде ящика с инструментами.

(пока только одна отвёртка и несколько редких бит)
TODO:
    * Покраска в таблицах
    * Добавление комментариев по тексту
    * Расширенная работа со стилями (шрифт, размер, написание, ...)
"""

from .core import RunsMap, RunFinder, color_run, highlight_run
from .main import color_text, highlight_text

__all__ = (
    'RunsMap', 'RunFinder', 'color_run', 'highlight_run',
    'color_text', 'highlight_text',
)
