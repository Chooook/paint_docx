"""Модуль с внешними функциями для покраски Run в объекте Document."""

# TODO: Реализовать возможность работы с копией документа
#  вместо изменения на месте (нужно ли это?)
from typing import Tuple

from docx import Document

from .core import (
    color_run, highlight_run,
    RunFinder
)


def color_text(document: Document,
               text: str,
               color: str | Tuple[int, int, int],
               first_only: bool = False,
               ) -> None:
    """Функция для покраски частей текста в .docx.

    Не изменяет структуры и стилей.
    Покраска происходит на месте, не забудьте сохранить документ в файл.

    :param document: Экземпляр документа, который красим.
    :param text: Строка текста, которую нужно покрасить.
    :param color: Цвет (из класса Color), в который хотим покрасить.
    :param first_only: Флаг для покраски только первого вхождения.
    """
    rf = RunFinder(document)
    for run in rf.get_runs_with_text_from_document(text, first_only):
        color_run(run, color)


def highlight_text(document: Document,
                   text: str,
                   color: str,
                   first_only: bool = False,
                   ) -> None:
    """Функция для покраски частей текста в .docx.

    Не изменяет структуры и стилей.
    Покраска происходит на месте, не забудьте сохранить документ в файл.

    :param document: Экземпляр документа, который красим.
    :param text: Строка текста, которую нужно покрасить.
    :param color: Цвет (из класса Color), в который хотим покрасить.
    :param first_only: Флаг для покраски только первого вхождения.
    """
    rf = RunFinder(document)
    for run in rf.get_runs_with_text_from_document(text, first_only):
        highlight_run(run, color)
