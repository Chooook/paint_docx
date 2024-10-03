"""Модуль с функциями для изменения элементов объекта Document."""
from typing import List

from docx import Document
from docx.text.run import Run

from .core import (
    Color,
    get_paragraphs_with_text,
    get_runs_with_text
)


def color_text(document: Document,
               text: str,
               first_only: bool = False,
               color: str = 'red'
               ) -> None:
    """Функция для покраски частей текста в .docx.

    Не изменяет структуры и стилей.
    Покраска происходит на месте, не забудьте сохранить документ в файл.

    :param document: Экземпляр документа, который красим.
    :param text: Строка текста, которую нужно покрасить.
    :param color: Цвет (из класса Color), в который хотим покрасить.
    :param first_only: Флаг для покраски только первого вхождения.
    """
    for run in __get_runs_with_text_from_document(document, text, first_only):
        Color.color_run(run, color)


def highlight_text(document: Document,
                   text: str,
                   first_only: bool = False,
                   color: str = 'red'
                   ) -> None:
    """Функция для покраски частей текста в .docx.

    Не изменяет структуры и стилей.
    Покраска происходит на месте, не забудьте сохранить документ в файл.

    :param document: Экземпляр документа, который красим.
    :param text: Строка текста, которую нужно покрасить.
    :param color: Цвет (из класса Color), в который хотим покрасить.
    :param first_only: Флаг для покраски только первого вхождения.
    """
    for run in __get_runs_with_text_from_document(document, text, first_only):
        Color.highlight_run(run, color)


def __get_runs_with_text_from_document(document: Document,
                                       text: str,
                                       first_only: bool
                                       ) -> List[Run]:
    text = text.strip()
    runs_to_color = []
    for paragraph in get_paragraphs_with_text(document, text, first_only):
        for runs_list in get_runs_with_text(
                paragraph, text, first_only=first_only):
            for run in runs_list:
                runs_to_color.append(run)
    return runs_to_color
