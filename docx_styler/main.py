"""Модуль с функциями для изменения элементов объекта Document."""
from typing import List

from docx.text.run import Run

from docx import Document

from .core.color import color_run
from .core.search import get_paragraphs_with_text, get_runs_with_text


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
        color_run(run, color)


def __get_runs_with_text_from_document(document: Document,
                                       text: str,
                                       first_only: bool
                                       ) -> List[Run]:
    text = text.strip()
    runs = []
    for paragraph in get_paragraphs_with_text(document, text, first_only):
        for run in get_runs_with_text(
                paragraph, text, first_only=first_only):
            runs.append(run)
    return runs
