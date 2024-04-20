"""Модуль с функциями для изменения элементов объекта Document."""
from typing import List

from docx import Document
from docx.text.run import Run

from .core import (
    color_run,
    fill_run_with_color,
    get_paragraphs_with_text,
    get_runs_with_text
)

# TODO передавать на покраску список ранов, но проверять сумму их текстов на
#  соответствие полному тексту, чтобы неполный текст не красился.
#  Вид списка на покраску: [[Run, Run, Run], [Run], [Run], [Run, Run]...]
#  Реализация должна быть в модуле поиска (search).


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


def fill_text_with_color(document: Document,
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
        fill_run_with_color(run, color)


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
