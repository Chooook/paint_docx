"""Модуль с функциями, которые изменяют структуру объекта Document."""

from copy import deepcopy
from typing import Tuple

from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .static import Index


def allocate_run_with_text(paragraph: Paragraph, run: Run, text: str) -> Run:
    """Функция для выделения объекта run, содержащего необходимый текст.

    Разделяет исходный Run на 3 Run`а для отделения Run`а с текстом.
    Перезаписывает весь параграф.
    Метод paragraph.append_runs добавляет Run с пробелом в начало,
    эта функция очищает Run с пробелом для сохранения структуры параграфа.
    После разделения все три Run`а сохраняют стиль исходного.
    Изменяет исходный объект Document.

    :param paragraph: Paragraph, содержащий необходимый Run.
    :param run: Run, который необходимо разделить.
    :param text: Текст, который необходимо выделить в отдельный Run.
    :return: Run, содержащий только необходимый текст.
    """
    runs = paragraph.runs
    try:
        run_index = [r.text for r in runs].index(run.text)
    except ValueError:
        # FIXME возможно неправильное определение индекса в случае идентичных
        run_index = [r.text for r in runs].index(text)
    new_runs, run_with_text = __split_run(run, text)
    paragraph.clear()
    paragraph.append_runs(runs[:run_index] + new_runs + runs[run_index + 1:])
    paragraph.runs[Index.first].clear()
    return run_with_text


def __split_run(run: Run, text: str) -> Tuple[list[Run], Run]:
    first_r = deepcopy(run)
    second_r = deepcopy(run)
    third_r = run
    first_r.text, third_r.text = run.text.split(text, maxsplit=1)
    second_r.text = text
    return [first_r, second_r, third_r], second_r
