"""Модуль с функциями, которые изменяют изначальный документ."""

from copy import deepcopy

from docx.text.paragraph import Paragraph
from docx.text.run import Run


def allocate_run_with_text(paragraph: Paragraph, run: Run, text: str) -> Run:
    """Функция для выделения объекта run, содержащего необходимый текст.

    Разделяет исходный Run на 3 Run`а для отделения Run`а с текстом.
    Перезаписывает весь параграф.
    После разделения все три Run`а сохраняют стиль исходного.
    Изменяет исходный объект Document.

    :param paragraph: Paragraph, содержащий необходимый Run.
    :param run: Run, который необходимо разделить.
    :param text: Текст, который необходимо выделить в отдельный Run.
    :return: Run, содержащий только необходимый текст.
    """
    run_with_text_after_split_index = 1
    runs = paragraph.runs
    run_index = [r.text for r in runs].index(run.text)
    new_runs = __split_run(run, text)
    paragraph.clear()
    add_runs(paragraph, runs[:run_index] + new_runs + runs[run_index + 1:])
    return new_runs[run_with_text_after_split_index]


def add_runs(paragraph: Paragraph, runs: list[Run]) -> None:
    """Функция для добавления списка Run`ов в параграф.

    Исходный метод append_runs добавляет Run с пробелом в начало, эта
    функция очищает Run с пробелом для сохранения структуры параграфа.
    Изменяет исходный объект Document.

    :param paragraph: Paragraph, в который нужно добавить Run`ы.
    :param runs: Run`ы, которые нужно добавить.
    """
    runs_number = len(paragraph.runs)
    paragraph.append_runs(runs)
    paragraph.runs[runs_number].clear()


def __split_run(run: Run, text: str) -> list[Run]:
    first_r = deepcopy(run)
    second_r = deepcopy(run)
    third_r = deepcopy(run)
    first_r.text, third_r.text = run.text.split(text, maxsplit=1)
    second_r.text = text
    return [first_r, second_r, third_r]
