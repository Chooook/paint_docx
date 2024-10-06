"""Модуль с функциями для поиска элементов объекта Document по тексту."""
# TODO: реализовать итеративную покраску не всех объектов, соответствующих
#  передаваемому тексту, а только тех, которые необходимо покрасить в этом
#  тексте, например:
#  Текст "Просто пример текста";
#  Покрасить нужно слова "Просто" и "текста";
#  В этом случае необходимо найти весь переданный текст
#  и уже в нём искать слова, которые нужно покрасить.
#  Необходимо также продумать структуру данных, соответствующую этой концепции
#  и сохранить обратную совместимость с текущей реализацией.
#  Для этого нужно выделить все run, соответствующие тексту и искать по ним.
#  Можно создать перегрузку только входной функции, так как дальше она
#  будет передавать работу в другие функции, в другие функции можно
#  просто добавить необязательный параметр. (?)

from typing import Generator, List, Tuple

from docx.text.paragraph import Paragraph
from docx.text.run import Run

from docx import Document

from .structure import allocate_run_with_text
from .utils import FIRST


def get_paragraphs_with_text(document: Document,
                             text: str,
                             first_only: bool = False
                             ) -> List[Paragraph]:
    """Функция для поиска объектов Paragraph, содержащих text.

    :param document: Объект Document, в котором осуществляется поиск.
    :param text: Искомый текст.
    :param first_only:
        True - возвращается список с первым соответствующим Paragraph.
        False - возвращается список со всеми соответствующими Paragraph.
    :return: Список объектов Paragraph, содержащих text.
    """
    paragraphs = []
    for paragraph in document.paragraphs:
        if check_text_in_element(paragraph, text, strict=False):
            paragraphs.append(paragraph)
        if first_only:
            return paragraphs
    return paragraphs


def check_text_in_element(element: Run | Paragraph,
                          text: str,
                          strict: bool = False
                          ) -> bool:
    """Функция для проверки объекта на содержание text.

    :param element: Проверяемый элемент.
    :param text: Искомый текст.
    :param strict:
        True - проверка объекта на полное вхождение text.
        False - проверка объекта на частичное вхождение text.
    :return: Bool, означающий, содержит объект text или нет.
    """
    if strict:
        return text == element.text.strip()
    return text in element.text


def get_runs_with_text(paragraph: Paragraph,
                       text: str,
                       first_only: bool = False,
                       ) -> List[List[Run]]:
    """Функция для поиска объектов Run, содержащих text.

    :param paragraph: Paragraph, в котором осуществляется поиск.
    :param text: Искомый текст.
    :param first_only:
        True - возвращается список с первым соответствующим Run.
        False - возвращается список со всеми соответствующими Run.
    :return:  Список объектов Run, содержащих text.
    """
    # TODO Использует модуль structure, неправильная зависимость,
    #  подумать как изменить
    runs = []
    possible_runs = list(__find_text_in_runs(paragraph.runs, text))
    for i, possible_run in enumerate(possible_runs):
        run, text_part = possible_run
        if check_text_in_element(run, text, strict=True):
            runs.append([run])
            if first_only:
                return runs
        else:
            temp_runs = []
            temp_text = []
            for temp_possible_run in possible_runs[i:]:
                temp_run, temp_text_part = temp_possible_run
                if check_text_in_element(
                        temp_run, temp_text_part, strict=True):
                    temp_runs.append(temp_run)
                    temp_text.append(temp_text_part)
                else:
                    temp_runs.append(allocate_run_with_text(
                        paragraph, temp_run, temp_text_part))
                    temp_text.append(temp_text_part)
                if ''.join(temp_text).strip() == text:
                    runs.append(temp_runs)
                    if first_only:
                        return runs
                    break
                elif ''.join(temp_text) in text:
                    continue
                else:
                    break
    return runs


def __find_text_in_runs(runs: List[Run],
                        text: str
                        ) -> Generator[Tuple[Run, str], None, None]:
    # FIXME красит лишнее если run заканчивается, пара букв в него попала,
    #  но в следующем run нет продолжения. Безумно редкий случай,
    #  скорее всего, можно создать только искусственно (см. template.docx)
    #  решение в заметке в модуле main

    text_symbols = list(text)
    for run in runs:
        run_contains: List[str] = []
        for run_symbol in run.text:
            try:
                symbol = text_symbols.pop(FIRST)
                if run_symbol != symbol:
                    run_contains.clear()
                    text_symbols = list(text)
                else:
                    run_contains.append(symbol)
            except IndexError:
                if run_contains:
                    yield run, ''.join(run_contains)
                run_contains.clear()
                text_symbols = list(text)
                continue
        if run_contains:
            yield run, ''.join(run_contains)
