"""Модуль с функциями для поиска элементов объекта Document по тексту."""
# TODO: реализовать итеративную покраску не всех объектов, соответствующих
#  передаваемому тексту, а только тех, которые необходимо покрасить в этом
#  тексте, например:
#  Текст "Просто пример текста";
#  Покрасить нужно слова "Просто" и "текста";
#  В этом случае необходимо найти весь переданный текст
#  и уже в нём искать слова, которые нужно покрасить.
#  Необходимо также продумать структуру данных, соответствующую этой концепции
#  и сохранить обратную совместимость с текущей реализацией (extensions.py).
#  Для этого нужно выделить все run, соответствующие тексту и искать по ним.
#  Можно создать перегрузку только входной функции, так как дальше она
#  будет передавать работу в другие функции, в другие функции можно
#  просто добавить необязательный параметр. (?)
# TODO: на данный момент не учитывается что фраза может быть в нескольких
#  параграфах, а это нужно как минимум для реализации поиска по таблицам.
# TODO: реализовать поиск по таблицам
from copy import deepcopy
from typing import Generator, List, Tuple

from docx.text.paragraph import Paragraph
from docx.text.run import Run

from docx import Document

from .utils import FIRST


class RunFinder:
    def __init__(self, document: Document):
        self.document = document

    def get_runs_with_text_from_document(self,
                                         text: str,
                                         first_only: bool
                                         ) -> List[Run]:
        """Функция для получения списка объектов Run с текстом.

        :param text: Искомый текст.
        :param first_only: Флаг для поиска только первого вхождения.
        :return:
        """
        text = text.strip()
        runs_to_color = []
        for paragraph in self.get_paragraphs_with_text(text, first_only):
            for runs_list in self.get_runs_with_text_from_paragraph(
                    paragraph, text, first_only=first_only):
                for run in runs_list:
                    runs_to_color.append(run)
        return runs_to_color

    def get_paragraphs_with_text(self,
                                 text: str,
                                 first_only: bool = False
                                 ) -> List[Paragraph]:
        """Ищет объекты Paragraph, содержащие text.

        :param text: Искомый текст.
        :param first_only:
            True - возвращается список с первым соответствующим Paragraph.
            False - возвращается список со всеми соответствующими Paragraph.
        :return: Список объектов Paragraph, содержащих text.
        """
        paragraphs = []
        for paragraph in self.document.paragraphs:
            if self.check_text_in_element(paragraph, text, strict=False):
                paragraphs.append(paragraph)
            if first_only:
                return paragraphs
        return paragraphs

    def get_runs_with_text_from_paragraph(self,
                                          paragraph: Paragraph,
                                          text: str,
                                          first_only: bool = False,
                                          ) -> List[List[Run]]:
        """Извлекает наборы объектов Run, которые в совокупности содержат text.

        :param paragraph: Paragraph, в котором осуществляется поиск.
        :param text: Искомый текст.
        :param first_only:
            True - возвращается список с первым соответствующим Run.
            False - возвращается список со всеми соответствующими Run.
        :return: Список наборов объектов Run, содержащих text.
        """
        runs = []
        possible_runs = list(self.__find_text_in_runs(paragraph.runs, text))
        for run_index, possible_run in enumerate(possible_runs):
            run, text_part = possible_run
            if self.check_text_in_element(run, text, strict=True):
                runs.append([run])
                if first_only:
                    return runs
            else:
                temp_runs = []
                temp_text = []
                for temp_possible_run in possible_runs[run_index:]:
                    temp_run, temp_text_part = temp_possible_run
                    if self.check_text_in_element(
                            temp_run, temp_text_part, strict=True):
                        temp_runs.append(temp_run)
                        temp_text.append(temp_text_part)
                    else:
                        temp_runs.append(self.__allocate_run_with_text(
                            temp_run, temp_text_part))
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

    @staticmethod
    def __find_text_in_runs(runs: List[Run],
                            text: str
                            ) -> Generator[Tuple[Run, str], None, None]:
        """Итеративно ищет объекты Run, которые содержат text или его часть.

        :param runs: Список объектов Run, по которым осуществляется поиск.
        :param text: Текст, по которому осуществляется поиск.
        :return: Кортеж с объектом Run, содержащим text или его часть
            и часть текста, которая была найдена.
        """
        text_symbols = list(text)
        for run in runs:
            if not run.text:
                continue
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

    @staticmethod
    def check_text_in_element(element: Run | Paragraph,
                              text: str,
                              strict: bool = False
                              ) -> bool:
        """Проверяет объект на содержание text.

        :param element: Проверяемый объект.
        :param text: Искомый текст.
        :param strict:
            True - проверка объекта на полное вхождение text.
            False - проверка объекта на наличие text в объекте.
        :return: Bool, Результат проверки.
        """
        if strict:
            # Run`ы часто содержат пробельные символы по краям,
            # которые не влияют на наличие/отсутствие искомого текста.
            # strip() применяется, чтобы не резать лишний раз структуру.
            return text == element.text.strip()
        return text in element.text

    def __allocate_run_with_text(
            self, run: Run, text: str) -> Run:
        """Выделяет объект Run, содержащий необходимый текст.

        Разделяет исходный Run на 3 Run`а для отделения Run`а с текстом.
        Перезаписывает весь параграф.
        После разделения все три Run`а сохраняют стиль исходного.
        Неявно изменяет объект Document.

        :param run: Run, который необходимо разделить.
        :param text: Текст, который необходимо выделить в отдельный Run.
        :return: Run, содержащий только необходимый текст.
        """
        paragraph = run._parent
        runs = paragraph.runs
        run_index = [r.text for r in runs].index(run.text)
        new_runs = self.__split_run(run, text)
        # Run с нужным текстом второй, см. __split_run
        run_with_text = new_runs[1]

        paragraph.clear()
        paragraph.append_runs(
            runs[:run_index] + new_runs + runs[run_index + 1:])
        # Очистка побочного Run`а с пробелом
        # для сохранения исходного текста параграфа
        paragraph.runs[FIRST].clear()
        return run_with_text

    @staticmethod
    def __split_run(run: Run, text: str) -> List[Run]:
        """Разделяет исходный Run на 3 Run`а для отделения Run`а с текстом.

        :param run: Исходный Run.
        :param text: Текст, который необходимо выделить в отдельный Run.
        :return: Набор объектов Run, в совокупности равные исходному Run.
        """
        first_r = deepcopy(run)
        second_r = deepcopy(run)
        third_r = run
        first_r.text, third_r.text = run.text.split(text, maxsplit=1)
        second_r.text = text
        return [first_r, second_r, third_r]
