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

from .run_collector import RunCollector
from docx_styler.utils import FIRST


class RunFinder:
    def __init__(self, document: Document):
        self.run_collector = RunCollector(document)

    # TODO Совместить с методом ниже и переделать на генератор, т.к.
    #  в случае пересборки параграфа, все ссылки на объекты Run меняются,
    #  следовательно, RunCollector постоянно перезапускать,
    #  что накладно, но пока другого выхода не вижу,
    #  т.к. переделать в генератор не получится полностью.
    #  Можно попробовать реализовать обновление только отдельных
    #  параграфов в RunCollector, если ссылка на параграф не меняется,
    #  для этого нужно сохранять также параграфы,
    #  но итерацию совершать только по объектам Run, при этом
    #  нужно обновлять индекс, с которого начинается итерация,
    #  скорее всего, как в поисковике, так и в RunCollector.
    #  !!! ССЫЛКА СБИВАЕТСЯ ПОСТОЯННО, ДАЖЕ ПРИ ПОЛУЧЕНИИ ОБЪЕКТА RUN !!!
    #  ТАК ЧТО НЕТ СМЫСЛА ПЕРЕСОБИРАТЬ, Т.К. ВСЕ ОБРАЩЕНИЯ ИДУТ ПО ССЫЛКАМ
    def search_runs(self,
                    text: str,
                    first_only: bool = False
                    ) -> List[Run]:
        """Функция для получения списка объектов Run с текстом.

        :param text: Искомый текст.
        :param first_only: Флаг для поиска только первого вхождения.
        :return:
        """
        text = text.strip()
        # TODO: класс не должен знать про покраску, изменить вывод
        #  на наборы run`ов, придумать, где привести вывод в плоский вид
        runs_to_color = []
        for runs_list in self.__get_runs_sequences(text, first_only=first_only):
            for run in runs_list:
                runs_to_color.append(run)
        return runs_to_color

    def __get_runs_sequences(self,
                             text: str,
                             first_only: bool = False,
                             ) -> List[List[Run]]:
        """Извлекает наборы объектов Run, которые в совокупности содержат text.

        :param text: Искомый текст.
        :param first_only:
            True - возвращается список с первым соответствующим Run.
            False - возвращается список со всеми соответствующими Run.
        :return: Список наборов объектов Run, содержащих text.
        """
        runs = []
        possible_runs = list(
            self.__search_text_parts(self.run_collector, text))
        for run_index, possible_run in enumerate(possible_runs):
            run, text_part = possible_run
            if self.__check_text_in_element(run, text, strict=True):
                runs.append([run])
                if first_only:
                    return runs
                continue
            temp_runs = []
            temp_text = []
            for temp_possible_run in possible_runs[run_index:]:
                temp_run, temp_text_part = temp_possible_run
                # FIXME по какой-то причине в аллокацию попадает run
                #  с несоответствующим текстом,
                #  ошибка в __search_text_parts, вероятно,
                #  нужно добиться повторяемости,
                #  т.к. пока что выглядит как случайность
                print('*'*200)
                print('='*100)
                print(temp_run.text)
                print(temp_text_part)
                print('='*100)
                if self.__check_text_in_element(
                        temp_run, temp_text_part, strict=True):
                    temp_runs.append(temp_run)
                    temp_text.append(temp_text_part)
                else:
                    try:
                        # TODO не аллоцировать пока не соберется полный текст.
                        # FIXME Есть вероятность, что в некоторых случаях, при нахождении слова,
                        #  поиск продолжается, и выражение ниже уходит в ошибку,
                        #  вследствие чего необходимый Run выделен, но не покрашен.
                        #  Нужно проверять на соответствие тексту в конце итерации или сразу после начала.
                        #  Альтернативы?
                        print(temp_run.text)
                        print(temp_text)
                        temp_runs.append(self.__allocate_run_with_text(
                            temp_run, temp_text_part))
                        temp_text.append(temp_text_part)
                    except ValueError:
                        print([r.text for r in temp_runs])
                        print(temp_text)
                        temp_runs = []
                        temp_text = []
                        continue
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
    def __search_text_parts(runs: List[Run],
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
    def __check_text_in_element(element: Run | Paragraph,
                                text: str,
                                strict: bool = False
                                ) -> bool:
        """Проверяет объект на содержание text.

        :param element: Проверяемый объект.
        :param text: Искомый текст.
        :param strict:
            True - проверка объекта на равенство без учёта пробельных символов.
            False - проверка на наличие text в объекте.
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
