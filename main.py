from copy import deepcopy
from typing import List, Dict, Generator

from docx import Document
from docx.text.paragraph import Paragraph, Run

from utility import Index, Color


#  TODO: покраска таблиц в .docx
class DocxPainter:

    def __init__(self, document: Document):
        self.__d = document  # Сокращенное имя атрибута.
        self.clr = Color()  # Сокращенное имя атрибута.
        self.__paragraphs = [
            el for el in self.__d.elements if isinstance(el, Paragraph)]  # Сокращенное имя переменной el. Что мешает просто взять .paragraphs?

    @property
    def paragraphs(self) -> list[Paragraph]:
        return self.__paragraphs

    def color_phrases_list(self,  # Можно убрать "list" из имени. Почему не classmethod?
                           phrases: list[str],
                           color: str = 'red',
                           first_only: bool = False
                           ) -> None:
        for phrase in phrases:
            self.color_phrase(phrase, color, first_only)

    def color_phrase(self,
                     phrase: str,
                     color: str = 'red',
                     first_only: bool = False
                     ) -> None:
        phrase = phrase.strip()
        for p in self.paragraphs:  # Сокращенное имя переменной.
            if not self.__find_phrase(p, phrase, strict=False):
                continue
            start = 0  # Одноразовая переменная.  Лучше вынести в параметр по умолчанию __get_runs_to_color.
            for r in self.__get_runs_to_color(  # Сокращенное имя переменной.
                    p, phrase, start, color, first_only):
                self.__color_r(r, color)

    @staticmethod
    def __find_phrase(el: Run | Paragraph,  # Сокращенное имя параметра.
                      phrase: str,
                      strict: bool = True
                      ) -> bool:
        if strict:  # return phrase in el.text.strip() if strict else el.text
            if phrase == el.text.strip():
                return True
        else:
            if phrase in el.text:
                return True
        return False

    def __get_runs_to_color(self,
                            p: Paragraph,  # Сокращенное имя параметра.
                            phrase: str,
                            start: int,
                            color: str,  # А оно здесь зачем?
                            first_only: bool = False
                            ) -> list[Run]:
        runs_to_check = p.runs[start:]  # Одноразовая переменная.
        runs_with_phrase = self.__find_phrase_in_runs(runs_to_check, phrase)  # Одноразовая переменная.
        runs_to_color = []
        for r, phrase in runs_with_phrase:  # Сокращенное имя переменной.
            if self.__find_phrase(r, phrase, strict=True):
                runs_to_color.append(r)
                if first_only:
                    return [r, ]
                continue
            if self.__find_phrase(r, phrase, strict=False):
                start = [r.text for r in p.runs].index(r.text)
                run = self.__reshape_r_with_phrase(p, r, phrase)
                if first_only:
                    return [run, ]
                runs_to_color.append(run)
                runs_after_reshape = self.__get_runs_to_color(  # Одноразовая переменная.
                    p, phrase, start, color)
                runs_to_color += runs_after_reshape
                break
        return runs_to_color

    def __find_phrase_in_runs(self,
                              runs: List[Run],
                              phrase: str
                              ) -> Generator[tuple[Run, str], None, None]:
        symbols = list(phrase)
        runs_combination: Dict[Run, str] = {}
        for r in runs:
            r_symbols = list(r.text)  # Одноразовая переменная. По тексту тоже можно итерироваться.
            r_contains: List[str] = []
            for r_symbol in r_symbols:
                try:
                    symbol = symbols.pop(Index.first)  # Каждую итерацию в списке перестраиваются индексы. Может лучше использовать указатели?
                    if r_symbol != symbol:
                        runs_combination.clear()
                        r_contains.clear()
                        symbols = self.__phrase_symbols_renew(phrase)
                        continue  # Можно заменить на else - читабельность. Можно инвертировать условие в if - сдвиг табуляции.
                    r_contains.append(symbol)
                except IndexError:
                    value = ''.join(r_contains)
                    runs_combination.update({r: value})
                    if value:
                        yield r, runs_combination[r]
                    runs_combination.clear()
                    r_contains.clear()
                    symbols = self.__phrase_symbols_renew(phrase)
                    continue  # Зачем?
            if r_contains:
                value = ''.join(r_contains)  # Код повторяется выше в except.
                runs_combination.update({r: value})
                if value:
                    yield r, runs_combination[r]

    @staticmethod
    def __phrase_symbols_renew(phrase: str) -> list[str]:
        return list(phrase)

    def __reshape_r_with_phrase(self,  # Сокращенное имя метода.
                                p: Paragraph,  # Сокращенное имя параметра.
                                r: Run,  # Сокращенное имя параметра.
                                phrase: str
                                ) -> Run:
        # TODO попробовать выделить отсюда часть по сборке параграфа
        r_with_phrase_after_split_index = 1  # Одноразовая переменная.
        r_index = [r.text for r in p.runs].index(r.text)
        runs_before_phrase = p.runs[:r_index]  # Одноразовая переменная.
        new_runs = self.__split_r(r, phrase)
        runs_after_phrase = p.runs[r_index + 1:]  # Одноразовая переменная.
        r_with_phrase = new_runs[r_with_phrase_after_split_index]
        runs = runs_before_phrase + new_runs + runs_after_phrase  # Одноразовая переменная.
        p.clear()
        self.__add_runs(p, runs)
        return r_with_phrase

    @staticmethod
    def __split_r(r: Run, phrase: str) -> list[Run]:  # Сокращенное имя метода.
        text_parts = r.text.split(phrase, maxsplit=1)
        first_r = deepcopy(r)  # first_r = second_r = third_r = deepcopy(r)
        second_r = deepcopy(r)
        third_r = deepcopy(r)
        first_r.text = text_parts[Index.first]
        second_r.text = phrase
        third_r.text = text_parts[Index.last]
        return [first_r, second_r, third_r]

    @staticmethod
    def __add_runs(p: Paragraph, runs: list[Run]) -> None:  # Сокращенное имя параметра.
        runs_number = len(p.runs)
        p.append_runs(runs)
        # append_runs ставит Run(' ') в начало, убираем
        p.runs[runs_number].clear()

    def __color_r(self, r: Run, color: str) -> None:  # Сокращенное имя метода. Сокращенное имя параметра.
        r.font.color.rgb = self.clr[color]


if __name__ == '__main__':
    expected = 'СЛОВО'  # Одноразовая переменная. Лучше в color_phrase указать имя параметра.
    d = Document('template.docx')  # Сокращенное имя переменной.
    painter = DocxPainter(d)  # Одноразовая переменная.
    painter.color_phrase(expected, 'magenta')
    d.save('new.docx')  # Document изменяется неявно???
