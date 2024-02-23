from copy import deepcopy
from typing import List, Dict, Generator

from docx import Document
from docx.text.paragraph import Paragraph, Run

from utility import Index, Color


#  TODO: покраска таблиц в .docx
class DocxPainter:

    def __init__(self, document: Document):
        self.__d = document
        self.clr = Color()
        self.__paragraphs = [
            el for el in self.__d.elements if isinstance(el, Paragraph)]

    @property
    def paragraphs(self) -> list[Paragraph]:
        return self.__paragraphs

    def color_phrases_list(self,
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
        for p in self.paragraphs:
            if not self.__find_phrase(p, phrase, strict=False):
                continue
            start = 0
            for r in self.__get_runs_to_color(
                    p, phrase, start, color, first_only):
                self.__color_r(r, color)

    @staticmethod
    def __find_phrase(el: Run | Paragraph,
                      phrase: str,
                      strict: bool = True
                      ) -> bool:
        if strict:
            if phrase == el.text.strip():
                return True
        else:
            if phrase in el.text:
                return True
        return False

    def __get_runs_to_color(self,
                            p: Paragraph,
                            phrase: str,
                            start: int,
                            color: str,
                            first_only: bool = False
                            ) -> list[Run]:
        runs_with_phrase = self.__find_phrase_in_runs(p.runs[start:], phrase)
        runs_to_color = []
        for r, phrase in runs_with_phrase:
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
                runs_after_reshape = self.__get_runs_to_color(
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
            r_symbols = list(r.text)
            r_contains: List[str] = []
            for r_symbol in r_symbols:
                try:
                    symbol = symbols.pop(Index.first)
                    if r_symbol != symbol:
                        runs_combination.clear()
                        r_contains.clear()
                        symbols = self.__phrase_symbols_renew(phrase)
                        continue
                    r_contains.append(symbol)
                except IndexError:
                    value = ''.join(r_contains)
                    runs_combination.update({r: value})
                    if value:
                        yield r, runs_combination[r]
                    runs_combination.clear()
                    r_contains.clear()
                    symbols = self.__phrase_symbols_renew(phrase)
                    continue
            if r_contains:
                value = ''.join(r_contains)
                runs_combination.update({r: value})
                if value:
                    yield r, runs_combination[r]

    @staticmethod
    def __phrase_symbols_renew(phrase: str) -> list[str]:
        return list(phrase)

    def __reshape_r_with_phrase(self,
                                p: Paragraph,
                                r: Run,
                                phrase: str
                                ) -> Run:
        # TODO попробовать выделить отсюда часть по сборке параграфа
        r_with_phrase_after_split_index = 1
        r_index = [r.text for r in p.runs].index(r.text)
        runs_before_phrase = p.runs[:r_index]
        new_runs = self.__split_r(r, phrase)
        runs_after_phrase = p.runs[r_index + 1:]
        r_with_phrase = new_runs[r_with_phrase_after_split_index]
        runs = runs_before_phrase + new_runs + runs_after_phrase
        p.clear()
        self.__add_runs(p, runs)
        return r_with_phrase

    @staticmethod
    def __split_r(r: Run, phrase: str) -> list[Run]:
        text_parts = r.text.split(phrase, maxsplit=1)
        first_r = deepcopy(r)
        second_r = deepcopy(r)
        third_r = deepcopy(r)
        first_r.text = text_parts[Index.first]
        second_r.text = phrase
        third_r.text = text_parts[Index.last]
        return [first_r, second_r, third_r]

    @staticmethod
    def __add_runs(p: Paragraph, runs: list[Run]) -> None:
        runs_number = len(p.runs)
        p.append_runs(runs)
        # append_runs ставит Run(' ') в начало, убираем
        p.runs[runs_number].clear()

    def __color_r(self, r: Run, color: str) -> None:
        r.font.color.rgb = self.clr[color]


if __name__ == '__main__':
    expected = 'СЛОВО'
    d = Document('template.docx')
    painter = DocxPainter(d)
    painter.color_phrase(expected, 'magenta')
    d.save('new.docx')
