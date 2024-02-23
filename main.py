from copy import deepcopy

from docx.text.paragraph import Paragraph, Run
from docx import Document

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

    def color_list_of_phrases(self, phrases: list[str],
                              first_only: bool = False, color: str = 'red'):
        for phrase in phrases:
            self.color_phrase(phrase, first_only, color)

    def color_phrase(self, phrase: str,
                     first_only: bool = False, color: str = 'red'):
        phrase = phrase.strip()
        for p in self.paragraphs:
            if not self.__find_phrase(p, phrase, strict=False):
                continue
            start = 0
            runs_to_color = self.__find_phrase_in_runs(p.runs[start:], phrase)
            for r, phrase in runs_to_color:
                if self.__find_phrase(r, phrase, strict=True):
                    self.__color_r(r, color)
                    if first_only:
                        return
                    continue
                if self.__find_phrase(r, phrase, strict=False):
                    start = [r.text for r in p.runs].index(r.text)
                    run = self.__reshape_r_with_phrase(p, r, phrase)
                    self.__color_r(run, color)
                    if first_only:
                        return
                    continue

    @staticmethod
    def __find_phrase(el: Run or Paragraph, phrase: str,
                      strict: bool = True) -> bool:
        if strict:
            if phrase == el.text.strip():
                return True
        else:
            if phrase in el.text:
                return True
        return False

    def __find_phrase_in_runs(self, runs: list[Run],
                              phrase: str) -> dict[str: Run]:
        symbols = list(phrase)
        runs_combination = {}
        for r in runs:
            r_symbols = list(r.text)
            r_contains = []
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
    def __phrase_symbols_renew(phrase):
        return list(phrase)

    def __color_r(self, r: Run, color: str):
        r.font.color.rgb = self.clr[color]

    def __reshape_r_with_phrase(self, p: Paragraph, r: Run, phrase: str):
        # TODO попробовать выделить отсюда часть по сборке параграфа
        r_with_phrase_after_split_index = 1
        r_index = [r.text for r in p.runs].index(r.text)
        runs_before_phrase = p.runs[:r_index]
        new_runs = self.__split_r(r, phrase)
        runs_after_phrase = p.runs[r_index+1:]
        r_with_phrase = new_runs[r_with_phrase_after_split_index]
        runs = runs_before_phrase + new_runs + runs_after_phrase
        p.clear()
        self.__add_runs(p, runs)
        return r_with_phrase

    @staticmethod
    def __split_r(r: Run, phrase: str):
        text_parts = r.text.split(phrase, maxsplit=1)
        first_r = deepcopy(r)
        first_r.text = text_parts[Index.first]
        second_r = deepcopy(r)
        second_r.text = phrase
        third_r = deepcopy(r)
        third_r.text = text_parts[Index.last]
        return [first_r, second_r, third_r]

    @staticmethod
    def __add_runs(p: Paragraph, runs: list[Run]):
        runs_number = len(p.runs)
        p.append_runs(runs)
        # append_runs ставит Run(' ') в начало, убираем
        p.runs[runs_number].clear()


if __name__ == '__main__':
    expected = 'СЛОВО'
    d = Document('template.docx')
    painter = DocxPainter(d)
    painter.color_phrase(expected)
    d.save('new.docx')
