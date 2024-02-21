from copy import deepcopy

from docx.text.paragraph import Paragraph, Run
from docx.shared import RGBColor
from docx import Document


class ColorPalette:
    red = RGBColor(255, 0, 0)
    orange = RGBColor(255, 255, 0)
    yellow = RGBColor(0, 255, 255)
    green = RGBColor(0, 255, 0)
    blue = RGBColor(0, 0, 255)
    purple = RGBColor(255, 0, 255)


class DocxPainter:

    DEBUG = False
    first_el = 0
    last_el = -1

    def __init__(self, document: Document):
        self.__d = document
        self.clr = ColorPalette
        self.__paragraphs = [
            el for el in self.__d.elements if isinstance(el, Paragraph)]

    @property
    def paragraphs(self) -> list[Paragraph]:
        return self.__paragraphs

    def color_list_of_phrases(
            self, phrases: list[str], first_only: bool = False):
        for phrase in phrases:
            self.color_phrase(phrase, first_only)

    def color_phrase(self, phrase: str, first_only: bool = False):
        phrase = phrase.strip()
        for p in self.paragraphs:
            if not self.__find_phrase(p, phrase, strict=False):
                continue
            for r in p.runs:
                if self.__find_phrase(r, phrase, strict=True):
                    self.__color_r(r)
                    if first_only:
                        return
                    continue
                if self.__find_phrase(r, phrase, strict=False):
                    run = self.__reshape_r_with_phrase(p, r, phrase)
                    self.__color_r(run)
                    if first_only:
                        return
                    continue
                # runs = self.__find_runs_with_phrase(p, phrase)
            # FIXME собьётся итерация если сделать reshape? мб нет потому
            #  что количество увеличивается
            # TODO иначе искать группу run`ов

    @staticmethod
    def __find_phrase(
            el: Run or Paragraph, phrase: str, strict: bool = True) -> bool:
        if strict:
            if phrase == el.text.strip():
                return True
        else:
            if phrase in el.text:
                return True
        return False

    def __color_r(self, r: Run):
        r.font.color.rgb = self.clr.red

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

    def __split_r(self, r: Run, phrase: str):
        text_parts = r.text.split(phrase, maxsplit=1)
        first_r = deepcopy(r)
        first_r.text = text_parts[self.first_el]
        second_r = deepcopy(r)
        second_r.text = phrase
        third_r = deepcopy(r)
        third_r.text = text_parts[self.last_el]
        return [first_r, second_r, third_r]

    @staticmethod
    def __add_runs(p: Paragraph, runs: list[Run]):
        runs_number = len(p.runs)
        p.append_runs(runs)
        # append_runs ставит Run(' ') в начало, убираем
        p.runs[runs_number].clear()

    def __find_runs_with_phrase(self, p: Paragraph, phrase: str):
        # собирать посимвольно всю фразу поочереди прогоняя раны
        ...


if __name__ == '__main__':
    # TODO варианты расположения `string` в документе:
    #  * `string` целиком в `run`, других слов нет           +
    #  * `string` целиком в `run`, есть слово до             +
    #  * `string` целиком в `run`, есть слово после          +
    #  * `string` целиком в `run`, есть слова до и после     +
    #  * `string` разбит на несколько `run`                  -
    #  * покраска таблиц в .docx                             -
    #  * покраска всех вхождений одного слова                +
    #  ------------------------------------------------------------------------

    expected = 'СЛОВО'
    # expected = 'СЛОВИЩЕ'

    d = Document('template.docx')
    painter = DocxPainter(d)
    painter.color_phrase(expected)
    d.save('new.docx')
