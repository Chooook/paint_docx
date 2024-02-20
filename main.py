from collections import namedtuple
from copy import deepcopy

from docx.text.paragraph import Paragraph, Run
from docx.shared import RGBColor
from docx import Document

# TODO искать подмножество run`ов, которые соответствуют фразе, дальше
#  думать что делать и как их разбивать.
MyRun = namedtuple('MyRun', ['text', 'style'])
NewOrigRun = namedtuple('NewOrigRun', ['new', 'original'])


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
        self.__paragraphs = set(
            [el for el in self.__d.elements if isinstance(el, Paragraph)])

    @property
    def paragraphs(self) -> set[Paragraph]:
        return self.__paragraphs

    def color_list_of_phrases(self, phrases: list[str]):
        for phrase in phrases:
            self.color_phrase(phrase)

    def color_phrase(self, phrase: str):
        phrase = phrase.strip()
        p = self.__get_paragraph_with_phrase(phrase)
        try:
            run = self.__find_run(p, phrase, strict=True)
            self.__color_run(run)
            return
        except ValueError:
            pass
        try:
            run = self.__find_run(p, phrase, strict=False)
            run = self.__reshape_phrase_run(p, run, phrase)
            self.__color_run(run)
        except ValueError:
            pass
        # иначе искать группу run`ов

    def __get_paragraph_with_phrase(self, phrase: str) -> Paragraph:
        for p in self.__paragraphs:
            if phrase in p.text:
                return p
        raise ValueError('Фраза не найдена в тексте.')

    @staticmethod
    def __find_run(p: Paragraph, phrase: str, strict: bool = True):
        if strict:
            for r in p.runs:
                if phrase == r.text.strip():
                    return r
            raise ValueError('Не найдено объекта run, соответствующего фразе.')
        else:
            for r in p.runs:
                if phrase in r.text:
                    return r
            raise ValueError('Не найдено объекта run, содержащего фразу.')

    def __color_run(self, run: Run):
        run.font.color.rgb = self.clr.red

    def __reshape_phrase_run(self, p: Paragraph, run: Run, phrase: str):
        # TODO попробовать выделить отсюда часть по сборке параграфа
        run_with_phrase_after_split_index = 1
        run_index = [r.text for r in p.runs].index(run.text)
        runs_before_phrase = p.runs[:run_index]
        new_runs = self.__split_run(run, phrase)
        runs_after_phrase = p.runs[run_index+1:]
        run_with_phrase = new_runs[run_with_phrase_after_split_index]
        runs = runs_before_phrase + new_runs + runs_after_phrase
        p.clear()
        self.__add_runs(p, runs)
        return run_with_phrase

    def __add_runs(self, p: Paragraph, runs: list[Run]):
        p.append_runs(runs)
        # append_runs ставит Run(' ') в начало, убираем
        self.__clear_first_run(p)

    def __split_run(self, run: Run, phrase: str):
        text_parts = run.text.split(phrase, maxsplit=1)
        first_run = deepcopy(run)
        first_run.text = text_parts[self.first_el]
        second_run = deepcopy(run)
        second_run.text = phrase
        third_run = deepcopy(run)
        third_run.text = text_parts[self.last_el]
        return [first_run, second_run, third_run]

    @staticmethod
    def __clear_first_run(p: Paragraph):
        p.runs[0].clear()


if __name__ == '__main__':
    # TODO варианты расположения `string` в документе:
    #  * `string` целиком в `run`, других слов нет           +
    #  * `string` целиком в `run`, есть слово до             +
    #  * `string` целиком в `run`, есть слово после          +
    #  * `string` целиком в `run`, есть слова до и после     +
    #  * `string` разбит на несколько `run`                  -
    #  ------------------------------------------------------------------------

    expected = 'ФРАЗА ЦЕЛИКОМ'
    # expected = 'ФРАЗА СО СЛОВАМИ ДО'
    # expected = 'ФРАЗА СО СЛОВАМИ ПОСЛЕ'
    # expected = 'ФРАЗА ДО И ПОСЛЕ'
    # expected = 'СЛОВИЩЕ'

    def debug(ptr):
        from pprint import pprint
        for par in ptr.paragraphs:
            # if string not in p.text:
            #     continue
            pprint([r.text for r in par.runs])

    d = Document('template.docx')
    painter = DocxPainter(d)
    # debug(painter)
    painter.color_phrase(expected)
    # debug(painter)
    d.save('new.docx')
