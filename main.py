from copy import deepcopy
from typing import List, Dict, Generator, Tuple

from docx import Document
from docx.text.paragraph import Paragraph, Run

from utility import Index, Color


#  TODO: Ящик с инструментами
#   покраска таблиц в .docx.
#   попробовать взять другую либу для цветов
class DocxPainter:

    @classmethod
    def color_text(cls,
                   document: Document,
                   text: str,
                   color: str = 'red',
                   first_only: bool = False
                   ) -> None:
        text = text.strip()
        for paragraph in document.paragraphs:
            if not cls.__find_text(paragraph, text, strict=False):
                continue
            for run in cls.__get_runs_to_color(
                    paragraph, text, first_only=first_only):
                cls.__color_run(run, color)

    @staticmethod
    def __find_text(element: Run | Paragraph,
                    text: str,
                    strict: bool = True
                    ) -> bool:
        if strict:
            return text == element.text.strip()
        return text in element.text

    @classmethod
    def __get_runs_to_color(cls,
                            paragraph: Paragraph,
                            text: str,
                            start: int = 0,
                            first_only: bool = False
                            ) -> List[Run]:
        runs_to_color = []
        for run, text in cls.__find_text_in_runs(
                paragraph.runs[start:], text):
            if cls.__find_text(run, text, strict=True):
                runs_to_color.append(run)
                if first_only:
                    return runs_to_color
                continue
            if cls.__find_text(run, text, strict=False):
                start = [r.text for r in paragraph.runs].index(run.text)
                runs_to_color.append(cls.__reshape_run_with_text(
                    paragraph, run, text))
                if first_only:
                    return runs_to_color
                runs_to_color += cls.__get_runs_to_color(
                    paragraph, text, start)
                break
        return runs_to_color

    @classmethod
    def __find_text_in_runs(cls,
                            runs: List[Run],
                            text: str
                            ) -> Generator[Tuple[Run, str], None, None]:
        # TODO Отрефакторить
        text_symbols = list(text)
        runs_combination: Dict[Run, str] = {}
        for run in runs:
            run_contains: List[str] = []
            for run_symbol in run.text:
                try:
                    symbol = text_symbols.pop(Index.first)
                    if run_symbol != symbol:
                        runs_combination.clear()
                        run_contains.clear()
                        text_symbols = cls.__text_symbols_renew(text)
                        continue
                    run_contains.append(symbol)
                except IndexError:
                    value = ''.join(run_contains)
                    runs_combination.update({run: value})
                    if value:
                        yield run, runs_combination[run]
                    runs_combination.clear()
                    run_contains.clear()
                    text_symbols = cls.__text_symbols_renew(text)
                    continue
            if run_contains:
                value = ''.join(run_contains)
                runs_combination.update({run: value})
                if value:
                    yield run, runs_combination[run]

    @staticmethod
    def __text_symbols_renew(text: str) -> list[str]:
        return list(text)

    @classmethod
    def __reshape_run_with_text(cls,
                                paragraph: Paragraph,
                                run: Run,
                                text: str
                                ) -> Run:
        # TODO попробовать выделить отсюда часть по сборке параграфа
        run_with_text_after_split_index = 1
        runs = paragraph.runs
        run_index = [r.text for r in runs].index(run.text)
        new_runs = cls.__split_run(run, text)
        paragraph.clear()
        cls.__add_runs(
            paragraph, runs[:run_index] + new_runs + runs[run_index + 1:])
        return new_runs[run_with_text_after_split_index]

    @staticmethod
    def __split_run(run: Run, text: str) -> list[Run]:
        run_text_parts = run.text.split(text, maxsplit=1)
        first_r = deepcopy(run)
        second_r = deepcopy(run)
        third_r = deepcopy(run)
        first_r.text = run_text_parts[Index.first]
        second_r.text = text
        third_r.text = run_text_parts[Index.last]
        return [first_r, second_r, third_r]

    @staticmethod
    def __add_runs(paragraph: Paragraph, runs: list[Run]) -> None:
        runs_number = len(paragraph.runs)
        paragraph.append_runs(runs)
        # append_runs ставит Run(' ') в начало, убираем
        paragraph.runs[runs_number].clear()

    @staticmethod
    def __color_run(run: Run, color: str) -> None:
        run.font.color.rgb = Color.get(color)


if __name__ == '__main__':
    expected = 'СЛОВО'
    doc = Document('template.docx')
    DocxPainter.color_text(doc, expected, 'purple')
    # неявное сохранение, пересмотреть базу класса
    doc.save('new.docx')
