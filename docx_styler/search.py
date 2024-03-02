from typing import Generator, List, Tuple

from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .change import allocate_run_with_text
from .static import Index


def get_runs_with_text(paragraph: Paragraph,
                       text: str,
                       first_only: bool = False,
                       start: int = 0
                       ) -> List[Run]:
    runs = []
    for run, text_part in __find_text_in_runs(paragraph.runs[start:], text):
        if check_text_in_element(run, text_part, strict=True):
            runs.append(run)
            if first_only:
                return runs
            continue
        if check_text_in_element(run, text_part, strict=False):
            start = [run.text for run in paragraph.runs].index(run.text)
            runs.append(allocate_run_with_text(
                paragraph, run, text_part))
            if first_only:
                return runs
            runs += get_runs_with_text(paragraph, text, start=start)
            break
    return runs


def check_text_in_element(element: Run | Paragraph,
                          text: str,
                          strict: bool = False
                          ) -> bool:
    if strict:
        return text == element.text.strip()
    return text in element.text


def __find_text_in_runs(runs: List[Run],
                        text: str
                        ) -> Generator[Tuple[Run, str], None, None]:
    # FIXME красит лишнее если run заканчивается, пара букв в него попала,
    #  но в следующем run нет продолжения. Безумно редкий случай,
    #  скорее всего, можно создать только искусственно (см. template.docx)
    text_symbols = list(text)
    for run in runs:
        run_contains: List[str] = []
        for run_symbol in run.text:
            try:
                symbol = text_symbols.pop(Index.first)
                if run_symbol != symbol:
                    run_contains.clear()
                    text_symbols = __text_symbols_renew(text)
                else:
                    run_contains.append(symbol)
            except IndexError:
                if run_contains:
                    yield run, ''.join(run_contains)
                run_contains.clear()
                text_symbols = __text_symbols_renew(text)
                continue
        if run_contains:
            yield run, ''.join(run_contains)


def __text_symbols_renew(text: str) -> list[str]:
    return list(text)
