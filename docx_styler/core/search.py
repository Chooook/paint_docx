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

import re
from copy import deepcopy

from docx import Document
from docx.text.paragraph import Paragraph
from typing_extensions import Match, Pattern

from docx_styler.core.runs_map import RunWithSpan, RunsMap


class RunFinder:
    def __init__(self, document: Document):
        self.runs_map = RunsMap(document)

    def search_and_allocate(self, text_pattern: str):
        result_runs = []
        pattern = re.compile(text_pattern)
        matches = list(re.finditer(pattern, self.runs_map.text))

        for match in matches:
            matching_runs = self.runs_map.find_runs_by_span(
                match.start(), match.end())

            for run_info in matching_runs:
                if run_info.run.text == text_pattern:
                    result_runs.append(run_info.run)
                    continue

                matching_text = self.__define_metching_text(match, run_info)
                allocated_runs = self.__allocate_runs(
                    matching_text, run_info, pattern)
                result_runs += allocated_runs

        return result_runs

    @staticmethod
    def __define_metching_text(match: Match, run_info: RunWithSpan):
        match_text = match.group()
        run_text = run_info.run.text
        m_start, m_end = match.start(), match.end()
        r_start, r_end = run_info.start, run_info.end

        if r_start < m_start:
            if r_end >= m_end:
                matching_text = match_text
            else:  # r_end < m_end
                matching_text = run_text[m_start - r_end:]

        elif r_start == m_start:
            if r_end >= m_end:
                matching_text = match_text
            else:  # r_end < m_end
                matching_text = run_text

        else:  # r_start > m_start
            if r_end <= m_end:
                matching_text = run_text
            else:  # r_end > m_end
                matching_text = run_text[r_start - m_end:]

        return matching_text

    def __allocate_runs(self,
                        matching_text: str,
                        run_info: RunWithSpan,
                        base_pattern: Pattern):
        run = run_info.run
        paragraph: Paragraph = run._parent
        runs_elements = [r.element for r in paragraph.runs]
        run_index = runs_elements.index(run.element)

        runs_before = paragraph.runs[:run_index]
        runs_after = paragraph.runs[run_index + 1:]

        result_runs = []
        middle_runs = []

        if re.search(base_pattern, matching_text):
            matching_run, new_runs = self.__split_run(matching_text, run)
            result_runs.append(matching_run)
            middle_runs += new_runs

            while matching_run.element != new_runs[-1].element:
                matching_run, new_runs = self.__split_run(
                    matching_text, new_runs[-1])
                result_runs.append(matching_run)
                middle_runs += new_runs
        else:
            matching_run, new_runs = self.__split_run(matching_text, run)
            result_runs.append(matching_run)
            middle_runs += new_runs

        paragraph.clear()
        paragraph_runs = runs_before + middle_runs + runs_after
        for run in paragraph_runs:
            paragraph._element.append(run.element)

        return result_runs

    @staticmethod
    def __split_run(split_by, run):
        before_text, middle_text, after_text = run.text.partition(split_by)
        result_runs = []

        if before_text:
            before_run = deepcopy(run)
            before_run.text = before_text
            result_runs.append(before_run)

        run.text = middle_text
        result_runs.append(run)

        if after_text:
            after_run = deepcopy(run)
            after_run.text = after_text
            result_runs.append(after_run)

        return run, result_runs
