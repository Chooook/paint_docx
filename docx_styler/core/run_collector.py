from typing import Iterator

from docx import Document
from docx.text.run import Run


class RunCollector:
    def __init__(self, document: Document) -> None:
        self.document = document
        self.text_runs = self.update_text_runs()
        self.tables_runs_by_rows = self.__update_tables_runs_by_rows()
        self.tables_runs_by_columns = self.__update_tables_runs_by_columns()
        self.footnotes_runs = self.update_footnotes_runs()

        self.document_runs = self.update_document_runs(reload_all=False)

    def __iter__(self) -> Iterator[Run]:
        return iter(self.document_runs)

    def update_text_runs(self) -> list[Run]:
        runs = []
        for paragraph in self.document.paragraphs:
            runs.extend(paragraph.runs)
        self.text_runs = runs
        if self.document_runs:
            self.update_document_runs(reload_all=False)
        return self.text_runs

    def update_tables_runs(self) -> tuple[list[Run], list[Run]]:
        self.tables_runs_by_rows = self.__update_tables_runs_by_rows()
        self.tables_runs_by_columns = self.__update_tables_runs_by_columns()
        if self.document_runs:
            self.update_document_runs(reload_all=False)
        return self.tables_runs_by_rows, self.tables_runs_by_columns

    def __update_tables_runs_by_rows(self) -> list[Run]:
        runs = []
        for table in self.document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        runs.extend(paragraph.runs)
        return runs

    def __update_tables_runs_by_columns(self) -> list[Run]:
        # TODO: проверить, что будет с объединёнными ячейками
        runs = []
        for table in self.document.tables:
            num_rows = len(table.rows)
            num_cols = len(table.columns)
            for col in range(num_cols):
                for row in range(num_rows):
                    cell = table.cell(row, col)
                    for paragraph in cell.paragraphs:
                        runs.extend(paragraph.runs)
        return runs

    def update_footnotes_runs(self) -> list[Run]:
        runs = []
        for footnote in self.document.footnotes:
            for paragraph in footnote.paragraphs:
                runs.extend(paragraph.runs)
        self.footnotes_runs = runs
        if self.document_runs:
            self.update_document_runs(reload_all=False)
        return self.footnotes_runs

    def update_document_runs(self, reload_all: bool = True) -> list[Run]:
        if reload_all:
            self.update_text_runs()
            self.update_tables_runs()
            self.update_footnotes_runs()
        # Для разделения объектов Run в различных сущностях документа
        run_with_space = self.__create_run_with_space()
        self.document_runs = (self.text_runs
                              + [run_with_space]
                              + self.tables_runs_by_rows
                              + [run_with_space]
                              + self.tables_runs_by_columns
                              + [run_with_space]
                              + self.footnotes_runs)
        return self.document_runs

    @staticmethod
    def __create_run_with_space() -> Run:
        # Для создания Run вне текущего документа, необходим объект Document
        return Document().add_paragraph().add_run(' ')
