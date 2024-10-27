import os

from docx import Document

import docx_styler

if __name__ == '__main__':
    filename = 'template.docx'
    input_dir = 'test'
    filepath = os.path.join(input_dir, filename)
    doc = Document(filepath)

    expected = 'СЛОВО'
    docx_styler.color_text(doc, expected, color='green')
    docx_styler.highlight_text(doc, expected, color='darkred')

    docx_styler.color_text(doc, 'параграф', color=(0, 0, 255))
    docx_styler.color_text(doc, 'текстом', color=(0, 0, 300))
    docx_styler.color_text(doc, 'форматированием', color=(0, 0, 255, 14))
    docx_styler.color_text(doc, 'перед', color=(0, 0))
    docx_styler.highlight_text(doc, 'Просто', color='dark')
    # p = docx_styler.get_paragraphs_with_text(
    #     doc, expected, first_only=True)[0]
    # r = docx_styler.get_runs_with_text(
    #     p, expected, first_only=True)[0]
    # # Комментарий к run
    # r.add_comment('Комментарий', author='Полное имя', initials='Инициалы')
    # # run.add_comment ломает документ по какой-то причине

    # # Комментарий к paragraph
    # p.add_comment('Комментарий', author='Полное имя', initials='Инициалы')
    # # p.add_comment работает нормально

    # # Примечание к paragraph
    # p.add_footnote('Примечание')
    # # add_footnote добавляет порядковый номер примечания в конец параграфа в
    # # обычном регистре, выглядит дерьмово
    # # можно брать последний Run в параграфе и менять его стиль.
    # # Создать под это функциональность в core

    doc.save(f'{filename[:-5]}_result.docx')
