from docx import Document

import docx_styler

if __name__ == '__main__':
    expected = 'СЛОВО'
    doc = Document('test/template.docx')

    docx_styler.color_text(doc, expected, 'red')
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

    doc.save('new.docx')
