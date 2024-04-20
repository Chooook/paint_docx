from docx import Document
# from docx.enum.text import WD_COLOR_INDEX

import docx_styler

if __name__ == '__main__':
    expected = 'СЛОВО'
    doc = Document('test/template.docx')

    docx_styler.color_text(doc, expected, color='red')
    docx_styler.fill_text_with_color(doc, expected, color='darkgreen')
    # p = docx_styler.get_paragraphs_with_text(
    #     doc, expected, first_only=True)[0]
    # r = docx_styler.get_runs_with_text(
    #     p, expected, first_only=True)[0]
    # r.font.highlight_color = 1
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
    # # TODO можно брать последний Run в параграфе и менять его стиль.
    # #  Создать под это функциональность в core

    doc.save('new.docx')
