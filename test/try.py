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
    docx_styler.highlight_text(doc, expected, color='magenta')
    docx_styler.color_text(doc, 'ФРАЗА КОТОРУЮ КРАСИМ', color='blue')
    docx_styler.highlight_text(doc, 'ФРАЗА КОТОРУЮ КРАСИМ', color='yellow')

    docx_styler.color_text(doc, 'параграф', color=(0, 0, 255))
    # Примеры для проверки warnings
    docx_styler.color_text(doc, 'текстом', color=(0, 0, 300))
    docx_styler.color_text(doc, 'форматированием', color=(0, 0, 255, 14))
    docx_styler.color_text(doc, 'перед', color=(0, 0))
    docx_styler.highlight_text(doc, 'Просто', color='dark')

    doc.save(f'{filename[:-5]}_result.docx')
