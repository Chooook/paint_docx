from docx import Document

import docx_styler

if __name__ == '__main__':
    example_json = [
        {
            'InputPath': 'some path',
            'Sentences': [
                {'Text': 'СЛОВО', 'Color': 'Red'},
                {'Text': 'параграф', 'Color': 'green'},
                {'Text': 'со словами', 'Color': 'blue'},
                {'Text': 'Тест со словами', 'Color': 'purple'},
                {'Text': 'после', 'Color': 'darkblue'},
                {'Text': 'Дальше должен идти', 'Color': 'darkyellow'},
                {'Text': 'ййй', 'Color': 'gray'},
                {'Text': 'Тест со словами до', 'Color': 'gray'},
            ],
            'OutputPath': 'same path'
        }
    ]
    doc = Document('test/template.docx')

    sentences = example_json[0]['Sentences']
    unique_sentences = list(
        {sentence['Text']: sentence for sentence in sentences}.values())
    sentences_to_color = sorted(unique_sentences,
                                key=lambda x: len(x['Text']),
                                reverse=True)
    # в зависимости от приоритета изменять значение ключа reverse:
    # если приоритет - покраска более общего значения, которое включает в
    # себя другие - отключить, более узкого - включить
    for sentence in sentences_to_color:
        docx_styler.color_text(doc, sentence['Text'], color=sentence['Color'])

    doc.save('new.docx')
