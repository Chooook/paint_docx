
https://habr.com/ru/articles/663028/

https://stackoverflow.com/questions/48654715/page-break-via-python-docx-in-ms-word-docx-file-appears-only-at-the-end


Объект Document представляет собой весь документ – его структура:

* Список объектов `paragraph` – абзацы документа
  * Список объектов `run` – фрагменты текста с различными стилями 
  форматирования (курсив, цвет шрифта и т.п.)
* Список объектов `table` – таблицы документа
  * Список объектов `row` – строки таблицы
    * Список объектов `cell` – ячейки в строке
      * Список объектов `cell.paragraphs` содержит все абзацы в ячейке
  * Список объектов `column` – столбцы таблицы
    * Список объектов `cell` – ячейки в столбце
      * Список объектов `cell.paragraphs` содержит все абзацы в ячейке
* Список объектов `InlineShape` – иллюстрации документа

``` Py
# добавляем разрыв страницы
doc.add_page_break()
```
``` Py
# данные таблицы без названий колонок

items = (
    (1, 'первая строка', 'первая строка'),
    (2, 'вторая строка', 'вторая строка'),
    (3, 'третья строка', 'третья строка'),
)
```
``` Py
# добавляем таблицу с одной строкой 
# для заполнения названий колонок
table = doc.add_table(1, len(items[0]))
```
``` Py
# определяем стиль таблицы
table.style = 'Light Shading Accent 1'
```
``` Py
# Получаем строку с колонками из добавленной таблицы
head_cells = table.rows[0].cells
```
``` Py
# добавляем названия колонок
for i, item in enumerate(['первая колонка', 'вторая колонка', 'третья колонка']):
    p = head_cells[i].paragraphs[0]
    # название колонки
    p.add_run(item).bold = True
    # выравниваем посередине
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
```
``` Py
# добавляем данные к существующей таблице
for row in items:
    # добавляем строку с ячейками к объекту таблицы
    cells = table.add_row().cells
    for i, item in enumerate(row):
        # вставляем данные в ячейки
        cells[i].text = str(item)
```
``` Py
def get_display_name():
    """
    This function return full name of user.
    out:
        string: full name of user
    """
    get_user_name_ex = ctypes.windll.secur32.GetUserNameExW
    name_display = 3
    size = ctypes.pointer(ctypes.c_ulong(0))
    get_user_name_ex(name_display, None, size)
    name_buffer = ctypes.create_unicode_buffer(size.contents.value)
    get_user_name_ex(name_display, name_buffer, size)
    return name_buffer.value
```
И, наконец, основная функция поиска строки и добавления комментария:
``` Py
def make_comment(text:str, paragraphs:list, user:str):
    """
    This function adds comments in docx files.
    :param text: the line we are looking for
    :param paragraphs: list of paragraphs to search for a string
    :param user: full name of user
    """
    for paragraph in paragraphs:
        if type(paragraph) == list: 
            text_in_table = [p.text for p in paragraph]
            text_in_table = ''.join(text_in_table) 
            if len(text_in_table) >= len(text)-5: 
                res = fuzz.partial_ratio(text.lower(), text_in_table.lower())
                if res >= 97: 
                    p = paragraph[-1]
                    run = p.add_run()
                    run.add_comment('Строчка которую искали', author=user) 
        else: 
            if len(paragraph.text) >= len(text):
                res = fuzz.partial_ratio(text.lower(), paragraph.text.lower())
                if res >= 97:
                    paragraph.add_comment('Строчка которую искали', author=user)
```

Разрывы страницы бывают двух видов:

1)  hard breaks – разрывы, вставленные с помощью Ctrl + Enter.

2)  soft breaks – разрывы, вставленные, когда автор печатал текст, и произошел автоматический переход на новую страницу.

Обнаружить эти два вида разрыва страницы можно в xml-разметке объектов run (run._element.xml), которые есть у каждого объекта paragraph.

Напишем функцию определения номера страницы для искомой строки:
``` Py
def number_page(text:str, paragraphs:list):
    """
    This funcion find number page.`

    :param text: string what we find
    :param paragraphs: list of paragraphs
    :return: pages
    """
    pages = []
    number_page = 1
    for paragraph in paragraphs:
        if type(paragraph) == list: 
            text_in_table = [p.text for p in paragraph]
            text_in_table = ''.join(text_in_table)
            for p in paragraph:
                for run in p.runs: 
                    if 'lastRenderedPageBreak' in run._element.xml:
                        number_page += 1 
                    elif 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                        number_page += 1
            if len(text_in_table) >= len(text)-10:
                res = fuzz.partial_ratio(text.lower(), text_in_table.lower())
                if res >= 97: 
                    pages.append(number_page)
        else: 
            for run in paragraph.runs:
                if 'lastRenderedPageBreak' in run._element.xml:
                    number_page += 1
                elif 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                    number_page += 1
            if len(paragraph.text) >= len(text):
                res = fuzz.partial_ratio(text.lower(), paragraph.text.lower())
                if res >= 97:
                    pages.append(number_page)
    return ', '.join(map(str, pages)
```
Вот здесь XML-разметка объекта run проверяется на наличие тегов, указывающих на наличие разрыва страницы:
``` Py
if 'lastRenderedPageBreak' in run._element.xml:
    number_page += 1 
elif 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
    number_page += 1
```
Как и в прошлый раз, подготавливаем данные для функции:

Для нахождения номера страницы нужно передать в функцию number_page первым аргументом строку, которую искали, вторым – список paragraphs.
``` Py
print(number_page(text, paragraphs))
```
