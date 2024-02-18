from collections import namedtuple
from pprint import pprint

from docx.text.paragraph import Paragraph
from docx.shared import RGBColor
from docx import Document

MyRun = namedtuple('MyRun', 'text style font')

d = Document('template.docx')

red_color = RGBColor(255, 0, 0)
first_el = 0
last_el = -1

# TODO варианты расположения `string` в документе:
#  * `string` целиком в `run`, других слов нет           +
#  * `string` целиком в `run`, есть слово до             +
#  * `string` целиком в `run`, есть слово после          +
#  * `string` целиком в `run`, есть слова до и после     +
#  * `string` разбит на несколько `run`                  -
#  ----------------------------------------------------------------------------
#  Проблема: если `run` с искомым `string` покрашен - цвет не сохраняется
#  Возможно, также не сохраняются другие свойства - проверить

string = 'ЦЕЛИКОМ'
# string = 'СЛОВА'
# string = 'СЛОВЕЧКО'
# string = 'СЛОВО'
# string = 'СЛОВИЩЕ'
# string = 'ФРАЗА ЦЕЛИКОМ'
# string = 'ФРАЗА СО СЛОВАМИ ДО'
# string = 'ФРАЗА СО СЛОВАМИ ПОСЛЕ'
# string = 'ФРАЗА ДО И ПОСЛЕ'

for p in [el for el in d.elements if isinstance(el, Paragraph)]:
    # if word not in p.text:
    #     continue
    pprint([r.text for r in p.runs])

for p in [el for el in d.elements if isinstance(el, Paragraph)]:

    if string not in p.text:
        continue

    if string in [r.text.strip() for r in p.runs]:
        for r in p.runs:
            if string == r.text.strip():
                r.font.color.rgb = RGBColor(255, 0, 0)
                break
        break

    p_runs_new = [MyRun(r.text, r.style, r.font) for r in p.runs]
    p_runs_same = [r for r in p.runs]
    p.clear()
    runs_number = len(p_runs_new)

    count = -1
    for i in range(runs_number):
        r_new = p_runs_new.pop(first_el)
        r_same = p_runs_same.pop(first_el)
        if string not in r_new.text:
            p.append_runs([r_same])
            count += 2
            p.runs[count-1].clear()
            continue
        divided_runs_text = r_new.text.split(string, maxsplit=1)
        run = p.add_run(divided_runs_text[first_el], r_new.style)
        run.element.font = r_new.font
        founded_text_run = p.add_run(string, r_new.style)
        founded_text_run.element.font = r_new.font
        run = p.add_run(divided_runs_text[last_el], r_new.style)
        run.element.font = r_new.font
        founded_text_run.font.color.rgb = red_color
        count += 3
        break
    for r in p_runs_same:
        run = p.append_runs([r])
        count += 2
        p.runs[count-1].clear()
    break

# print('-'*100)
# for p in [el for el in d.elements if isinstance(el, Paragraph)]:
#     if word not in p.text:
#         continue
#     pprint([r.text for r in p.runs])


d.save('new.docx')
