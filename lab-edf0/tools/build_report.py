from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'Отчет_Лабораторная_работа_EDF0_Траньков.docx'
EXPORTS = ROOT / 'exports'

def set_font(run, size=12, bold=False, name='Times New Roman'):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:ascii'), name)
    run._element.rPr.rFonts.set(qn('w:hAnsi'), name)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def border_cell(cell, color='000000'):
    tcpr = cell._tc.get_or_add_tcPr()
    borders = tcpr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcpr.append(borders)
    for edge in ('top','left','bottom','right','insideH','insideV'):
        tag = qn(f'w:{edge}')
        el = borders.find(tag)
        if el is None:
            el = OxmlElement(f'w:{edge}')
            borders.append(el)
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '6')
        el.set(qn('w:color'), color)

def cell_padding(cell):
    tcpr = cell._tc.get_or_add_tcPr()
    margins = tcpr.first_child_found_in('w:tcMar')
    if margins is None:
        margins = OxmlElement('w:tcMar')
        tcpr.append(margins)
    for side, amount in (('top', 80), ('bottom', 80), ('left', 110), ('right', 110)):
        node = margins.find(qn(f'w:{side}'))
        if node is None:
            node = OxmlElement(f'w:{side}')
            margins.append(node)
        node.set(qn('w:w'), str(amount))
        node.set(qn('w:type'), 'dxa')

def para(doc, text='', size=12, bold=False, align=None, before=0, after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.15
    if align is not None:
        p.alignment = align
    elif not bold:
        p.paragraph_format.first_line_indent = Cm(1.0)
    r = p.add_run(text)
    set_font(r, size, bold)
    return p

def heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(text)
    set_font(r, 15 if level == 1 else 13, True)
    return p

def image(doc, filename, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(str(EXPORTS / filename), width=Cm(16.8))
    c = para(doc, caption, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, after=8)
    c.runs[0].italic = True

doc = Document()
sec = doc.sections[0]
sec.top_margin = Cm(2.0)
sec.bottom_margin = Cm(2.0)
sec.left_margin = Cm(2.5)
sec.right_margin = Cm(1.5)
doc.styles['Normal'].font.name = 'Times New Roman'
doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
doc.styles['Normal'].font.size = Pt(12)

title = doc.add_paragraph(style='Title')
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_after = Pt(9)
run = title.add_run('Лабораторная работа EDF0')
set_font(run, 18, True)
para(doc, 'Построение функциональной модели предметной области', 14, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
para(doc, 'Вариант 18. Учёт выдачи и возврата изданий в библиотеке', 13, align=WD_ALIGN_PARAGRAPH.CENTER, after=28)
para(doc, 'Исполнитель: Роман Траньков', 12, align=WD_ALIGN_PARAGRAPH.RIGHT, after=2)
para(doc, 'Дата: 29.09.2026', 12, align=WD_ALIGN_PARAGRAPH.RIGHT, after=24)

heading(doc, 'Цель работы')
para(doc, 'Освоить построение трёхуровневой функциональной модели в нотации IDEF0 и применить её к учёту библиотечной выдачи и возврата изданий.')
heading(doc, 'Цель, точка зрения и границы модели')
para(doc, 'Цель модели: описать учёт выдачи и возврата библиотечных изданий, чтобы определить функции, данные и роли будущей библиотечной информационной системы.')
para(doc, 'Точка зрения: библиотекарь, который обслуживает читателя, оформляет выдачу и возврат экземпляров и контролирует сроки пользования.')
para(doc, 'Границы системы: модель начинается с обращения читателя или возврата экземпляра и заканчивается регистрацией операции, изменением статуса фонда и формированием уведомления или отчёта. Комплектование фонда, закупка изданий и финансовый учёт штрафов в модель не входят.')

heading(doc, 'Структура модели')
table = doc.add_table(rows=1, cols=3)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.style = 'Table Grid'
headers = ['Узел', 'Название', 'Содержание']
for c, text in zip(table.rows[0].cells, headers):
    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(text); set_font(r, 10, True)
    border_cell(c); cell_padding(c)
rows = [
    ('A-0', 'Учитывать выдачу и возврат изданий', 'Контекстная диаграмма с входами, управлениями, выходами и механизмами.'),
    ('A0', 'Декомпозиция главной функции', 'Проверка заявки, выдача, возврат с контролем срока, уведомления и отчёты.'),
    ('A1', 'Проверять читателя и доступность издания', 'Проверка читателя, экземпляра и возможности выдачи.'),
    ('A2', 'Регистрировать выдачу издания', 'Оформление записи выдачи и передача экземпляра читателю.'),
    ('A3', 'Регистрировать возврат и контролировать срок', 'Приём издания, контроль срока, регистрация возврата и изменение статуса фонда.'),
]
for i,row in enumerate(rows):
    cells=table.add_row().cells
    for j,text in enumerate(row):
        cells[j].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p=cells[j].paragraphs[0]; p.paragraph_format.space_after=Pt(2)
        r=p.add_run(text); set_font(r, 9)
        border_cell(cells[j]); cell_padding(cells[j])

doc.add_page_break()
heading(doc, 'Построенные диаграммы')
para(doc, 'На диаграммах ICOM-коды позволяют проследить миграцию граничных стрелок между уровнями. На A0 показана обратная связь W3 «Просрочка», которая влияет на проверку новой заявки.')
image(doc, 'A-0_контекстная_диаграмма.png', 'Рисунок 1. Контекстная диаграмма A-0')
doc.add_page_break()
image(doc, 'A0_декомпозиция.png', 'Рисунок 2. Декомпозиция главной функции A0')
doc.add_page_break()
image(doc, 'A1_проверка_заявки.png', 'Рисунок 3. Декомпозиция A1 «Проверять читателя и доступность издания»')
doc.add_page_break()
image(doc, 'A2_выдача_издания.png', 'Рисунок 4. Декомпозиция A2 «Регистрировать выдачу издания»')
doc.add_page_break()
image(doc, 'A3_возврат_и_контроль.png', 'Рисунок 5. Декомпозиция A3 «Регистрировать возврат и контролировать срок»')

heading(doc, 'Словарь стрелок')
dict_table = doc.add_table(rows=1, cols=4)
dict_table.style = 'Table Grid'
dict_table.alignment = WD_TABLE_ALIGNMENT.CENTER
for c, text in zip(dict_table.rows[0].cells, ['Код', 'Наименование', 'Тип', 'Определение']):
    c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(text); set_font(r,9,True); border_cell(c); cell_padding(c)
dictionary = [
('I1','Запрос читателя','Вход','Обращение читателя на получение конкретного издания.'),
('I2','Возвращённое издание','Вход','Экземпляр, переданный читателем обратно в библиотеку.'),
('C1','Правила пользования библиотекой','Управление','Нормы регистрации читателя, выдачи, возврата и продления.'),
('C2','Каталог и учётные правила','Управление','Сведения о фонде и правила ведения карточек экземпляров.'),
('C3','Сроки пользования и ограничения','Управление','Допустимые сроки, лимиты и условия ограничения выдачи.'),
('O1','Выданное издание','Выход','Экземпляр, оформленный для передачи читателю.'),
('O2','Обновлённый статус фонда','Выход','Сведения о доступности или занятости экземпляра после операции.'),
('O3','Уведомление читателю','Выход','Сообщение о сроке возврата, просрочке или отказе.'),
('O4','Отчёт по операциям','Выход','Сводные сведения о выдаче, возврате и задолженности.'),
('M1','Библиотекарь','Механизм','Сотрудник, выполняющий обслуживание и учёт.'),
('M2','Библиотечная информационная система','Механизм','Программа для поиска, регистрации операций и формирования отчётов.'),
('M3','Библиотечный фонд','Механизм','Набор экземпляров, доступных для выдачи и возврата.'),
('W1','Подтверждённая заявка','Внутренняя','Результат проверки читателя и доступности издания.'),
('W2','Запись выдачи','Внутренняя','Карточка операции с читателем, экземпляром и сроком возврата.'),
('W3','Сведения о просрочке','Обратная связь','Данные, ограничивающие новую выдачу до урегулирования задолженности.'),
]
for i,row in enumerate(dictionary):
    cells=dict_table.add_row().cells
    for j,text in enumerate(row):
        cells[j].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p=cells[j].paragraphs[0]; p.paragraph_format.space_after=Pt(1); p.paragraph_format.line_spacing=1.0
        r=p.add_run(text); set_font(r,8)
        border_cell(cells[j]); cell_padding(cells[j])

doc.add_page_break()
heading(doc, 'Требования к информационной системе')
para(doc, 'Из модели следует, что система должна идентифицировать читателя, проверять доступность экземпляра и ограничения перед выдачей. Она должна хранить записи выдачи и возврата, автоматически менять статус экземпляра, контролировать сроки пользования и формировать уведомления о просрочке. Отчётность должна использовать сведения о выполненных операциях и задолженности.')
heading(doc, 'Ответы на контрольные вопросы')
answers = [
('1. Что такое функциональная модель?', 'Функциональная модель описывает, какие действия выполняет система, какие данные она получает, что использует как управление и ресурсы, а также какие результаты формирует.'),
('2. Чем цель модели отличается от точки зрения?', 'Цель отвечает на вопрос, для чего строится модель. Точка зрения определяет, чьими глазами рассматривается система. До начала моделирования они задают единое направление работы.'),
('3. Какие типы интерфейсных дуг использует IDEF0?', 'Вход присоединяется слева, управление сверху, выход справа, механизм снизу. Вызов является специальной стрелкой для функции, описанной в другой модели.'),
('4. Как отличить вход от управления?', 'Вход преобразуется работой, а управление задаёт правила её выполнения. Например, возвращённое издание является входом при регистрации возврата, а правила пользования библиотекой являются управлением.'),
('5. Почему на диаграмме должно быть от трёх до шести блоков?', 'Такое число позволяет раскрыть функцию без потери читаемости. Меньшее число обычно не даёт содержательной декомпозиции, большее перегружает диаграмму.'),
('6. Что такое контекстная диаграмма?', 'Это верхний уровень модели с одной главной функцией и граничными стрелками. Ей присваивается номер узла A-0.'),
('7. Как связаны номера блоков и диаграмм?', 'Блок A1 на диаграмме A0 раскрывается на диаграмме A1. Аналогично блоки A2 и A3 имеют одноимённые дочерние диаграммы.'),
('8. Что такое миграция стрелок?', 'Это перенос граничных стрелок родительского блока на дочернюю диаграмму. Корректность проверяют по совпадению ICOM-кодов.'),
('9. Что такое несвязанная стрелка?', 'Это стрелка, которая не соединена с блоком. Она считается синтаксической ошибкой, потому что не участвует в описании функции.'),
('10. Какие связи между работами применяются?', 'Связь по входу, связь по управлению, обратная связь по входу, обратная связь по управлению и связь выход-механизм.'),
('11. Когда применяется обратная связь по управлению?', 'Она применяется, когда результат нижестоящей работы задаёт ограничение или правило для предыдущей. В модели сведения о просрочке влияют на проверку новой заявки.'),
('12. Для чего нужно туннелирование?', 'Оно скрывает стрелку на смежном уровне, если её показ делает диаграмму перегруженной. Туннель обозначается квадратными скобками у начала или конца стрелки.'),
('13. Как именуются ветви стрелок?', 'Стрелку подписывают до точки разветвления. Неподписанные ветви наследуют общее содержимое, а подписанные обозначают его часть.'),
('14. В чём разница AS-IS и TO-BE?', 'AS-IS описывает текущую организацию процесса. TO-BE фиксирует целевую организацию после устранения найденных недостатков.'),
('15. Какие недостатки выявляют в AS-IS?', 'Дублирующиеся и неуправляемые работы, неэффективный документооборот, отсутствие контроля и обратных связей.'),
('16. Чем BPMN отличается от IDEF0?', 'IDEF0 показывает функции и их связи. BPMN описывает последовательность действий, события и ветвления процесса во времени.'),
('17. Какие поля содержит каркас диаграммы?', 'Каркас содержит USED AT, AUTHOR, PROJECT, DATE, REV, NODE, TITLE и NUMBER. Они позволяют идентифицировать диаграмму и её версию.'),
('18. Почему модель хранят в Git?', 'Файл drawio является XML-документом. Git хранит историю правок, позволяет сравнивать версии и возвращаться к предыдущему варианту модели.'),
]
for q,a in answers:
    para(doc, q, 10, True, before=3, after=1)
    para(doc, a, 10, after=3)

heading(doc, 'Вывод')
para(doc, 'Построенная модель разделяет учёт выдачи и возврата на проверку заявки, регистрацию выдачи, регистрацию возврата и формирование результатов. Для автоматизации нужны проверка ограничений читателя, ведение карточек экземпляров, контроль сроков и формирование уведомлений. Исходная модель сохранена в формате drawio, а диаграммы экспортированы в PNG.')

footer=sec.footer.paragraphs[0]
footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
set_font(footer.add_run('Лабораторная работа EDF0, вариант 18'),9)
doc.save(OUT)
print(OUT)
