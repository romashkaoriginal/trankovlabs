from pathlib import Path
from copy import deepcopy
import textwrap
import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = Path(r"C:\Users\trank\Downloads\Telegram Desktop\IDEF0_шаблон.drawio")
MODEL_DIR = ROOT / "model"
EXPORT_DIR = ROOT / "exports"
MODEL_DIR.mkdir(exist_ok=True)
EXPORT_DIR.mkdir(exist_ok=True)

AUTHOR = "Роман Траньков"
DATE = "29.09.2026"
MAIN = "Учитывать выдачу и возврат изданий"

PAGES = [
    {
        "node": "A-0", "title": MAIN, "name": "A-0 Контекст",
        "functions": [("Учитывать выдачу и\nвозврат изданий", "0")],
        "inputs": [("I1", "Запрос читателя"), ("I2", "Возвращённое издание")],
        "controls": [("C1", "Правила\nбиблиотеки"), ("C2", "Каталог\nфонда"), ("C3", "Сроки и\nограничения")],
        "outputs": [("O1", "Выданное\nиздание"), ("O2", "Статус\nфонда"), ("O3", "Уведомление"), ("O4", "Отчёт")],
        "mechanisms": [("M1", "Библиотекарь"), ("M2", "Библиотечная\nИС"), ("M3", "Фонд")],
        "internal": [], "file": "A-0_контекстная_диаграмма.png"
    },
    {
        "node": "A0", "title": MAIN, "name": "A0 Декомпозиция",
        "functions": [("Проверять читателя\nи доступность издания", "1"), ("Регистрировать\nвыдачу издания", "2"), ("Регистрировать возврат\nи контролировать срок", "3"), ("Формировать уведомления\nи отчёты", "4")],
        "inputs": [("I1", "Запрос читателя"), ("I2", "Возвращённое издание")],
        "controls": [("C1", "Правила\nбиблиотеки"), ("C2", "Каталог\nфонда"), ("C3", "Сроки и\nограничения")],
        "outputs": [("O1", "Выданное издание"), ("O2", "Статус фонда"), ("O3", "Уведомление"), ("O4", "Отчёт")],
        "mechanisms": [("M1", "Библиотекарь"), ("M2", "Библиотечная\nИС"), ("M3", "Фонд")],
        "internal": [(0, 1, "W1 Заявка"), (1, 2, "W2 Выдача"), (2, 0, "W3 Просрочка")],
        "file": "A0_декомпозиция.png"
    },
    {
        "node": "A1", "title": "Проверять читателя и доступность издания", "name": "A1 Проверка заявки",
        "functions": [("Идентифицировать\nчитателя", "1"), ("Проверять доступность\nэкземпляра", "2"), ("Подтверждать заявку\nили оформлять отказ", "3")],
        "inputs": [("I1", "Запрос читателя")],
        "controls": [("C1", "Правила\nбиблиотеки"), ("C2", "Каталог\nфонда"), ("C3", "Сроки и\nограничения")],
        "outputs": [("W1", "Подтверждённая заявка"), ("O3", "Уведомление об отказе")],
        "mechanisms": [("M1", "Библиотекарь"), ("M2", "Библиотечная\nИС"), ("M3", "Фонд")],
        "internal": [(0, 1, "Статус читателя"), (1, 2, "Доступный экземпляр")],
        "file": "A1_проверка_заявки.png"
    },
    {
        "node": "A2", "title": "Регистрировать выдачу издания", "name": "A2 Выдача издания",
        "functions": [("Получать подтверждённую\nзаявку", "1"), ("Оформлять запись\nвыдачи", "2"), ("Передавать экземпляр\nчитателю", "3")],
        "inputs": [("W1", "Подтверждённая заявка")],
        "controls": [("C1", "Правила\nбиблиотеки"), ("C3", "Сроки и\nограничения")],
        "outputs": [("O1", "Выданное издание"), ("W2", "Запись выдачи")],
        "mechanisms": [("M1", "Библиотекарь"), ("M2", "Библиотечная\nИС"), ("M3", "Фонд")],
        "internal": [(0, 1, "Данные выдачи"), (1, 2, "Оформленная операция")],
        "file": "A2_выдача_издания.png"
    },
    {
        "node": "A3", "title": "Регистрировать возврат и контролировать срок", "name": "A3 Возврат и контроль срока",
        "functions": [("Принимать возвращённое\nиздание", "1"), ("Проверять состояние\nи срок возврата", "2"), ("Регистрировать возврат\nи ограничения", "3"), ("Обновлять статус\nфонда", "4")],
        "inputs": [("I2", "Возвращённое издание"), ("W2", "Запись выдачи")],
        "controls": [("C1", "Правила\nбиблиотеки"), ("C3", "Сроки и\nограничения")],
        "outputs": [("O2", "Обновлённый статус фонда"), ("W3", "Сведения о просрочке")],
        "mechanisms": [("M1", "Библиотекарь"), ("M2", "Библиотечная\nИС"), ("M3", "Фонд")],
        "internal": [(0, 1, "Принятое издание"), (1, 2, "Результат проверки"), (2, 3, "Запись возврата")],
        "file": "A3_возврат_и_контроль.png"
    },
]

def add_cell(root, cid, value, style, x, y, w, h, vertex=True):
    attrs = {"id": cid, "value": value, "style": style, "vertex": "1", "parent": "1"}
    cell = ET.SubElement(root, "mxCell", attrs)
    ET.SubElement(cell, "mxGeometry", {"x": str(x), "y": str(y), "width": str(w), "height": str(h), "as": "geometry"})
    return cell

def add_edge(root, cid, value, x1, y1, x2, y2):
    style = "edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;endArrow=block;endFill=1;strokeWidth=1.5;"
    cell = ET.SubElement(root, "mxCell", {"id": cid, "value": value, "style": style, "edge": "1", "parent": "1"})
    geo = ET.SubElement(cell, "mxGeometry", {"relative": "1", "as": "geometry"})
    ET.SubElement(geo, "mxPoint", {"x": str(x1), "y": str(y1), "as": "sourcePoint"})
    ET.SubElement(geo, "mxPoint", {"x": str(x2), "y": str(y2), "as": "targetPoint"})
    return cell

def add_text(root, cid, value, x, y, w=160, h=28, align="center"):
    style = f"text;html=1;strokeColor=none;fillColor=none;align={align};verticalAlign=middle;whiteSpace=wrap;overflow=hidden;fontSize=10;"
    return add_cell(root, cid, value.replace("\n", "<br>"), style, x, y, w, h)

def update_template_page(diagram, page, index):
    model = diagram.find("mxGraphModel")
    root = model.find("root")
    keep_until = [19, 44, 71, 98, 125][index]
    for c in list(root):
        cid = c.attrib.get("id", "")
        if cid.startswith("n") and cid[1:].isdigit() and int(cid[1:]) > keep_until:
            root.remove(c)
        elif not cid.startswith("n") and cid not in {"0", "1"}:
            root.remove(c)
    lookup = {c.attrib.get("id"): c for c in root}
    base = [8, 33, 60, 87, 114][index]
    lookup[f"n{base+1}"].set("value", f"AUTHOR: {AUTHOR}<br>PROJECT: ЛР IDEF0, вариант 18<br>DATE: {DATE} &nbsp;&nbsp; REV: 1")
    lookup[f"n{base+7}"].set("value", page["node"])
    lookup[f"n{base+9}"].set("value", page["title"])
    lookup[f"n{base+11}"].set("value", f"ЛР-IDEF0-18-{index+1}")
    # Node title and diagrams content
    block_style = "rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=2;fontSize=13;align=center;verticalAlign=middle;"
    n = 1
    if len(page["functions"]) == 1:
        positions = [(420, 320, 310, 130)]
    elif len(page["functions"]) == 3:
        positions = [(240, 270, 210, 100), (490, 345, 210, 100), (740, 420, 210, 100)]
    else:
        positions = [(170, 230, 190, 90), (395, 310, 190, 90), (620, 390, 190, 90), (845, 470, 190, 90)]
    for (label, no), (x, y, w, h) in zip(page["functions"], positions):
        bid = f"p{index}_b{n}"
        add_cell(root, bid, f"{label.replace(chr(10), '<br>')}<br><font style='font-size:10px'>A{page['node'].replace('A-','').replace('A','')}{no if page['node'] != 'A-0' else ''}</font>", block_style, x, y, w, h)
        n += 1
    # External arrows around whole model
    center = positions[0] if len(positions)==1 else positions[len(positions)//2]
    bx, by, bw, bh = center
    for k,(code,label) in enumerate(page["inputs"]):
        y = by + 20 + k*28
        add_edge(root, f"p{index}_i{k}", "", 55, y, bx, y)
        add_text(root, f"p{index}_it{k}", f"{code} {label}", 60, y-25, max(120,bx-75), 22, "left")
    for k,(code,label) in enumerate(page["controls"]):
        x = bx + 10 + k*min(100, max(55,(bw-25)//max(1,len(page['controls']))))
        add_edge(root, f"p{index}_c{k}", "", x, 100, x, by)
        add_text(root, f"p{index}_ct{k}", f"{code} {label}", x-70, 108, 140, 30)
    for k,(code,label) in enumerate(page["outputs"]):
        y = by + 18 + k*25
        add_edge(root, f"p{index}_o{k}", "", bx+bw, y, 1100, y)
        add_text(root, f"p{index}_ot{k}", f"{code} {label}", bx+bw+5, y-25, max(130,1080-(bx+bw)), 28, "left")
    for k,(code,label) in enumerate(page["mechanisms"]):
        x = bx + 20 + k*min(100, max(55,(bw-30)//max(1,len(page['mechanisms']))))
        add_edge(root, f"p{index}_m{k}", "", x, 700, x, by+bh)
        add_text(root, f"p{index}_mt{k}", f"{code} {label}", x-70, 652, 140, 30)
    # Internal flow labels and arrows across blocks
    for k,(src,dst,label) in enumerate(page["internal"]):
        sx,sy,sw,sh = positions[src]
        dx,dy,dw,dh = positions[dst]
        if dst < src:
            # feedback loop above the blocks
            add_edge(root, f"p{index}_w{k}", "", sx+sw/2, sy, dx+dw/2, dy)
            add_text(root, f"p{index}_wt{k}", label, min(sx,dx)+20, min(sy,dy)-38, 170, 24)
        else:
            add_edge(root, f"p{index}_w{k}", "", sx+sw, sy+sh/2, dx, dy+dh/2)
            add_text(root, f"p{index}_wt{k}", label, (sx+sw+dx)/2-75, (sy+sh/2+dy+dh/2)/2-30, 150, 24)

def font(size, bold=False):
    candidates = [r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\timesbd.ttf" if bold else r"C:\Windows\Fonts\times.ttf"]
    for c in candidates:
        if Path(c).exists(): return ImageFont.truetype(c, size)
    return ImageFont.load_default()

def multiline(draw, xy, text, fnt, fill, width, align="center"):
    lines=[]
    for para in text.split("\n"):
        lines.extend(textwrap.wrap(para, width=width) or [""])
    rendered = "\n".join(lines)
    left, top, right, bottom = draw.multiline_textbbox(xy, rendered, font=fnt, anchor="mm", align=align, spacing=4)
    draw.rectangle((left-4, top-3, right+4, bottom+3), fill="white")
    draw.multiline_text(xy, rendered, font=fnt, fill=fill, anchor="mm", align=align, spacing=4)

def draw_arrow(draw, p1, p2, width=3):
    draw.line([p1,p2], fill="#202020", width=width)
    x,y=p2; dx=x-p1[0]; dy=y-p1[1]
    if abs(dx)>=abs(dy):
        sign=1 if dx>0 else -1
        tri=[(x,y),(x-12*sign,y-7),(x-12*sign,y+7)]
    else:
        sign=1 if dy>0 else -1
        tri=[(x,y),(x-7,y-12*sign),(x+7,y-12*sign)]
    draw.polygon(tri, fill="#202020")

def render_page(page):
    img=Image.new("RGB", (1800, 1270), "white")
    d=ImageDraw.Draw(img)
    # frame and header/footer
    d.rectangle((60,50,1740,1170), outline="#111111", width=3)
    d.line((60,150,1740,150), fill="#111111", width=2)
    d.line((60,1075,1740,1075), fill="#111111", width=2)
    d.text((85,75), f"AUTHOR: {AUTHOR}", font=font(22), fill="#111111")
    d.text((85,105), f"PROJECT: ЛР IDEF0, вариант 18     DATE: {DATE}     REV: 1", font=font(18), fill="#111111")
    d.text((900,105), page["name"], font=font(20, True), fill="#111111", anchor="mm")
    d.text((85,1105), f"NODE: {page['node']}", font=font(21, True), fill="#111111")
    d.text((900,1105), f"TITLE: {page['title']}", font=font(18), fill="#111111", anchor="mm")
    d.text((1710,1105), "ЛР-IDEF0-18", font=font(18), fill="#111111", anchor="rm")
    count=len(page['functions'])
    if count==1: pos=[(700,500,430,180)]
    elif count==3: pos=[(360,390,320,150),(740,520,320,150),(1120,650,320,150)]
    else: pos=[(220,330,270,135),(545,445,270,135),(870,560,270,135),(1195,675,270,135)]
    # external first
    cidx=min(len(pos)-1,1)
    bx,by,bw,bh=pos[cidx]
    for k,(code,label) in enumerate(page['inputs']):
        yy=by+30+k*85; draw_arrow(d,(80,yy),(bx,yy)); multiline(d,(165,yy-42),f"{code} {label}",font(15),"#111",20)
    for k,(code,label) in enumerate(page['controls']):
        xx=bx+25+k*max(92,(bw-50)//max(1,len(page['controls']))); draw_arrow(d,(xx,175),(xx,by)); multiline(d,(xx,215),f"{code} {label}",font(15),"#111",13)
    for k,(code,label) in enumerate(page['outputs']):
        yy=by+30+k*55; draw_arrow(d,(bx+bw,yy),(1715,yy)); multiline(d,(1605,yy-30),f"{code} {label}",font(14),"#111",13)
    for k,(code,label) in enumerate(page['mechanisms']):
        xx=bx+25+k*135; draw_arrow(d,(xx,1035),(xx,by+bh)); multiline(d,(xx,970 if k % 2 == 0 else 1010),f"{code} {label}",font(15),"#111",18)
    # internal arrows
    for src,dst,label in page['internal']:
        sx,sy,sw,sh=pos[src]; dx,dy,dw,dh=pos[dst]
        if dst<src:
            d.line([(sx+sw//2,sy),(sx+sw//2,760),(dx+dw//2,760),(dx+dw//2,dy)], fill="#202020",width=3)
            draw_arrow(d,(dx+dw//2,760),(dx+dw//2,dy))
            multiline(d,((sx+dx)//2+110,735),label,font(15),"#111",22)
        else:
            draw_arrow(d,(sx+sw,sy+sh//2),(dx,dy+dh//2)); multiline(d,((sx+sw+dx)//2,(sy+sh//2+dy+dh//2)//2-30),label,font(15),"#111",20)
    # functions top
    for (label,no),(x,y,w,h) in zip(page['functions'],pos):
        d.rectangle((x,y,x+w,y+h),outline="#111111",width=4,fill="#f8fbff")
        multiline(d,(x+w//2,y+h//2),label,font(23,True),"#111",19)
        d.text((x+w-16,y+h-15), no, font=font(20,True), fill="#111", anchor="rm")
    img.save(EXPORT_DIR/page['file'])

if __name__ == '__main__':
    tree=ET.parse(TEMPLATE)
    mxfile=tree.getroot()
    for i,(diagram,page) in enumerate(zip(mxfile.findall('diagram'),PAGES)):
        diagram.set('name',page['name'])
        update_template_page(diagram,page,i)
        render_page(page)
    ET.indent(tree, space="  ")
    tree.write(MODEL_DIR/'roman_trankov_idef0_library.drawio', encoding='utf-8', xml_declaration=True)
    print('Built source model and five PNG exports')
