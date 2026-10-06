"""Build the editable draw.io model and its PNG views from one geometry plan."""
from pathlib import Path
import textwrap
import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw, ImageFont
from build_idef0_assets import PAGES, TEMPLATE, MODEL_DIR, EXPORT_DIR, AUTHOR, DATE

S = 2
W, H = 1169, 826
INK = "#171717"
OUTPUT = MODEL_DIR / "roman_trankov_idef0_library.drawio"

def fnt(size=12, bold=False):
    path = Path(r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf")
    return ImageFont.truetype(str(path), size * S)

def geom(page):
    n = len(page["functions"])
    if n == 1:
        return [(430, 310, 300, 140)]
    if n == 3:
        return [(175, 185, 210, 98), (465, 310, 210, 98), (755, 435, 210, 98)]
    return [(145, 165, 180, 90), (385, 270, 180, 90), (625, 375, 180, 90), (865, 480, 180, 90)]

def text_bounds(draw, x, y, value, size, bold=False, anchor="la", spacing=3):
    return draw.multiline_textbbox((x*S,y*S), value, font=fnt(size,bold), anchor=anchor, spacing=spacing*S)

class Canvas:
    def __init__(self, root):
        self.root = root
        self.img = Image.new("RGB", (W*S, H*S), "white")
        self.d = ImageDraw.Draw(self.img)
        self.count = 0

    def ident(self):
        self.count += 1
        return "clean" + str(self.count)

    def line(self, points, arrow=False, weight=1):
        p = [(round(x*S),round(y*S)) for x,y in points]
        self.d.line(p, fill=INK, width=max(2,weight*S), joint="curve")
        if arrow:
            x,y=p[-1]; px,py=p[-2]
            dx,dy=x-px,y-py
            if abs(dx)>=abs(dy):
                sign=1 if dx>0 else -1
                tri=[(x,y),(x-8*S*sign,y-4*S),(x-8*S*sign,y+4*S)]
            else:
                sign=1 if dy>0 else -1
                tri=[(x,y),(x-4*S,y-8*S*sign),(x+4*S,y-8*S*sign)]
            self.d.polygon(tri, fill=INK)
        style = "edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;strokeColor=#171717;strokeWidth=1;" + ("endArrow=block;endFill=1;" if arrow else "endArrow=none;")
        cell=ET.SubElement(self.root,"mxCell",{"id":self.ident(),"style":style,"edge":"1","parent":"1"})
        g=ET.SubElement(cell,"mxGeometry",{"relative":"1","as":"geometry"})
        ET.SubElement(g,"mxPoint",{"x":str(points[0][0]),"y":str(points[0][1]),"as":"sourcePoint"})
        if len(points)>2:
            arr=ET.SubElement(g,"Array",{"as":"points"})
            for xx,yy in points[1:-1]:
                ET.SubElement(arr,"mxPoint",{"x":str(xx),"y":str(yy)})
        ET.SubElement(g,"mxPoint",{"x":str(points[-1][0]),"y":str(points[-1][1]),"as":"targetPoint"})

    def label(self, x, y, value, size=11, bold=False, width=None, height=None, anchor="la"):
        # Text is placed in whitespace beside the edge; no opaque label background.
        value=value.replace("<br>","\n")
        self.d.multiline_text((x*S,y*S),value,font=fnt(size,bold),fill=INK,anchor=anchor,spacing=3*S,align="center" if anchor=="ma" else "left")
        b=text_bounds(self.d,x,y,value,size,bold,anchor)
        if width is None: width=max(25,(b[2]-b[0])/S+5)
        if height is None: height=max(16,(b[3]-b[1])/S+5)
        style=("text;html=1;strokeColor=none;fillColor=none;whiteSpace=wrap;"
               f"align={'center' if anchor=='mt' else 'left'};verticalAlign=middle;fontSize={size};")
        xx=x-width/2 if anchor=="ma" else x
        yy=y
        cell=ET.SubElement(self.root,"mxCell",{"id":self.ident(),"value":value.replace("\n","<br>"),"style":style,"vertex":"1","parent":"1"})
        ET.SubElement(cell,"mxGeometry",{"x":str(round(xx)),"y":str(round(yy)),"width":str(round(width)),"height":str(round(height)),"as":"geometry"})

    def block(self, x,y,w,h,title,no):
        title="\n".join(line for part in title.split("\n") for line in textwrap.wrap(part, width=19, break_long_words=False))
        size=13
        while size>10:
            box=self.d.multiline_textbbox((0,0),title,font=fnt(size,True),spacing=4*S,align="center")
            if box[2]-box[0] <= (w-18)*S and box[3]-box[1] <= (h-22)*S:
                break
            size-=1
        self.d.rectangle((x*S,y*S,(x+w)*S,(y+h)*S),outline=INK,width=2*S,fill="white")
        self.d.multiline_text(((x+w/2)*S,(y+h/2)*S),title,font=fnt(size,True),fill=INK,anchor="mm",align="center",spacing=4*S)
        self.d.text(((x+w-9)*S,(y+h-8)*S),no,font=fnt(10,True),fill=INK,anchor="rd")
        cell=ET.SubElement(self.root,"mxCell",{"id":self.ident(),"value":title.replace("\n","<br>")+f"<br><font style='font-size:10px'>{no}</font>","style":f"rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#171717;strokeWidth=2;fontSize={size};fontStyle=1;align=center;verticalAlign=middle;","vertex":"1","parent":"1"})
        ET.SubElement(cell,"mxGeometry",{"x":str(x),"y":str(y),"width":str(w),"height":str(h),"as":"geometry"})

def draw_frame(c,page,index):
    d=c.d
    def line(a,b): d.line((a[0]*S,a[1]*S,b[0]*S,b[1]*S),fill=INK,width=2)
    d.rectangle((40*S,40*S,1120*S,740*S),outline=INK,width=2*S)
    for a,b in [((40,102),(1120,102)),((40,695),(1120,695)),
                ((200,40),(200,102)),((510,40),(510,102)),((775,40),(775,102)),
                ((190,695),(190,740)),((955,695),(955,740))]:
        line(a,b)
    c.label(50,48,"USED AT: БИС",10)
    c.label(210,46,f"AUTHOR: {AUTHOR}\nPROJECT: ЛР IDEF0, вариант 18\nDATE: {DATE}   REV: 1",9)
    c.label(520,47,"WORKING\nDRAFT [X]\nRECOMMENDED",9)
    c.label(785,47,"READER: —\nDATE: "+DATE+"\nCONTEXT: "+page["node"],9)
    c.label(50,701,"NODE:",10)
    c.label(92,715,page["node"],13,True)
    c.label(205,701,"TITLE:",10)
    c.label(320,714,page["title"],12,True)
    c.label(966,701,"NUMBER:",10)
    c.label(972,718,f"ЛР-18-{index+1}",10)

def trim_frame(root,index,page):
    # Retain the supplied four-column header and three-part footer.
    maxid=[19,44,71,98,125][index]
    for cell in list(root):
        cid=cell.get("id","")
        if cid in {"0","1"}: continue
        if not cid.startswith("n") or not cid[1:].isdigit() or int(cid[1:])>maxid:
            root.remove(cell)
    base=[8,33,60,87,114][index]
    cells={x.get("id"):x for x in root}
    cells[f"n{base}"].set("value","USED AT: БИС")
    cells[f"n{base+1}"].set("value",f"AUTHOR: {AUTHOR}<br>PROJECT: ЛР IDEF0, вариант 18<br>DATE: {DATE}   REV: 1")
    cells[f"n{base+2}"].set("value","WORKING<br>DRAFT [X]<br>RECOMMENDED")
    cells[f"n{base+3}"].set("value",f"READER: —<br>DATE: {DATE}<br>CONTEXT: {page['node']}")
    cells[f"n{base+7}"].set("value",page["node"])
    cells[f"n{base+9}"].set("value",page["title"])
    cells[f"n{base+11}"].set("value",f"ЛР-18-{index+1}")

def ext_input(c, box, code, name, slot=0):
    x,y,w,h=box
    yy=y+25+slot*50
    c.line([(55,yy),(x,yy)],True)
    short=(code+"  "+name).replace("Запрос читателя","Запрос\nчитателя").replace("Возвращённое издание","Возвращённое\nиздание").replace("Подтверждённая заявка","Подтверждённая\nзаявка")
    c.label(58,yy+5,short,9)

def ext_control(c, box, code, name, slot=0):
    x,y,w,h=box
    xx=x+round(w*(0.32+slot*0.3))
    c.line([(xx,112),(xx,y)],True)
    c.label(xx+8,118,code+"  "+name,9)

def ext_output(c, box, code, name, slot=0):
    x,y,w,h=box
    yy=y+28+slot*30
    c.line([(x+w,yy),(1105,yy)],True)
    # The caption sits below its own edge, with enough clearance before the next one.
    if x+w > 1000:
        c.label(x+w+2,yy+5,code+"  "+name,8)
    else:
        label=(code+"  "+name).replace("Подтверждённая заявка","Подтверждённая\nзаявка").replace("Выданное издание","Выданное\nиздание")
        c.label(max(x+w+8,930),yy+5,label,9)

def ext_mechanism(c, box, code, name, slot=0):
    x,y,w,h=box
    xx=x+([45,160,280][slot] if w>250 else round(w*(.2+.3*slot)))
    c.line([(xx,660),(xx,y+h)],True)
    caption=code+"  "+name
    if w>250:
        caption=caption.replace("Библиотечная ИС","Библиотечная\nИС")
    c.label(xx+9,660,caption,8)

def flow(c, a,b,name,compact=False):
    ax,ay,aw,ah=a; bx,by,bw,bh=b
    y1=ay+round(ah*.68); y2=by+round(bh*.38)
    xx=ax+aw+round((bx-ax-aw)/2)
    c.line([(ax+aw,y1),(xx,y1),(xx,y2),(bx,y2)],True)
    # Put the caption in the open space above the receiving box, to the
    # right of its control edge.  It never crosses the elbow of the flow.
    if compact:
        c.label(ax+aw+6,y1+14,name.split()[0],9)
    else:
        control_x=bx+round(bw*.32)
        c.label(control_x+10,by-29,name,9)

def make_page(c,page,index):
    boxes=geom(page)
    for (title,no),box in zip(page["functions"],boxes): c.block(*box,title,no)
    if index==0:
        for i,(code,name) in enumerate(page["inputs"]): ext_input(c,boxes[0],code,name,i)
        for i,(code,name) in enumerate(page["controls"]): ext_control(c,boxes[0],code,name,i)
        for i,(code,name) in enumerate(page["outputs"]): ext_output(c,boxes[0],code,name,i)
        for i,(code,name) in enumerate(page["mechanisms"]): ext_mechanism(c,boxes[0],code,name,i)
        return
    if index==1:
        ext_input(c,boxes[0],"I1","Запрос читателя")
        ext_input(c,boxes[2],"I2","Возвращённое издание")
        # The child diagrams use the same external controls and mechanisms.
        for box in boxes[:3]:
            for slot,code in enumerate(("C1","C2","C3")):
                ext_control(c,box,code,"",slot)
        ext_control(c,boxes[3],"C1","Правила библиотеки")
        for box,code,name,slot in [(boxes[1],"O1","Выданное издание",0),(boxes[2],"O2","Статус фонда",0),(boxes[3],"O3","Уведомление",0),(boxes[3],"O4","Отчёт",1)]:
            ext_output(c,box,code,name,slot)
        for box in boxes[:3]:
            for slot,code in enumerate(("M1","M2","M3")):
                ext_mechanism(c,box,code,"",slot)
        ext_mechanism(c,boxes[3],"M2","Библиотечная ИС")
        flow(c,boxes[0],boxes[1],"W1  Решение по заявке",True)
        flow(c,boxes[1],boxes[2],"W2  Запись выдачи",True)
        flow(c,boxes[2],boxes[3],"W3  Сведения о просрочке",True)
        return
    if index==2:
        ext_input(c,boxes[0],"I1","Запрос читателя")
        for box,code,name in [(boxes[0],"C1","Правила библиотеки"),(boxes[1],"C2","Каталог фонда"),(boxes[2],"C3","Сроки и ограничения")]:
            ext_control(c,box,code,name)
        ext_output(c,boxes[2],"W1","Решение по заявке")
        for box,code,name in [(boxes[0],"M1","Библиотекарь"),(boxes[1],"M3","Фонд"),(boxes[2],"M2","Библиотечная ИС")]:
            ext_mechanism(c,box,code,name)
        flow(c,boxes[0],boxes[1],"Статус читателя")
        flow(c,boxes[1],boxes[2],"Доступный экземпляр")
        return
    if index==3:
        ext_input(c,boxes[0],"W1","Решение по заявке")
        for box,code,name in [(boxes[0],"C1","Правила библиотеки"),(boxes[1],"C2","Каталог фонда"),(boxes[2],"C3","Сроки и ограничения")]:
            ext_control(c,box,code,name)
        ext_output(c,boxes[2],"O1","Выданное издание")
        ext_output(c,boxes[1],"W2","Запись выдачи")
        for box,code,name in [(boxes[0],"M2","Библиотечная ИС"),(boxes[1],"M3","Фонд"),(boxes[2],"M1","Библиотекарь")]:
            ext_mechanism(c,box,code,name)
        flow(c,boxes[0],boxes[1],"Данные выдачи")
        flow(c,boxes[1],boxes[2],"Оформленная операция")
        return
    ext_input(c,boxes[0],"I2","Возвращённое издание")
    ext_input(c,boxes[1],"W2","Запись выдачи")
    for box,code,name in [(boxes[0],"C1","Правила библиотеки"),(boxes[1],"C3","Сроки пользования"),(boxes[2],"C1","Правила библиотеки"),(boxes[3],"C2","Каталог фонда")]:
        ext_control(c,box,code,name)
    ext_output(c,boxes[3],"O2","Статус фонда")
    ext_output(c,boxes[2],"W3","Сведения о просрочке")
    for box,code,name in [(boxes[0],"M1","Библиотекарь"),(boxes[1],"M2","Библиотечная ИС"),(boxes[2],"M2","Библиотечная ИС"),(boxes[3],"M3","Фонд")]:
        ext_mechanism(c,box,code,name)
    flow(c,boxes[0],boxes[1],"Принятое издание")
    flow(c,boxes[1],boxes[2],"Результат проверки")
    flow(c,boxes[2],boxes[3],"Запись возврата")

tree=ET.parse(TEMPLATE)
for i,(diagram,page) in enumerate(zip(tree.getroot().findall("diagram"),PAGES)):
    diagram.set("name",page["name"])
    root=diagram.find("mxGraphModel").find("root")
    trim_frame(root,i,page)
    c=Canvas(root)
    draw_frame(c,page,i)
    make_page(c,page,i)
    c.img.save(EXPORT_DIR/page["file"])
ET.indent(tree,space="  ")
tree.write(OUTPUT,encoding="utf-8",xml_declaration=True)
print(OUTPUT)
