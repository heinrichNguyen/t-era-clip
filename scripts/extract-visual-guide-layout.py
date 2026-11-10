"""Record the PDF guide's original layout without flattening its text or callouts."""
import ast
import json
from pathlib import Path
from html.parser import HTMLParser
from functools import wraps
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
TMP = ROOT / 'tmp' / 'discussion-pptx'
TMP.mkdir(parents=True, exist_ok=True)
SOURCE = ROOT / 'scripts' / 'build-visual-discussion-guide.py'
slides = []
scope = {'__file__': str(SOURCE), '__name__': '__guide_layout__'}

def color(c):
    return '#' + ''.join(f'{round(v*255):02X}' for v in [c.red, c.green, c.blue]) if c else None

class PlainText(HTMLParser):
    def __init__(self):
        super().__init__(); self.parts = []
    def handle_starttag(self, tag, attrs):
        if tag == 'br': self.parts.append('\n')
    def handle_data(self, data):
        self.parts.append(data)

def wrap(name, fn):
    @wraps(fn)
    def record(*args, **kwargs):
        if name == 'page': slides.append([])
        result = fn(*args, **kwargs)
        op = None
        if name == 'text':
            s, x, y = args[:3]
            size = args[3] if len(args)>3 else kwargs.get('size', 15)
            bold = args[4] if len(args)>4 else kwargs.get('bold', False)
            ink = args[5] if len(args)>5 else kwargs.get('color', scope['NAVY'])
            width = scope['pdfmetrics'].stringWidth(s, 'UIB' if bold else 'UI', size)
            op = dict(type='text', text=s, x=x, y=y, w=width+5, h=size*1.35+2,
                      size=size, bold=bold, color=color(ink), wrap=False)
        elif name == 'para':
            s, x, y, w = args[:4]
            size = args[4] if len(args)>4 else kwargs.get('size', 16)
            bold = args[5] if len(args)>5 else kwargs.get('bold', False)
            ink = args[6] if len(args)>6 else kwargs.get('color', scope['NAVY'])
            plain = PlainText(); plain.feed(s)
            op = dict(type='text', text=''.join(plain.parts), x=x, y=y, w=w, h=result+4,
                      size=size, bold=bold, color=color(ink), wrap=True)
        elif name == 'rect':
            x,y,w,h = args[:4]
            fill = args[4] if len(args)>4 else kwargs.get('fill')
            stroke = args[5] if len(args)>5 else kwargs.get('stroke')
            radius = args[6] if len(args)>6 else kwargs.get('radius',0)
            line = args[7] if len(args)>7 else kwargs.get('line',1)
            op = dict(type='rect',x=x,y=y,w=w,h=h,fill=color(fill),stroke=color(stroke),radius=radius,line=line)
        elif name == 'photo':
            name_,x,y,w = args[:4]
            h = args[4] if len(args)>4 else kwargs.get('h')
            crop = args[5] if len(args)>5 else kwargs.get('crop')
            with Image.open(scope['IMG']/name_) as im:
                if crop: im = im.crop(crop)
                dest = TMP / f'photo-{len(slides):02d}-{len(slides[-1]):02d}.png'
                im.save(dest)
                h = h or w*im.height/im.width
            op = dict(type='image',path=str(dest),x=x,y=y,w=w,h=h,alt=f'Ảnh giao diện với dữ liệu minh họa: {name_}')
        elif name == 'badge':
            n,x,y = args
            op = dict(type='badge',n=n,x=x,y=y)
        elif name == 'arrow':
            op = dict(type='arrow',x1=args[0],y1=args[1],x2=args[2],y2=args[3])
        if op: slides[-1].append(op)
        return result
    return record

for node in ast.parse(SOURCE.read_text()).body:
    # Render an intermediate copy only, never rewrite the existing user PDF.
    if isinstance(node, ast.Assign) and any(isinstance(t,ast.Name) and t.id=='PDF' for t in node.targets):
        scope['PDF'] = TMP / 'reference-render.pdf'
        continue
    exec(compile(ast.Module(body=[node], type_ignores=[]), str(SOURCE), 'exec'), scope)
    if isinstance(node, ast.FunctionDef) and node.name in {'page','text','para','rect','photo','badge','arrow'}:
        scope[node.name] = wrap(node.name, scope[node.name])
assert len(slides)==8
(TMP/'layout.json').write_text(json.dumps({'width':960,'height':540,'slides':slides},ensure_ascii=False,indent=2))
print(f'Recorded {len(slides)} editable slide layouts')
