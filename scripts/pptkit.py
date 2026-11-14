"""pptkit：在模板上生成页面和图示。

用法见 references/render.md。单位都是英寸。颜色写“风格角色名”（primary、accent、text、
muted、line、tint、surface、on_primary），由风格文件映射到主题色或色值，不在页面代码里写色值。
"""
import copy
import json
import math
import os
import re

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

EMU = 914400
LS = 1.1    # 行距倍数
LH = 1.3    # 中文字体一行的实际高度约为字号的 1.3 倍
NO_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
THEME = {
    "tx1": MSO_THEME_COLOR.TEXT_1, "tx2": MSO_THEME_COLOR.TEXT_2,
    "bg1": MSO_THEME_COLOR.BACKGROUND_1, "bg2": MSO_THEME_COLOR.BACKGROUND_2,
    "accent1": MSO_THEME_COLOR.ACCENT_1, "accent2": MSO_THEME_COLOR.ACCENT_2,
    "accent3": MSO_THEME_COLOR.ACCENT_3, "accent4": MSO_THEME_COLOR.ACCENT_4,
    "accent5": MSO_THEME_COLOR.ACCENT_5, "accent6": MSO_THEME_COLOR.ACCENT_6,
}
SIZES = {"name": 14, "h": 13, "body": 11.5, "small": 10, "label": 10.5, "num": 22, "stat": 40}
ANCHOR = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}
ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}


# ------------------------------------------------------------------ 文字量算
def em(text):
    n = 0.0
    for ch in text:
        if ch == " ":
            n += 0.3
        elif ord(ch) < 0x2E80:
            n += 0.56
        else:
            n += 1.0
    return n


def lines(text, size, w):
    total = 0
    for part in str(text).split("\n"):
        total += max(1, math.ceil(em(part) * size * 1.03 / (w * 72)))
    return total


def P(text=None, size=None, bold=False, color="text", runs=None, **kw):
    """一个段落。runs 是 [(文字, {size, bold, color}), ...]，用于一段里有不同样式的文字。"""
    d = dict(text=text, size=size, bold=bold, color=color, runs=runs)
    d.update(kw)
    return d


def _split(spec):
    m = re.match(r"^([A-Za-z0-9]+?)([+-]\d*\.?\d+)?$", spec)
    return m.group(1), float(m.group(2)) if m.group(2) else 0.0


def _apply_color(cf, spec):
    name, br = _split(spec)
    if name in THEME:
        cf.theme_color = THEME[name]
    else:
        cf.rgb = RGBColor.from_string(name)
    if br:
        cf.brightness = br


def _color_el(parent, spec):
    name, br = _split(spec)
    if name in THEME:
        el = etree.SubElement(parent, qn("a:schemeClr"))
    else:
        el = etree.SubElement(parent, qn("a:srgbClr"))
    el.set("val", name)
    if br > 0:
        etree.SubElement(el, qn("a:lumMod")).set("val", str(int((1 - br) * 100000)))
        etree.SubElement(el, qn("a:lumOff")).set("val", str(int(br * 100000)))
    elif br < 0:
        etree.SubElement(el, qn("a:lumMod")).set("val", str(int((1 + br) * 100000)))


def _nostyle(s):
    st = s._element.find(qn("p:style"))
    if st is not None:
        s._element.remove(st)


# ------------------------------------------------------------------ 风格
class Style:
    def __init__(self, d):
        self.d = d
        self.colors = d["colors"]
        self.fonts = d.get("fonts") or {}
        self.sizes = dict(SIZES)
        self.sizes.update(d.get("sizes") or {})
        self.block = d.get("block", "line")
        self.shadow = bool(d.get("shadow"))
        self.num_font = d.get("num_font")
        self.radius = d.get("radius", 0.06)
        self.roles = d["roles"]

    def col(self, c):
        return self.colors.get(c, c)

    def sz(self, s):
        return self.sizes.get(s, s) if isinstance(s, str) else s


# ------------------------------------------------------------------ 整份
class Deck:
    def __init__(self, template, style=None, fit=True):
        """template：模板或带示例页的 pptx。style：风格文件路径或 dict；不给时找模板同名的 .json。
        fit：图示的字号按主体区的高度自动选档；传 False 则固定用风格文件里的字号。"""
        self.fit = fit
        self.prs = Presentation(template)
        if style is None:
            style = os.path.splitext(template)[0] + ".json"
        if isinstance(style, str):
            with open(style, encoding="utf-8") as f:
                style = json.load(f)
        self.style = Style(style)
        self.W = self.prs.slide_width / EMU
        self.H = self.prs.slide_height / EMU
        self._samples = len(self.prs.slides)
        self._pages = 0
        self._made = []
        self._used = {}
        self._notes = []

    def _layout(self, name):
        for lay in self.prs.slide_layouts:
            if lay.name == name:
                return lay
        raise ValueError("模板里没有版式“%s”。现有版式：%s" % (name, "、".join(l.name for l in self.prs.slide_layouts)))

    def _copy_sample(self, n):
        src = self.prs.slides[n - 1]
        dst = self.prs.slides.add_slide(src.slide_layout)
        tree = dst.shapes._spTree
        for el in list(tree)[2:]:
            tree.remove(el)
        rid = {}
        for rel in src.part.rels.values():
            if rel.reltype.endswith("/slideLayout") or rel.reltype.endswith("/notesSlide"):
                continue
            if rel.is_external:
                rid[rel.rId] = dst.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            else:
                rid[rel.rId] = dst.part.relate_to(rel.target_part, rel.reltype)
        for el in list(src.shapes._spTree)[2:]:
            new = copy.deepcopy(el)
            for node in new.iter():
                for k, v in list(node.attrib.items()):
                    if k.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}") and v in rid:
                        node.set(k, rid[v])
            tree.append(new)
        return dst

    def page(self, role, notes=None, pick=None, **fields):
        """按风格文件里的页面角色新增一页，填入标题等固定位置的文字。返回 Page。
        角色配了几张示例页时，pick 指定用第几张（从 1 起）；不给就按出现的先后轮流用。"""
        st = self.style
        if role not in st.roles:
            raise ValueError("风格文件里没有页面角色“%s”。现有：%s" % (role, "、".join(st.roles)))
        spec = st.roles[role]
        known = spec.get("fields", {})
        for k in fields:
            if k not in known:
                raise ValueError("页面角色“%s”没有“%s”这一栏。现有：%s" % (role, k, "、".join(known)))
        if "sample" in spec:
            n = spec["sample"]
            if isinstance(n, list):          # 几张示例页轮流用，例如每章的章首页背景不同
                used = self._used.get(role, 0)
                self._used[role] = used + 1
                n = n[(pick - 1) % len(n)] if pick else n[used % len(n)]
            slide = self._copy_sample(n)
        else:
            slide = self.prs.slides.add_slide(self._layout(spec["layout"]))
        self._pages += 1
        body = list(spec["body"]) if spec.get("body") else None
        if body and "lead" in known and not fields.get("lead"):
            up = spec.get("lead_h")
            if up is None:                   # 没写就按导语文本框的位置算：主体区上移到导语原来的地方
                up = max(0.0, body[1] - known["lead"]["box"][1]) if "box" in known["lead"] else 0.0
            body[1] -= up
            body[3] += up
        if body and "conclusion" in known and not fields.get("conclusion"):
            more = spec.get("conclusion_h")
            if more is None:
                box = known["conclusion"].get("box")
                more = max(0.0, box[1] + box[3] - body[1] - body[3]) if box else 0.0
            body[3] += more
        pg = Page(self, slide, body)
        self._made.append(pg)
        for name, f in known.items():
            val = fields.get(name)
            if "ph" in f:
                ph = next((s for s in slide.placeholders if s.placeholder_format.idx == f["ph"]), None)
                if ph is None:
                    raise ValueError("版式“%s”里没有编号 %s 的占位符" % (spec.get("layout"), f["ph"]))
                if val:
                    ph.text_frame.text = str(val)
                else:
                    ph._element.getparent().remove(ph._element)
            elif "sample_text" in f or "shape" in f:
                sh = _find_text(slide.shapes, f["sample_text"]) if "sample_text" in f else _find_shape(slide.shapes, f["shape"])
                if sh is None:
                    raise ValueError("示例页 %s 上找不到%s" % (spec.get("sample"), "文字“%s”" % f["sample_text"]
                                                             if "sample_text" in f else "名为“%s”的形状" % f["shape"]))
                if val:
                    old = sh.text_frame.text
                    sz = next((r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.font.size), None)
                    wd = (sh.width or 0) / EMU
                    if sz and wd > 0.5 and lines(str(val), sz, wd) > max(1, lines(old, sz, wd)):
                        print("第 %d 页“%s”：模板这个位置是 %gpt 的字，一行约放 %d 个字；现在的文字会多出一行，可能压到下面的内容。"
                              "改法：缩短文字，或在风格文件里把这一栏改成固定位置的文本框。" % (self._pages, name, sz, int(wd * 72 / sz)))
                    _replace_text(sh, str(val))
                else:
                    sh._element.getparent().remove(sh._element)
            elif "box" in f and val:
                x, y, w, h = f["box"]
                para = [P(str(val), f.get("size", 12), f.get("bold", False), f.get("color", "text"), align=f.get("align", "l"))]
                if f.get("fill"):
                    pg.box(x, y, w, h, f["fill"], r=f.get("radius"), name=name, paras=para, anchor=f.get("anchor", "m"),
                           inset=tuple(f.get("inset", (0.2, 0, 0.2, 0))))
                else:
                    pg.text(x, y, w, h, para, name, f.get("anchor", "t"), tuple(f.get("inset", (0, 0, 0, 0))))
        if spec.get("page_number"):
            f = spec["page_number"]
            x, y, w, h = f["box"]
            pg.text(x, y, w, h, [P(str(self._pages), f.get("size", 9), False, f.get("color", "muted"),
                                   align=f.get("align", "r"))], "页码", f.get("anchor", "m"))
        if notes:
            slide.notes_slide.notes_text_frame.text = notes
        return pg

    def save(self, path):
        """删掉模板自带的示例页后保存。"""
        lst = self.prs.slides._sldIdLst
        for sid in list(lst)[: self._samples]:
            self.prs.part.drop_rel(sid.rId)
            lst.remove(sid)
        self._samples = 0
        self.prs.save(path)
        for note in self._notes:
            print(note)
        for pg in self._made:
            if pg.body and pg.bottom > pg.body[1] + pg.body[3] + 0.03:
                print("第 %d 页：字号已缩到最小一档，内容仍比主体区多出 %.2f 英寸，会压到下面的结论条或页脚。"
                      "改法：拆成两页，或删去次要的行、次要的方面。不要把每一格都压成短语。" % (pg.no, pg.bottom - pg.body[1] - pg.body[3]))
        return path


def _find_text(shapes, text):
    for s in shapes:
        if s.shape_type == 6:
            r = _find_text(s.shapes, text)
            if r is not None:
                return r
        elif s.has_text_frame and s.text_frame.text.strip() == text.strip():
            return s
    return None


def _find_shape(shapes, name):
    for s in shapes:
        if s.name == name and s.has_text_frame:
            return s
        if s.shape_type == 6:
            r = _find_shape(s.shapes, name)
            if r is not None:
                return r
    return None


def _replace_text(sh, text):
    """换文字，保留原来第一段第一个 run 的格式。多行用换行符分开。"""
    tf = sh.text_frame
    paras = tf.paragraphs
    first = paras[0]
    for p in paras[1:]:
        p._p.getparent().remove(p._p)
    runs = first.runs
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    for br in first._p.findall(qn("a:br")):
        first._p.remove(br)
    parts = text.split("\n")
    if runs:
        runs[0].text = parts[0]
    else:
        first.add_run().text = parts[0]
    for part in parts[1:]:
        np = copy.deepcopy(first._p)
        np.find(qn("a:r")).find(qn("a:t")).text = part
        tf._txBody.append(np)


FIT_STEPS = (1.25, 1.15, 1.08, 1.0, 0.92, 0.85, 0.78, 0.72)   # 字号倍数，从大到小试，取放得下的最大一档
FIT_MIN = 0.92                                       # 一般图示最小缩到这一档；架构图、能力地图、流程图可以缩到最后一档
FIT_AIR = 0.2                                        # 放大字号时，图示下面至少留这么多空（英寸）
FIT_KEYS = ("name", "h", "body", "small", "label")   # 跟着调的字号；大数字和编号不调


def _component(fn, grow=None, shrink=None):
    def wrap(self, x, y, w, *a, leave=0.0, **k):
        return self._place(fn, x, y, w, a, k, leave, grow, shrink)
    wrap.__doc__ = fn.__doc__
    wrap.__name__ = fn.__name__
    wrap.__wrapped__ = fn
    return wrap


def _dense(grow=None):
    """格子多、字小也看得清的图示：字号可以缩到最小一档。grow 是最多放大到几倍。"""
    return lambda fn: _component(fn, grow, FIT_STEPS[-1])


# ------------------------------------------------------------------ 一页
class Page:
    def __init__(self, deck, slide, body):
        self.deck, self.slide, self.style = deck, slide, deck.style
        self.body = tuple(body) if body else None   # 主体区 (x, y, w, h)
        self.bottom = 0.0
        self.no = deck._pages
        self.scale = 1.0
        self.f = 1.0            # 当前这一档的倍数，图示用它收紧自己的边距
        self._recs = []

    def _place(self, fn, x, y, w, a, k, leave, grow=None, shrink=None):
        """画一个图示。一页上的图示共用一档字号：从大到小试，取整页放得进主体区的最大一档。
        每加一个图示，前面画过的图示按新的字号重画，上下的间距保持不变。"""
        recs = self._recs
        above = [r for r in recs if r["bottom"] <= y + 0.01 and y - r["bottom"] <= 1.0]
        rec = {"fn": fn, "x": x, "y": y, "w": w, "a": a, "k": k, "leave": leave, "grow": grow or FIT_STEPS[0], "shrink": shrink or FIT_MIN,
               "els": [], "bottom": 0.0,
               "above": above, "gap": y - max(r["bottom"] for r in above) if above else 0.0}
        recs.append(rec)
        limit = self.body[1] + self.body[3] if self.body else None
        steps = FIT_STEPS if (limit is not None and self.deck.fit) else (1.0,)
        base = self.style.sizes
        tree = self.slide.shapes._spTree
        try:
            for s in steps:
                for r in recs:
                    f = self.f = max(r["shrink"], min(s, r["grow"]))
                    self.style.sizes = dict(base, **{key: round(base[key] * f * 2) / 2 for key in FIT_KEYS})
                    for el in r["els"]:
                        tree.remove(el)
                    had = set(tree)
                    yy = max(q["bottom"] for q in r["above"]) + r["gap"] if r["above"] else r["y"]
                    r["bottom"] = r["fn"](self, r["x"], yy, r["w"], *r["a"], **r["k"])
                    r["els"] = [el for el in tree if el not in had]
                air = FIT_AIR if s > 1 else -0.02
                if limit is None or all(r["bottom"] <= limit - r["leave"] - air for r in recs):
                    break
                if s <= min(r["shrink"] for r in recs):
                    break
        finally:
            self.style.sizes = base
            self.f = 1.0
        self.scale = s
        self.bottom = max(r["bottom"] for r in recs)
        return rec["bottom"]

    # ---- 基础形状
    def _font(self, run, size, bold, color, font=None):
        st = self.style
        f = run.font
        f.size = Pt(st.sz(size) or st.sizes["body"])
        f.bold = bool(bold)
        _apply_color(f.color, st.col(color))
        f.name = (st.num_font if font == "num" and st.num_font else st.fonts.get("latin", "+mn-lt"))
        rPr = run._r.get_or_add_rPr()
        for tag in ("a:ea", "a:cs"):
            el = rPr.find(qn(tag))
            if el is None:
                el = etree.SubElement(rPr, qn(tag))
            el.set("typeface", st.fonts.get("ea", "+mn-ea"))

    def _fill_text(self, tf, paras, anchor="t", inset=(0, 0, 0, 0)):
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [Inches(v) for v in inset]
        tf.vertical_anchor = ANCHOR[anchor]
        first = True
        for p in paras:
            para = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            para.alignment = ALIGN[p.get("align", "l")]
            para.line_spacing = p.get("ls", LS)
            if p.get("after") is not None:
                para.space_after = Pt(p["after"])
            if p.get("before") is not None:
                para.space_before = Pt(p["before"])
            for t, o in p.get("runs") or [(p["text"], {})]:
                for k, part in enumerate(str(t).split("\n")):
                    if k:
                        para.add_line_break()
                    r = para.add_run()
                    r.text = part
                    self._font(r, o.get("size", p["size"]), o.get("bold", p["bold"]), o.get("color", p["color"]),
                               o.get("font", p.get("font")))

    def paras_h(self, paras, w):
        """这些段落在宽度 w 里排下来的高度。"""
        st, h = self.style, 0.0
        for p in paras:
            if p.get("runs"):
                text = "".join(str(t) for t, _ in p["runs"])
                size = max(st.sz(o.get("size", p["size"])) or st.sizes["body"] for _, o in p["runs"])
            else:
                text, size = p["text"], st.sz(p["size"]) or st.sizes["body"]
            h += lines(text, size, w) * size * p.get("ls", LS) * LH / 72
            h += ((p.get("after") or 0) + (p.get("before") or 0)) / 72
        return h

    def text(self, x, y, w, h, paras, name=None, anchor="t", inset=(0, 0, 0, 0)):
        s = self.slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        self._fill_text(s.text_frame, paras, anchor, inset)
        if name:
            s.name = name
        return s

    def box(self, x, y, w, h, fill=None, line=None, lw=0.75, r=None, name=None, shape=None,
            paras=None, anchor="m", inset=(0.1, 0.05, 0.1, 0.05), shadow=False):
        st = shape or (MSO_SHAPE.ROUNDED_RECTANGLE if r else MSO_SHAPE.RECTANGLE)
        s = self.slide.shapes.add_shape(st, Inches(x), Inches(y), Inches(w), Inches(h))
        _nostyle(s)
        if fill:
            s.fill.solid()
            _apply_color(s.fill.fore_color, self.style.col(fill))
        else:
            s.fill.background()
        if line:
            _apply_color(s.line.color, self.style.col(line))
            s.line.width = Pt(lw)
        else:
            s.line.fill.background()
        if r and shape is None:
            s.adjustments[0] = min(0.5, r / min(w, h))
        if shadow:
            fx = etree.SubElement(s._element.spPr, qn("a:effectLst"))
            sh = etree.SubElement(fx, qn("a:outerShdw"))
            for k, v in (("blurRad", "228600"), ("dist", "50800"), ("dir", "5400000"), ("algn", "t"), ("rotWithShape", "0")):
                sh.set(k, v)
            c = etree.SubElement(sh, qn("a:srgbClr"))
            c.set("val", "101828")
            etree.SubElement(c, qn("a:alpha")).set("val", "10000")
        if paras:
            self._fill_text(s.text_frame, paras, anchor, inset)
        if name:
            s.name = name
        return s

    def card(self, x, y, w, h, fill=None, name=None):
        """一块承载内容的底。卡片风格是白底加柔和阴影，线条风格是浅底色。"""
        card = self.style.block == "card"
        return self.box(x, y, w, h, fill or ("surface" if card else "tint"), r=self.style.radius, name=name,
                        shadow=card and self.style.shadow and fill is None)

    def badge(self, x, y, d, text, fill="primary", color="on_primary", line=None, name=None):
        """圆形数字徽标。"""
        return self.box(x, y, d, d, fill, line, 1.25, shape=MSO_SHAPE.OVAL, name=name,
                        paras=[P(str(text), round(d * 72 * 0.42, 1), True, color, align="c")], inset=(0, 0, 0, 0))

    def pill(self, x, y, text, fill="tint", color="primary", size="small", h=0.3, name=None):
        """圆角小标签。返回 (形状, 宽度)。"""
        w = em(text) * self.style.sz(size) / 72 + 0.36
        return self.box(x, y, w, h, fill, r=h / 2, name=name, paras=[P(text, size, True, color, align="c")],
                        inset=(0.04, 0, 0.04, 0)), w

    def line(self, x1, y1, x2, y2, color="line", lw=0.75, name=None):
        c = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        _nostyle(c)
        _apply_color(c.line.color, self.style.col(color))
        c.line.width = Pt(lw)
        if name:
            c.name = name
        return c

    def note(self, msg):
        """记一条提醒，保存时打印。同一条只记一次。"""
        msg = "第 %d 页：%s" % (self.no, msg)
        if msg not in self.deck._notes:
            self.deck._notes.append(msg)

    def group(self, shapes, name):
        """把一个图示的全部形状组合并起名。文字和它的底都要放进来。"""
        g = self.slide.shapes.add_group_shape(shapes)
        g.name = name
        return g

    def table(self, x, y, colw, rowh, cells, name="表格", borders=None, fills=None,
              inset=(0.12, 0.05, 0.12, 0.05), anchor="m"):
        """原生表格。cells[r][c] 是段落列表；borders(r,c) 返回 {'b': (颜色, 磅)}；fills(r,c) 返回颜色或 None。"""
        nr, nc = len(cells), len(colw)
        gf = self.slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(sum(colw)), Inches(sum(rowh)))
        gf.name = name
        t = gf.table
        tblPr = t._tbl.tblPr
        for a in ("firstRow", "bandRow", "firstCol", "lastRow", "lastCol", "bandCol"):
            tblPr.set(a, "0")
        sid = tblPr.find(qn("a:tableStyleId"))
        if sid is None:
            sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
        sid.text = NO_STYLE
        for c, w in enumerate(colw):
            t.columns[c].width = Inches(w)
        for r, h in enumerate(rowh):
            t.rows[r].height = Inches(h)
        for r in range(nr):
            for c in range(nc):
                cell = t.cell(r, c)
                a = anchor(r, c) if callable(anchor) else anchor
                self._fill_text(cell.text_frame, cells[r][c], a)
                cell.margin_left, cell.margin_top, cell.margin_right, cell.margin_bottom = [Inches(v) for v in inset]
                cell.vertical_anchor = ANCHOR[a]
                self._border(cell, borders(r, c) if borders else {})
                f = fills(r, c) if fills else None
                if f:
                    cell.fill.solid()
                    _apply_color(cell.fill.fore_color, self.style.col(f))
                else:
                    cell.fill.background()
        return gf

    def _border(self, cell, spec):
        tcPr = cell._tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            el = tcPr.find(qn(tag))
            if el is not None:
                tcPr.remove(el)
        for i, (side, tag) in enumerate((("l", "a:lnL"), ("r", "a:lnR"), ("t", "a:lnT"), ("b", "a:lnB"))):
            ln = etree.Element(qn(tag))
            if side in spec:
                color, pt = spec[side]
                ln.set("w", str(int(pt * 12700)))
                ln.set("cap", "flat")
                ln.set("cmpd", "sng")
                ln.set("algn", "ctr")
                _color_el(etree.SubElement(ln, qn("a:solidFill")), self.style.col(color))
                etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
            else:
                ln.set("w", "0")
                etree.SubElement(ln, qn("a:noFill"))
            tcPr.insert(i, ln)

    def _rows_h(self, cells, colw, inset, min_h):
        out = []
        for row in cells:
            h = max(self.paras_h(c, colw[i] - inset[0] - inset[2]) for i, c in enumerate(row))
            out.append(max(min_h, h + inset[1] + inset[3] + 0.12))
        return out

    # ---- 图示。都从 (x, y) 起、宽 w，高度跟着内容走，返回图示的下边缘 y
    @_component
    def matrix(self, x, y, w, head, rows, colw=None, hi_col=None, name="对照表"):
        """对照表。rows 的每格是文字；第一列可以写成 (编号, 名称)。hi_col 是要用底色突出的那一列。"""
        marks = set()
        if rows and isinstance(rows[0], dict):
            marks = {i + 1 for i, it in enumerate(rows) if it.get("mark")}
            rows = [[it["name"]] + list(it.get("cells") or []) for it in rows]
            if len(head) == len(rows[0]) - 1:
                head = [""] + list(head)
        n = len(head)
        if not colw:
            first = min(0.24 * w, max(1.4, max(em(str(r[0] if not isinstance(r[0], (tuple, list)) else r[0][1])) for r in rows)
                                      * self.style.sizes["body"] / 72 + 0.5))
            colw = [first] + [(w - first) / (n - 1)] * (n - 1) if n > 1 else [w]
        cells = [[[P(h, "label", True, "muted")] for h in head]]
        for row in rows:
            line = []
            for c, v in enumerate(row):
                if isinstance(v, (tuple, list)):
                    line.append([P(runs=[(v[0], {"size": "num", "color": "primary", "font": "num"}), ("　" + v[1], {})],
                                   size="name", bold=True)])
                elif c == 0 or c == hi_col:
                    line.append([P(v, "body", True)])
                else:
                    line.append([P(v, "body")])
            cells.append(line)
        inset = (0.14, 0.06, 0.14, 0.06)
        rowh = self._rows_h(cells[:1], colw, inset, 0.38) + self._rows_h(cells[1:], colw, inset, 0.46)
        self.table(x, y, colw, rowh, cells, name,
                   borders=lambda r, c: {"b": ("primary", 1.5)} if r == 0 else {"b": ("line", 0.75)},
                   fills=lambda r, c: "tint" if ((c == hi_col and r > 0) or r in marks) else None, inset=inset)
        return y + sum(rowh)

    @_component
    def stats(self, x, y, w, items, name="大数字"):
        """一排大数字。items 是 [(数字加单位, 含义), ...]，例如 ("4 轮", "流程全部走通")，最多四个。"""
        n = len(items)
        card = self.style.block == "card"
        gap = 0.2 if card else 0.0
        cw = (w - gap * (n - 1)) / n
        h = 1.05
        parts = []
        for i, (num, label) in enumerate(items):
            cx = x + i * (cw + gap)
            m = re.match(r"^([\d.,+\-%]+)\s*(.*)$", num)
            number, unit = (m.group(1), m.group(2)) if m else (num, "")
            pad = 0.24 if (card or i > 0) else 0.0
            if card:
                parts.append(self.card(cx, y, cw, h, name="%s-%d-底" % (name, i + 1)))
            elif i > 0:
                parts.append(self.line(cx, y + 0.14, cx, y + h - 0.14, "line", 0.75, "%s-%d-分隔" % (name, i + 1)))
            runs = [(number, {"size": "stat", "color": "primary", "font": "num"})]
            if unit:
                runs.append((" " + unit, {"size": "name"}))
            parts.append(self.text(cx + pad, y, cw - pad - 0.1, h,
                                   [P(runs=runs, size="stat", bold=True, ls=0.95), P(label, "body", False, "muted")],
                                   "%s-%d-文字" % (name, i + 1), "m"))
        self.group(parts, name)
        return y + h

    @_component
    def chips(self, x, y, w, label, items, name="标签组"):
        """一组短条目，画成圆角标签，前面一个小栏名。"""
        z = self.style.sizes
        card = self.style.block == "card"
        parts = [self.text(x, y, w, 0.25, [P(label, "label", True, "muted")], name + "-栏名")] if label else []
        px, py, h = x, y + (0.34 if label else 0.0), 0.38
        for i, t in enumerate(items):
            pw = em(t) * z["body"] / 72 + 0.48
            if px + pw > x + w + 0.01:
                px, py = x, py + h + 0.1
            parts.append(self.box(px, py, pw, h, "surface" if card else "tint", r=h / 2, name="%s-第%d项" % (name, i + 1),
                                  paras=[P(t, "body", align="c")], inset=(0.05, 0, 0.05, 0),
                                  shadow=card and self.style.shadow))
            px += pw + 0.12
        self.group(parts, name)
        return py + h

    @_component
    def cards(self, x, y, w, items, cols=None, name="卡片组"):
        """一排卡片。items 的每项是 dict：tag（小标签）、title、sub（一行补充）、body（正文）。"""
        cols = cols or len(items)
        gap = 0.2
        cw = (w - gap * (cols - 1)) / cols
        card = self.style.block == "card"
        pad = 0.22
        tagh = 0.44 if any(it.get("tag") for it in items) else 0.0
        bodies = []
        for it in items:
            ps = [P(it["title"], "name", True, after=3)]
            if it.get("sub"):
                ps.append(P(it["sub"], "small", False, "muted", after=9))
            if it.get("body"):
                ps.append(P(it["body"], "body"))
            bodies.append(ps)
        parts, yy = [], y
        for r in range(0, len(items), cols):
            row = bodies[r:r + cols]
            ch = max(self.paras_h(b, cw - 2 * pad) for b in row) + 2 * pad + tagh + 0.04
            for i, b in enumerate(row):
                cx = x + i * (cw + gap)
                parts.append(self.card(cx, yy, cw, ch, name="%s-第%d项-底" % (name, r + i + 1)))
                if items[r + i].get("tag"):
                    parts.append(self.pill(cx + pad, yy + pad, items[r + i]["tag"], "tint" if card else "surface",
                                           name="%s-第%d项-标签" % (name, r + i + 1))[0])
                parts.append(self.text(cx + pad, yy + pad + tagh, cw - 2 * pad, ch - 2 * pad - tagh, b,
                                       "%s-第%d项-文字" % (name, r + i + 1)))
            yy += ch + gap
        self.group(parts, name)
        return yy - gap

    @_component
    def layers(self, x, y, w, layers, side=None, name="分层"):
        """分层。layers 的每项是 dict：no、name、desc、sub（小字补充，可无）。
        第一项可带 desc_title、sub_title 作栏名。
        side 是 {"title": 栏名, "items": [每层对应的一句话]}，画在右侧并与各层对齐。"""
        if self.style.block != "card":
            return self._layers_table(x, y, w, layers, side, name)
        if any(l.get("sub") for l in layers):
            pad = 0.22
            bottom = self._layers_table(x + pad, y + pad * 0.6, w - 2 * pad, layers, side, name)
            bg = self.card(x, y, w, bottom - y + pad * 0.6, name=name + "-底")
            tree = bg._element.getparent()
            tree.remove(bg._element)
            tree.insert(len(tree) - 1, bg._element)     # 底放到表的下面
            return bottom + pad * 0.6
        sw = 0.29 * w if side else 0
        xr = x + w - sw - (0.25 if side else 0)
        parts, sparts = [], []
        yy = y
        if side or layers[0].get("desc_title"):
            if layers[0].get("desc_title"):
                self.text(x, y, 6, 0.25, [P(layers[0]["desc_title"], "label", True, "muted")], name + "-栏名")
            if side:
                self.text(xr + 0.25, y, sw, 0.25, [P(side["title"], "label", True, "muted")], name + "-侧栏名")
            yy += 0.34
        step = min(0.34, (xr - x) * 0.05)
        hs = []
        for i, l in enumerate(layers):
            ps = [P(l["desc"], "body", after=2)]
            if l.get("sub"):
                ps.append(P(l["sub"], "small", False, "muted"))
            hs.append((ps, self.paras_h(ps, xr - x - step * i - 1.75)))
        ch = max(h for _, h in hs) + 0.24
        gap = 0.08
        total = len(layers) * ch + (len(layers) - 1) * gap
        if side:
            sparts.append(self.box(xr + 0.25, yy, sw, total, "tint", r=self.style.radius, name=name + "-侧栏底"))
        for i, l in enumerate(layers):
            cx, cy = x + step * i, yy + i * (ch + gap)
            parts.append(self.card(cx, cy, xr - cx, ch, name="%s-第%s层-底" % (name, l["no"])))
            parts.append(self.badge(cx + 0.16, cy + ch / 2 - 0.19, 0.38, l["no"], name="%s-第%s层-号" % (name, l["no"])))
            parts.append(self.text(cx + 0.68, cy, 0.85, ch, [P(l["name"], "name", True)], "%s-第%s层-名" % (name, l["no"]), "m"))
            parts.append(self.text(cx + 1.58, cy, xr - cx - 1.75, ch, hs[i][0], "%s-第%s层-说明" % (name, l["no"]), "m"))
            if side:
                sparts.append(self.text(xr + 0.47, cy, sw - 0.4, ch, [P(side["items"][i], "body", True)],
                                        "%s-侧栏-第%s层" % (name, l["no"]), "m"))
        self.group(parts, name)
        if side:
            self.group(sparts, name + "-侧栏")
        return yy + total

    def _layers_table(self, x, y, w, layers, side, name):
        has_sub = any(l.get("sub") for l in layers)
        head = ["", layers[0].get("desc_title", "")]
        if has_sub:
            head.append(layers[0].get("sub_title", ""))
        if side:
            head.append(side["title"])
        rows = []
        for i, l in enumerate(layers):
            row = [(l["no"], l["name"]), l["desc"]]
            if has_sub:
                row.append(l.get("sub", ""))
            if side:
                row.append(side["items"][i])
            rows.append(row)
        sw = 0.19 * w if side else 0
        rest = w - 1.45 - sw
        colw = [1.45] + ([rest * 0.5, rest * 0.5] if has_sub else [rest]) + ([sw] if side else [])
        return self.matrix.__wrapped__(self, x, y, w, head, rows, colw, hi_col=len(colw) - 1 if side else None, name=name)

    @_component
    def stages(self, x, y, w, phases, labels, extra=None, name="分阶段"):
        """分阶段。上面一条进度轨道，下面一张表：每列一个阶段，每行一个方面，行名只写一次。
        phases 的每项是 dict：name、cells（与 labels 一一对应）、tag（如“第一期”，可无）、
        current 或 mark（当前阶段）、status（只在当前阶段显示）、extra（单独一行的内容，可无）。
        extra 参数是那一行的行名。"""
        phases = [dict(p, current=p.get("current") or bool(p.get("mark"))) for p in phases]
        n = len(phases)
        cur = next((i for i, p in enumerate(phases) if p.get("current")), -1)
        card = self.style.block == "card"
        pad = 0.24 if card else 0.0
        tree = self.slide.shapes._spTree
        n0 = len(tree)
        x, y, w = x + pad, y + pad, w - 2 * pad
        lw = 1.3
        cw = (w - lw) / n
        d = 0.32
        cx0 = x + lw + 0.14
        parts = [self.line(cx0, y + d / 2, x + w, y + d / 2, "line", 2.25, name + "-轴线")]
        if cur > 0:
            parts.append(self.line(cx0, y + d / 2, cx0 + cur * cw, y + d / 2, "primary", 2.25, name + "-已完成"))
        for i in range(n):
            done = i <= cur
            parts.append(self.badge(cx0 + i * cw, y, d, i + 1, "primary" if done else "surface",
                                    "on_primary" if done else "muted", None if done else "line", "%s-节点%d" % (name, i + 1)))
        self.group(parts, name + "-进度")
        cells = [[[P("", "small")]]]
        for p in phases:
            runs = [(p["tag"] + "　", {"color": "muted"})] if p.get("tag") else []
            if p.get("current") and p.get("status"):
                runs.append((p["status"], {"color": "primary"}))
            cells[0].append(([P(runs=runs, size="label", bold=True, after=2)] if runs else []) + [P(p["name"], 16, True)])
        for k, label in enumerate(list(labels) + ([extra] if extra else [])):
            row = [[P(label, "label", True, "muted", before=1)]]
            for p in phases:
                v = (p["cells"][k] if k < len(labels) else p.get("extra")) or "无"
                row.append([P(v, "body", False, "text" if v != "无" else "muted")])
            cells.append(row)
        inset = (0.14, 0.09, 0.14, 0.07)
        colw = [lw] + [cw] * n
        rowh = self._rows_h(cells[:1], colw, inset, 0.6) + self._rows_h(cells[1:], colw, inset, 0.5)
        self.table(x, y + 0.46, colw, rowh, cells, name,
                   borders=lambda r, c: {"b": ("primary", 1.5)} if r == 0 else {"b": ("line", 0.75)},
                   fills=lambda r, c: "tint" if c == cur + 1 else None,
                   anchor=lambda r, c: "m" if r == 0 else "t", inset=inset)
        bottom = y + 0.46 + sum(rowh)
        if card:
            bg = self.card(x - pad, y - pad, w + 2 * pad, bottom - y + 2 * pad, name=name + "-底")
            tree.remove(bg._element)
            tree.insert(n0, bg._element)        # 底放到轨道和表的下面
            bottom += pad
        return bottom

    @_component
    def compare(self, x, y, w, cols, hi=None, name="对照栏"):
        """对照栏。cols 的每项是 dict：title、items（文字列表）。hi 是要突出的那一栏的序号（从 0 起）。"""
        n = len(cols)
        gap = 0.3
        cw = (w - gap * (n - 1)) / n
        pad = 0.24
        bodies = []
        for i, c in enumerate(cols):
            on = i == hi
            ps = [P(c["title"], "name", True, "on_primary" if on else "text", after=9)]
            for k, t in enumerate(c["items"]):
                ps.append(P(t, "body", False, "on_primary" if on else "text", after=7 if k < len(c["items"]) - 1 else 0))
            bodies.append(ps)
        ch = max(self.paras_h(b, cw - 2 * pad) for b in bodies) + 2 * pad + 0.05
        parts = []
        for i, b in enumerate(bodies):
            cx = x + i * (cw + gap)
            parts.append(self.card(cx, y, cw, ch, "primary" if i == hi else None, name="%s-第%d栏-底" % (name, i + 1)))
            parts.append(self.text(cx + pad, y + pad, cw - 2 * pad, ch - 2 * pad, b, "%s-第%d栏-文字" % (name, i + 1)))
        self.group(parts, name)
        return y + ch

    @_component
    def flow(self, x, y, w, steps, name="步骤流程"):
        """步骤流程或因果链。steps 的每项是 dict：name、desc（可无）。相邻两步之间画箭头。"""
        n = len(steps)
        gap = 0.36
        cw = (w - gap * (n - 1)) / n
        pad = 0.22
        bodies = []
        for s in steps:
            ps = [P(s["name"], "name", True, after=4)]
            if s.get("desc"):
                ps.append(P(s["desc"], "body"))
            bodies.append(ps)
        ch = max(self.paras_h(b, cw - 2 * pad) for b in bodies) + 2 * pad + 0.5
        parts = []
        for i, b in enumerate(bodies):
            cx = x + i * (cw + gap)
            parts.append(self.card(cx, y, cw, ch, name="%s-第%d步-底" % (name, i + 1)))
            parts.append(self.badge(cx + pad, y + pad, 0.36, i + 1, name="%s-第%d步-号" % (name, i + 1)))
            parts.append(self.text(cx + pad, y + pad + 0.5, cw - 2 * pad, ch - 2 * pad - 0.5, b, "%s-第%d步-文字" % (name, i + 1)))
            if i < n - 1:
                parts.append(self.box(cx + cw + 0.11, y + ch / 2 - 0.14, 0.14, 0.28, "muted", shape=MSO_SHAPE.CHEVRON,
                                      name="%s-箭头%d" % (name, i + 1)))
        self.group(parts, name)
        return y + ch

    @_component
    def callout(self, x, y, w, num, label, name="强调数字"):
        """一个带底色的数字加一句含义。"""
        m = re.match(r"^([\d.,+\-%]+)\s*(.*)$", num)
        number, unit = (m.group(1), m.group(2)) if m else (num, "")
        runs = [(number, {"size": "stat", "color": "accent", "font": "num"})]
        if unit:
            runs.append((" " + unit, {"size": "name"}))
        ps = [P(runs=runs, size="stat", bold=True, ls=0.95), P(label, "body")]
        h = self.paras_h(ps, w - 0.44) + 0.34
        parts = [self.box(x, y, w, h, "tint", r=self.style.radius, name=name + "-底"),
                 self.text(x + 0.22, y + 0.14, w - 0.44, h - 0.28, ps, name + "-文字", "m")]
        self.group(parts, name)
        return y + h

    @_component
    def bullets(self, x, y, w, items, name="条目"):
        """目录页的条目。items 的每项是 (名称, 说明)。每条前面是大号编号，条目之间有细线。"""
        parts, yy = [], y
        for i, it in enumerate(items):
            title, desc = it if isinstance(it, (tuple, list)) else (it, "")
            ps = [P(title, 20, True, after=3)]
            if desc:
                ps.append(P(desc, "body", False, "muted"))
            h = self.paras_h(ps, w - 1.25) + 0.36
            if i > 0:
                parts.append(self.line(x, yy, x + w, yy, "line", 0.75, "%s-线%d" % (name, i)))
            parts.append(self.text(x, yy, 1.1, h, [P("%02d" % (i + 1), 30, True, "primary", font="num")], "%s-%d-号" % (name, i + 1), "m"))
            parts.append(self.text(x + 1.25, yy, w - 1.25, h, ps, "%s-%d-文字" % (name, i + 1), "m"))
            yy += h
        self.group(parts, name)
        return yy

    # ---- 复杂构图：架构图、能力地图、流程图
    def _arrow(self, shape, end="tail"):
        ln = shape.line._get_or_add_ln()
        el = etree.SubElement(ln, qn("a:%sEnd" % end))
        el.set("type", "triangle")
        el.set("w", "med")
        el.set("len", "med")

    def connect(self, a, ia, b, ib, kind="elbow", color="muted", lw=1.0, name=None, first=None):
        """用连接符把形状 a 的连接点 ia 连到形状 b 的连接点 ib，箭头指向 b。连接点：0 上、1 左、2 下、3 右。
        kind：straight 直线；elbow 先横后竖再横；hv 先横后竖（从 a 的左右边出发，进 b 的上下边）；
        vh 先竖后横（从 a 的上下边出发，进 b 的左右边）。
        first：elbow 的第一段横线走多长（英寸），不给就在中间拐。"""
        st = MSO_CONNECTOR.STRAIGHT if kind == "straight" else MSO_CONNECTOR.ELBOW
        c = self.slide.shapes.add_connector(st, 0, 0, 0, 0)
        _nostyle(c)
        if kind == "vh":                # 折一次的连接符从起点先走横向，所以把起点放在 b，箭头画在起点一端
            c.begin_connect(b, ib)
            c.end_connect(a, ia)
        else:
            c.begin_connect(a, ia)
            c.end_connect(b, ib)
        geom = c._element.find(qn("p:spPr")).find(qn("a:prstGeom"))
        if kind in ("vh", "hv"):
            geom.set("prst", "bentConnector2")
        elif kind == "elbow" and first:
            dx = abs(c.end_x - c.begin_x) / float(EMU)
            if dx > first + 0.05:
                av = geom.find(qn("a:avLst"))
                if av is None:
                    av = etree.SubElement(geom, qn("a:avLst"))
                gd = etree.SubElement(av, qn("a:gd"))
                gd.set("name", "adj1")
                gd.set("fmla", "val %d" % int(100000 * first / dx))
        _apply_color(c.line.color, self.style.col(color))
        c.line.width = Pt(lw)
        self._arrow(c, "head" if kind == "vh" else "tail")
        if name:
            c.name = name
        return c

    def polyline(self, pts, color="muted", lw=1.0, arrow=True, name=None):
        """折线，pts 是 [(x, y), ...]。用于回退线这类连接符画不出的走线。"""
        fb = self.slide.shapes.build_freeform(Inches(pts[0][0]), Inches(pts[0][1]))
        fb.add_line_segments([(Inches(px), Inches(py)) for px, py in pts[1:]], close=False)
        s = fb.convert_to_shape()
        _nostyle(s)
        s.fill.background()
        _apply_color(s.line.color, self.style.col(color))
        s.line.width = Pt(lw)
        if arrow:
            self._arrow(s)
        if name:
            s.name = name
        return s

    @staticmethod
    def _item(it):
        if isinstance(it, dict):
            return str(it.get("name", "")), it.get("mark")
        return str(it), None

    def _cell(self, x, y, w, h, it, size, base, legend, name):
        """一个小格子。base 是没有标记时的底色；有标记的按图例的档次加重。"""
        text, mark = self._item(it)
        level = None
        if mark:
            if legend and mark not in legend and mark is not True:
                self.note("格子“%s”的 mark 写的是“%s”，图例里没有这一档（图例：%s），这一格按不加重画了。把两处写成一样。"
                          % (text, mark, "、".join(legend)))
            level = legend.index(mark) if legend and mark in legend else (0 if not legend else None)
            if legend and level == len(legend) - 1 and len(legend) > 1:
                level = None                      # 图例的最后一档就是不加重的样子
        if level == 0:
            fill, line, color, bold = "primary", None, "on_primary", True
        elif level == 1:
            fill, line, color, bold = base, "primary", "primary", True
        else:
            fill, line, color, bold = base, ("line" if base == "surface" else None), "text", False
        return self.box(x, y, w, h, fill, line, 1.0 if level == 1 else 0.75, r=min(self.style.radius, 0.05), name=name,
                        paras=[P(text, size, bold, color, align="c", ls=1.0)], inset=(0.05, 0.02, 0.05, 0.02))

    def _cells_h(self, items, cw, size):
        g = min(1.0, self.f) ** 1.6
        return max(0.34 * g, max(self.paras_h([P(self._item(it)[0], size, ls=1.0)], cw - 0.12) for it in items) + 0.14 * g)

    def _legend(self, x, y, w, legend, base, name):
        """图例，靠右排在图示上方。返回占用的高度。"""
        if not legend:
            return 0.0
        z = self.style.sz("small")
        widths = [0.3 + em(t) * z / 72 + 0.28 for t in legend]
        px = x + w - sum(widths)
        parts = []
        for i, t in enumerate(legend):
            parts.append(self._cell(px, y + 0.03, 0.22, 0.16, {"name": "", "mark": t}, "small", base, legend,
                                    "%s-图例%d-色块" % (name, i + 1)))
            parts.append(self.text(px + 0.3, y, widths[i] - 0.3, 0.22, [P(t, "small", False, "muted")],
                                   "%s-图例%d" % (name, i + 1), "m"))
            px += widths[i]
        self.group(parts, name + "-图例")
        return 0.34

    @_dense()
    def arch(self, x, y, w, layers, left=None, right=None, legend=None, name="架构图"):
        """分层架构图。layers 自上而下，每层是 dict：name（层名）、items（这一层的组成）、note（层名下的一行小字，可无）。
        items 的每项是文字、{"name", "mark"}，或一个分组 {"group": 组名, "items": [...]}，分组在层内画成一个小容器。
        left、right 是贯穿各层的侧柱：{"name": 名称, "items": [可无]}，也可以传一组。
        legend 是图例，例如 ["本期新建", "已有"]；mark 等于图例第一项的格子用主色，最后一项是不加重的样子。"""
        card = self.style.block == "card"
        A, B = ("surface", "tint") if card else ("tint", "surface")       # 层的底色、层内的底色，逐级交替

        def named(d):                       # 大纲里层名和侧柱名写在 group 里
            return dict(d, name=d.get("name") or d.get("group") or "")
        layers = [named(l) for l in layers]
        pillars_l = [named(p) for p in ([left] if isinstance(left, dict) else list(left or []))]
        pillars_r = [named(p) for p in ([right] if isinstance(right, dict) else list(right or []))]
        n_cells = sum(len(g.get("items") or [1]) if isinstance(g, dict) and "group" in g else 1
                      for l in layers for g in l.get("items") or [])
        size = "body" if n_cells <= 14 else "small"
        y += self._legend(x, y, w, legend, B, name)
        g = min(1.0, self.f) ** 1.6
        gap, pad = 0.1 * g, 0.12 * g

        def pw(p):
            return 1.2 if p.get("items") else 0.5
        xl = x + sum(pw(p) + gap for p in pillars_l)
        wl = w - sum(pw(p) + gap for p in pillars_l + pillars_r)
        lw = min(1.5, max(1.0, max(em(l["name"]) for l in layers) * self.style.sz("name") / 72 + 0.36))
        xin, win = xl + lw + pad, wl - lw - 2 * pad
        cg = 0.08
        parts, yy = [], y
        for li, l in enumerate(layers):
            items = l.get("items") or []
            groups = [g for g in items if isinstance(g, dict) and "group" in g]
            plan = []                       # (画法, 参数)
            if groups:
                if len(groups) != len(items):
                    groups = [g if isinstance(g, dict) and "group" in g else {"group": "", "items": [g]} for g in items]
                weights = [max(2, len(g["items"])) for g in groups]
                gw_all = win - cg * (len(groups) - 1)
                gx, inner_h = xin, 0.0
                for g, wt in zip(groups, weights):
                    gw = gw_all * wt / float(sum(weights))
                    n = len(g["items"])
                    minw = min(7, max(em(self._item(it)[0]) for it in g["items"])) * self.style.sz(size) / 72 + 0.2
                    cols = max(1, min(n, int((gw - 0.16 + cg) / (minw + cg))))
                    rows = int(math.ceil(n / float(cols)))
                    cols = int(math.ceil(n / float(rows)))
                    cw = (gw - 0.16 - cg * (cols - 1)) / cols
                    ch = self._cells_h(g["items"], cw, size)
                    th = 0.3 * min(1.0, self.f) if g.get("group") else 0.08
                    gh = th + rows * ch + (rows - 1) * cg + 0.08
                    plan.append((g, gx, gw, cols, cw, ch, th))
                    inner_h = max(inner_h, gh)
                    gx += gw + cg
            else:
                n = len(items)
                minw = min(7, max(em(self._item(it)[0]) for it in items)) * self.style.sz(size) / 72 + 0.2
                per = max(1, min(n, int((win + cg) / (minw + cg))))
                rows = int(math.ceil(n / float(per)))
                per = int(math.ceil(n / float(rows)))
                cw = (win - cg * (per - 1)) / per
                ch = self._cells_h(items, cw, size)
                inner_h = rows * ch + (rows - 1) * cg
            label = [P(l["name"], "name", True, "on_primary", align="c")]
            if l.get("note"):
                label.append(P(l["note"], "small", False, "on_primary", align="c", before=2))
            lh = max(inner_h, self.paras_h(label, lw - 0.2)) + 2 * pad
            parts.append(self.card(xl, yy, wl, lh, name="%s-%s-底" % (name, l["name"])))
            parts.append(self.box(xl, yy, lw, lh, "primary", r=self.style.radius, name="%s-%s-层名" % (name, l["name"]),
                                  paras=label, inset=(0.1, 0.04, 0.1, 0.04)))
            top = yy + pad
            if groups:
                for g, gx, gw, cols, cw, ch, th in plan:
                    parts.append(self.box(gx, top, gw, lh - 2 * pad, B, None if card else None, r=min(self.style.radius, 0.05),
                                          name="%s-%s-%s-底" % (name, l["name"], g.get("group") or "组")))
                    if g.get("group"):
                        parts.append(self.text(gx + 0.08, top + 0.02, gw - 0.16, th - 0.04, [P(g["group"], "label", True, "muted")],
                                               "%s-%s-%s-组名" % (name, l["name"], g["group"]), "m"))
                    for i, it in enumerate(g["items"]):
                        r_, c_ = divmod(i, cols)
                        parts.append(self._cell(gx + 0.08 + c_ * (cw + cg), top + th + r_ * (ch + cg), cw, ch, it, size, A, legend,
                                                "%s-%s-%s" % (name, l["name"], self._item(it)[0])))
            else:
                for i, it in enumerate(items):
                    r_, c_ = divmod(i, per)
                    parts.append(self._cell(xin + c_ * (cw + cg), top + r_ * (ch + cg), cw, ch, it, size, B, legend,
                                            "%s-%s-%s" % (name, l["name"], self._item(it)[0])))
            yy += lh + gap
        bottom = yy - gap
        px = x
        for side, ps in (("左", pillars_l), ("右", pillars_r)):
            if side == "右":
                px = xl + wl + gap
            for p in ps:
                wd = pw(p)
                parts.append(self.card(px, y, wd, bottom - y, name="%s-%s-底" % (name, p["name"])))
                if p.get("items"):
                    parts.append(self.text(px + 0.08, y + 0.1, wd - 0.16, 0.3, [P(p["name"], "body", True, "primary", align="c")],
                                           "%s-%s-名" % (name, p["name"]), "m"))
                    ch = self._cells_h(p["items"], wd - 0.2, size)
                    free = bottom - y - 0.5 - 0.1
                    step = min(ch + 0.3, free / len(p["items"]))
                    for i, it in enumerate(p["items"]):
                        parts.append(self._cell(px + 0.1, y + 0.5 + i * step, wd - 0.2, min(ch, step - 0.04), it, size, B, legend,
                                                "%s-%s-%s" % (name, p["name"], self._item(it)[0])))
                else:
                    parts.append(self.text(px, y, wd, bottom - y, [P("\n".join(p["name"]), "body", True, "primary", align="c")],
                                           "%s-%s-名" % (name, p["name"]), "m"))
                px += wd + gap
        self.group(parts, name)
        return bottom

    @_dense()
    def grid(self, x, y, w, groups, cols=None, legend=None, name="能力地图"):
        """能力地图：分组的密集小格子，各组的格子上下对齐。groups 的每项是 {"group": 组名, "items": [...]}，
        items 的每项是文字或 {"name", "mark"}。cols 是每行几格，不给就按最长的名称算。
        legend 是图例，例如 ["已上线", "建设中", "规划"]；mark 写图例里的一项，第一项用主色，第二项描边，最后一项不加重。"""
        size = "small"
        z = self.style.sz(size)
        y += self._legend(x, y, w, legend, "surface", name)
        lw = min(1.8, max(1.0, max(em(g["group"]) for g in groups) * self.style.sz("body") / 72 + 0.4))
        cg = 0.07
        win = w - lw - 0.12
        every = [it for g in groups for it in g["items"]]
        if not cols:
            minw = min(7, max(em(self._item(it)[0]) for it in every)) * z / 72 + 0.2
            cols = max(3, min(10, int((win + cg) / (minw + cg))))
        cw = (win - cg * (cols - 1)) / cols
        ch = self._cells_h(every, cw, size)
        parts, yy = [], y
        for gi, g in enumerate(groups):
            rows = int(math.ceil(len(g["items"]) / float(cols)))
            gh = rows * ch + (rows - 1) * cg
            if gi:
                parts.append(self.line(x, yy - 0.07 * min(1.0, self.f), x + w, yy - 0.07 * min(1.0, self.f), "line", 0.75,
                                       "%s-分隔%d" % (name, gi)))
            parts.append(self.text(x, yy, lw, gh, [P(g["group"], "body", True, "primary")], "%s-%s-组名" % (name, g["group"]), "m"))
            for i, it in enumerate(g["items"]):
                r_, c_ = divmod(i, cols)
                parts.append(self._cell(x + lw + 0.12 + c_ * (cw + cg), yy + r_ * (ch + cg), cw, ch, it, size, "surface", legend,
                                        "%s-%s-%s" % (name, g["group"], self._item(it)[0])))
            yy += gh + 0.14 * min(1.0, self.f)
        self.group(parts, name)
        return yy - 0.14 * min(1.0, self.f)

    @_dense(1.15)
    def flowchart(self, x, y, w, nodes, lanes=None, name="流程图"):
        """流程图，可带判断、分支、回退和泳道。nodes 按流程先后列出，每项是 dict：
        id、name、kind（开始、步骤、判断、结束，默认步骤）、to（去向，[id 或 (id, 线上的字)]）、
        lane（所在泳道的名称，有 lanes 时必填）、desc（步骤名下的一行小字，可无）、mark（要突出的步骤）、
        col 和 row（第几列、第几行，从 1 起，一般不用写）。
        列和行自动排：每个节点排在它前一个节点的右边，分支往下排；有泳道时，交给另一条泳道的那一步排在正上方或正下方，
        同一条泳道里并行的分支上下排开。to 指向前面的节点时画成回退线。lanes 是泳道名称的列表，自上而下。"""
        KIND = {"开始": "start", "结束": "end", "判断": "decision", "步骤": "step"}
        ns = []
        for n in nodes:
            n = dict(n)
            n["kind"] = KIND.get(n.get("kind"), n.get("kind") or "step")
            to = []
            for t in n.get("to") or []:
                if isinstance(t, dict):
                    to.append((t["id"], t.get("label")))
                elif isinstance(t, (list, tuple)):
                    to.append((t[0], t[1] if len(t) > 1 else None))
                else:
                    to.append((t, None))
            n["to"] = to
            ns.append(n)
        ids = {n["id"]: n for n in ns}
        order = {n["id"]: i for i, n in enumerate(ns)}
        preds = {n["id"]: [] for n in ns}
        for n in ns:
            for t, _ in n["to"]:
                if t not in ids:
                    raise ValueError("流程图里没有 id 为“%s”的节点（出现在“%s”的 to 里）" % (t, n["id"]))
                if order[t] > order[n["id"]]:
                    preds[t].append(n["id"])
            if lanes and n.get("lane") not in lanes:
                raise ValueError("节点“%s”的 lane 要从 %s 里选" % (n["id"], "、".join(lanes)))

        # ---- 排列和行。有泳道时行是“泳道里的第几行”，最后再换算成全图的第几行
        def goes_on(q):
            return any(order[t] > order[q["id"]] for t, _ in q["to"])
        occ = {}
        for n in ns:
            ps = [ids[p] for p in preds[n["id"]]]
            col = int(n["col"]) - 1 if n.get("col") else max([p["c"] + 1 for p in ps] or [0])
            lane = lanes.index(n["lane"]) if lanes else 0
            if lanes and not n.get("col") and len(ps) == 1 and ps[0]["lane"] != n["lane"] and not ps[0].get("handed") \
                    and len([1 for t, _ in ps[0]["to"] if order[t] > order[ps[0]["id"]]]) == 1 \
                    and not any(k[0] == lane and k[1] == ps[0]["c"] for k in occ):
                lo, hi = sorted((lanes.index(ps[0]["lane"]), lane))
                if not any(lo < k[0] < hi and k[1] == ps[0]["c"] for k in occ):
                    col = ps[0]["c"]                  # 交给另一条泳道：排在正上方或正下方
                    ps[0]["handed"] = True
            first = ps[0] if ps else None
            sub = int(n["row"]) - 1 if n.get("row") else (first["sub"] if first and (not lanes or first["lane"] == n["lane"]) else 0)
            if first and not goes_on(n) and any(goes_on(ids[t]) and t != n["id"] and (not lanes or ids[t].get("lane") == n.get("lane"))
                                                for t, _ in first["to"] if order[t] > order[first["id"]]):
                sub += 1                              # 到此结束的分支让出主线的位置
            while (lane, col, sub) in occ:
                sub += 1
            n["c"], n["sub"], n["ln"] = col, sub, lane
            occ[(lane, col, sub)] = n
        nlane = len(lanes) if lanes else 1
        subs = [max([n["sub"] for n in ns if n["ln"] == l] or [0]) + 1 for l in range(nlane)]
        for n in ns:
            n["r"] = sum(subs[:n["ln"]]) + n["sub"]
        ncol = max(n["c"] for n in ns) + 1
        nrow = sum(subs)
        cell = {(n["c"], n["r"]): n for n in ns}

        def free(cols, rows):
            return not any((c_, r_) in cell for c_ in cols for r_ in rows)

        # ---- 先定每条线怎么走，再按线上的字给列间留空
        plan, trunks = [], {}
        for n in ns:
            n["out"], n["in"] = set(), set()
        for n in ns:
            for k, (t, label) in enumerate(n["to"]):
                m = ids[t]
                down = m["r"] > n["r"]
                v = 2 if down else 0
                rows_between = range(min(n["r"], m["r"]) + 1, max(n["r"], m["r"]))
                above = False
                if order[t] <= order[n["id"]]:
                    if m["r"] == n["r"]:
                        above = (n["r"] == 0 and not lanes) or (2 in n["out"] | n["in"] and 0 not in n["out"] | n["in"])
                        kind, so, si = "loop", (0 if above else 2), (0 if above else 2)
                    elif m["c"] == n["c"]:
                        pair = any(e["n"] is m and e["m"] is n and e["kind"] == "vpair" for e in plan)
                        kind, so, si = ("vpair" if pair else "vline"), v, (0 if down else 2)
                    elif 1 not in n["in"] | n["out"] and free(range(m["c"], n["c"]), [n["r"]]) and free([m["c"]], rows_between):
                        kind, so, si = "hv", 1, (0 if down else 2)
                    else:
                        kind, so, si = "detour", (0 if down else 2), (0 if down else 2)
                elif m["c"] == n["c"]:
                    pair = any(t2 == n["id"] for t2, _ in m["to"])
                    kind, so, si = ("vpair" if pair else "vline"), v, (0 if down else 2)
                elif m["r"] == n["r"]:
                    kind, so, si = ("straight" if free(range(n["c"] + 1, m["c"]), [n["r"]]) else "chan"), 3, 1
                elif n["kind"] == "decision" and v not in n["in"] | n["out"] \
                        and free([n["c"]], list(rows_between) + [m["r"]]) and free(range(n["c"] + 1, m["c"]), [m["r"]]):
                    kind, so, si = "vh", v, 1
                else:
                    kind, so, si = "elbow", 3, 1
                n["out"].add(so)
                m["in"].add(si)
                e = {"n": n, "m": m, "label": label, "kind": kind, "above": above}
                if kind == "elbow":                 # 竖的一段放在哪：隔了几列且本行没有东西挡着，就贴着目标；否则贴着出发的节点
                    e["near"] = "target" if m["c"] - n["c"] > 1 and free(range(n["c"] + 1, m["c"]), [n["r"]]) else "source"
                    if e["near"] == "source" and not free(range(n["c"] + 1, m["c"]), [m["r"]]):
                        e["kind"] = "chan"          # 两条路都有节点挡着：走空当
                    elif e["near"] == "source":       # 同一列里的竖线：出发点相同或目标相同的共用一条，其余左右错开
                        groups = trunks.setdefault(n["c"], [])
                        hit = next((g_ for g_ in groups if n["id"] in g_[0] or m["id"] in g_[1]), None)
                        if hit is None:
                            hit = (set(), set())
                            groups.append(hit)
                        hit[0].add(n["id"])
                        hit[1].add(m["id"])
                        e["idx"] = groups.index(hit)
                plan.append(e)
        card = self.style.block == "card"
        tree = self.slide.shapes._spTree
        n0 = len(tree)
        g = min(1.0, self.f) ** 1.6
        pad = 0.2 * g if card else 0.0
        x0, y0, w0 = x, y, w
        x, y, w = x + pad, y + pad, w - 2 * pad
        zs = self.style.sz("small")

        def tw(t):
            return em(t) * zs / 72 + 0.1
        lw = min(1.5, max(0.8, max(em(l) for l in lanes) * self.style.sz("body") / 72 + 0.36)) if lanes else 0.0
        gaps = [0.34] * (ncol - 1)
        for e in plan:
            if e["kind"] == "elbow" and e.get("near") == "source":
                i = e["n"]["c"]
                gaps[i] = max(gaps[i], 0.2 + 0.07 * e["idx"] + (tw(e["label"]) + 0.08 if e["label"] and e["m"]["c"] == i + 1 else 0.1))
                if e["label"] and e["m"]["c"] > i + 1:
                    gaps[e["m"]["c"] - 1] = max(gaps[e["m"]["c"] - 1], tw(e["label"]) + 0.16)
            elif e["label"] and e["kind"] == "straight" and e["m"]["c"] > 0:
                i = e["m"]["c"] - 1
                gaps[i] = max(gaps[i], tw(e["label"]) + 0.16)
        left = x + lw + (0.12 if lanes else 0.0)
        nw = (x + w - 0.04 - left - sum(gaps)) / ncol
        if nw < 0.7:
            self.note("流程图有 %d 列，每个节点只有 %.2f 英寸宽，字会很挤。减少节点、缩短线上的字，或把后半段拆到下一页。" % (ncol, nw))
        cx0 = [left + c_ * nw + sum(gaps[:c_]) for c_ in range(ncol)]
        tight = nw < 1.0
        if tight and any(n.get("desc") for n in ns):
            self.note("流程图节点较多，节点名下面的小字没有画。要保留小字，就减少节点或拆成两页。")
        backs = [e for e in plan if e["kind"] == "loop"]
        over = [e for e in backs if e["above"] and e["n"]["r"] == 0 and not lanes]      # 从第一行上方绕回去的线
        if over:
            y += 0.5 if any(e["label"] for e in over) else 0.3
        for n in ns:
            dec = n["kind"] == "decision"
            ps = [P(n["name"], "small" if dec else "body", True, align="c", ls=1.0)]
            if n.get("desc") and not tight:
                ps.append(P(n["desc"], "small", False, "muted", align="c", ls=1.0, before=2))
            n["w"] = nw
            if dec:
                n["diamond"] = nw >= 1.3                 # 窄的时候菱形里放不下字，改用两头尖的六边形
                if n["diamond"]:
                    n["h"] = max(0.8 * g, self.paras_h(ps, nw * 0.5 - 0.04) * 2 + 0.16)
                else:
                    n["h"] = max(0.46 * g, self.paras_h(ps, nw * 0.62) + 0.16 * g)
            elif n["kind"] in ("start", "end"):
                n["h"] = max(0.4 * g, self.paras_h(ps, nw - 0.24) + 0.14 * g)
            else:
                n["h"] = max(0.46 * g, self.paras_h(ps, nw - 0.14) + 0.16 * g)
            n["ps"] = ps
        rh = [max([n["h"] for n in ns if n["r"] == r] or [0.4]) for r in range(nrow)]
        vgap = max(0.2, 0.5 * g) if (backs or nrow > 1) else 0.3
        lpad = max(0.09, 0.22 * g)                      # 泳道上下的留空
        top, yy, lane_box = [], y, []
        r = 0
        for l in range(nlane):
            ly = yy
            if lanes:
                yy += lpad
            for k in range(subs[l]):
                top.append(yy)
                yy += rh[r] + (vgap if k < subs[l] - 1 else 0)
                r += 1
            if lanes:
                yy += lpad
                lane_box.append((ly, yy - ly))
            elif l < nlane - 1:
                yy += vgap
        under = [e for e in backs if not e["above"] and e["n"]["r"] == nrow - 1 and not lanes]
        bottom = yy + (0.3 if under else 0)
        if lanes:
            lane_parts = []
            for l, name_ in enumerate(lanes):
                by, bh = lane_box[l]
                lane_parts.append(self.box(x, by, lw, bh, "tint", name="%s-泳道-%s" % (name, name_),
                                           paras=[P(name_, "body", True, "primary", align="c")]))
                if l:
                    lane_parts.append(self.line(x, by, x + w, by, "line", 0.75, "%s-泳道线%d" % (name, l)))
            lane_parts.append(self.line(x, y, x + w, y, "line", 0.75, name + "-泳道线-上"))
            lane_parts.append(self.line(x, bottom, x + w, bottom, "line", 0.75, name + "-泳道线-下"))
            self.group(lane_parts, name + "-泳道")
        parts = []
        for n in ns:
            n["x"] = cx0[n["c"]]
            n["y"] = top[n["r"]] + (rh[n["r"]] - n["h"]) / 2
            on = bool(n.get("mark"))
            if n["kind"] == "decision":
                for p in n["ps"]:
                    p["color"] = "on_primary" if on else "primary"
                s = self.box(n["x"], n["y"], n["w"], n["h"], "primary" if on else "surface", "primary", 1.25,
                             shape=MSO_SHAPE.DIAMOND if n["diamond"] else MSO_SHAPE.FLOWCHART_PREPARATION,
                             name="%s-判断-%s" % (name, n["id"]), paras=n["ps"], inset=(0, 0, 0, 0))
            elif n["kind"] in ("start", "end"):
                for p in n["ps"]:
                    p["color"] = "on_primary"
                s = self.box(n["x"], n["y"], n["w"], n["h"], "primary", r=n["h"] / 2,
                             name="%s-%s-%s" % (name, "开始" if n["kind"] == "start" else "结束", n["id"]),
                             paras=n["ps"], inset=(0.1, 0.02, 0.1, 0.02))
            else:
                if on:
                    for p in n["ps"]:
                        p["color"] = "on_primary"
                s = self.box(n["x"], n["y"], n["w"], n["h"], "primary" if on else "tint", r=self.style.radius,
                             name="%s-步骤-%s" % (name, n["id"]), paras=n["ps"], inset=(0.06, 0.03, 0.06, 0.03))
            n["s"] = s
            parts.append(s)

        def tag(tx, ty, t, right=False):
            if t:
                parts.append(self.text(tx - tw(t) if right else tx, ty, tw(t), 0.2,
                                       [P(t, "small", True, "primary", align="r" if right else "l")],
                                       "%s-线上的字-%s" % (name, t), "m"))

        def vch(c_):                                   # 第 c_ 列左边那条竖向空当的 x
            if c_ <= 0:
                return cx0[0] - 0.1
            if c_ >= ncol:
                return cx0[-1] + nw + 0.1
            return cx0[c_] - gaps[c_ - 1] / 2

        def hch(r_):                                   # 第 r_ 行上边那条横向空当的 y
            if r_ <= 0:
                return top[0] - 0.1
            if r_ >= nrow:
                return top[-1] + rh[-1] + 0.1
            return top[r_] - min(0.09, (top[r_] - top[r_ - 1] - rh[r_ - 1]) * 0.4)

        def edge(q, dx, bottom_):                      # 节点上边或下边在偏离中线 dx 处的 y
            if q["kind"] == "decision" and q.get("diamond"):
                d = abs(dx) / (q["w"] / 2) * (q["h"] / 2)
                return q["y"] + q["h"] - d if bottom_ else q["y"] + d
            return q["y"] + q["h"] if bottom_ else q["y"]
        for e in plan:
            n, m, label, kind = e["n"], e["m"], e["label"], e["kind"]
            cx, cy = n["x"] + n["w"] / 2, n["y"] + n["h"] / 2
            mx, my = m["x"] + m["w"] / 2, m["y"] + m["h"] / 2
            down = m["r"] > n["r"]
            nm = "%s-线-%s到%s" % (name, n["id"], m["id"])
            if kind == "loop":
                above = e["above"]
                if above:
                    ch_y = min(n["y"], m["y"]) - 0.22
                    pts = [(cx, n["y"]), (cx, ch_y), (mx, ch_y), (mx, m["y"])]
                    tag(cx - 0.06, ch_y - 0.21, label, True)
                else:
                    ch_y = max(n["y"] + n["h"], m["y"] + m["h"]) + min(0.2, vgap * 0.45, lpad * 0.8 if lanes else 1)
                    pts = [(cx, n["y"] + n["h"]), (cx, ch_y), (mx, ch_y), (mx, m["y"] + m["h"])]
                    tag(cx - 0.06, ch_y + 0.01, label, True)
                parts.append(self.polyline(pts, name=nm))
            elif kind == "vpair":                           # 同一列里一去一回：两条线左右错开
                dx = 0.1 if order[m["id"]] > order[n["id"]] else -0.1
                y1, y2 = edge(n, dx, down), edge(m, dx, not down)
                parts.append(self.polyline([(cx + dx, y1), (cx + dx, y2)], name=nm))
                if dx > 0:
                    tag(cx + dx + 0.05, (y1 + y2) / 2 - 0.1, label)
                else:
                    tag(cx + dx - 0.05, (y1 + y2) / 2 - 0.1, label, True)
            elif kind == "vline":
                parts.append(self.connect(n["s"], 2 if down else 0, m["s"], 0 if down else 2, "straight", name=nm))
                tag(cx + 0.05, (edge(n, 0, down) + edge(m, 0, not down)) / 2 - 0.1, label)
            elif kind == "hv":                              # 回到左边另一行：从左边出发，拐一次进目标
                parts.append(self.connect(n["s"], 1, m["s"], 0 if down else 2, "hv", name=nm))
                tag(n["x"] - 0.04, cy - 0.23, label, True)
            elif kind == "detour":                          # 左边走不通：沿行间的空当绕过去，不穿过别的节点
                up = m["r"] < n["r"]
                d = min(0.08, n["w"] / 6)
                sx, ex = cx - d, mx + d
                hy1 = hch(n["r"]) if up else hch(n["r"] + 1)
                hy2 = hch(m["r"] + 1) if up else hch(m["r"])
                pts = [(sx, edge(n, -d, not up)), (sx, hy1)]
                if abs(hy1 - hy2) > 0.01:
                    vx = vch(m["c"] + 1) if m["c"] < n["c"] else vch(n["c"] + 1)
                    pts += [(vx, hy1), (vx, hy2)]
                pts += [(ex, hy2), (ex, edge(m, d, up))]
                parts.append(self.polyline(pts, name=nm))
                tag(sx - 0.05, (hy1 - 0.2) if up else (hy1 + 0.01), label, True)
            elif kind == "chan":                            # 往右的线被节点挡住：先进列间的空当，再沿行间的空当过去
                hy = hch(m["r"]) if m["r"] > n["r"] else hch(m["r"] + 1) if m["r"] < n["r"] else \
                    (hch(n["r"]) if (n["r"] > 0 or lanes) else hch(n["r"] + 1))
                vx1, vx2 = vch(n["c"] + 1), vch(m["c"])
                parts.append(self.polyline([(n["x"] + n["w"], cy), (vx1, cy), (vx1, hy), (vx2, hy), (vx2, my), (m["x"], my)], name=nm))
                tag(n["x"] + n["w"] + 0.05, cy - 0.22, label)
            elif kind == "straight":
                parts.append(self.connect(n["s"], 3, m["s"], 1, "straight", name=nm))
                tag(m["x"] - 0.07, my - 0.22, label, True)
            elif kind == "vh":                              # 分支：从上下边出发，拐一次进目标的左边
                parts.append(self.connect(n["s"], 2 if down else 0, m["s"], 1, "vh", name=nm))
                tag(cx + 0.05, (n["y"] + n["h"] + 0.02) if down else (n["y"] - 0.22), label)
            elif e["near"] == "target":                     # 换行：先横后竖再横，竖的一段贴着目标
                parts.append(self.connect(n["s"], 3, m["s"], 1, "elbow", name=nm, first=m["x"] - n["x"] - n["w"] - 0.14))
                tag(n["x"] + n["w"] + 0.05, cy - 0.22, label)
            else:                                           # 竖的一段贴着出发的节点
                parts.append(self.connect(n["s"], 3, m["s"], 1, "elbow", name=nm, first=0.1 + 0.07 * e["idx"]))
                tag(m["x"] - 0.07, my - 0.22, label, True)
        self.group(parts, name)
        if card:
            bg = self.card(x0, y0, w0, bottom - y0 + pad, name=name + "-底")
            tree.remove(bg._element)
            tree.insert(n0, bg._element)
            bottom += pad
        return bottom
