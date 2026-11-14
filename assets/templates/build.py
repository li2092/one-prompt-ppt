"""生成 skill 自带的通用模板。每套模板是一个 pptx（主题、母版、五个带占位符的版式）加一个同名的风格文件。

    python build.py            # 生成全部
    python build.py 航线        # 只生成一套
    python art.py              # 重新生成底图（放在 art/ 下）

改模板就改这个文件里的定义再重新生成，不手工改 pptx。
"""
import json
import os
import re
import sys

from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches

HERE = os.path.dirname(os.path.abspath(__file__))
W, H = 13.333, 7.5
NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
EA, LATIN = "微软雅黑", "Arial"


def e(v):
    return str(int(round(v * 914400)))


def clr(c):
    """颜色写法：'tx1'、'accent2'，或带透明度 'bg1@25'（25% 不透明）。"""
    name, _, alpha = c.partition("@")
    inner = '<a:alpha val="%d"/>' % (int(alpha) * 1000) if alpha else ""
    return '<a:solidFill><a:schemeClr val="%s">%s</a:schemeClr></a:solidFill>' % (name, inner)


class Ids:
    def __init__(self):
        self.n = 1

    def next(self):
        self.n += 1
        return self.n


def shape(ids, name, x, y, w, h, fill=None, line=None, lw=1.0, geom="rect", adj=None):
    av = '<a:gd name="adj" fmla="val %d"/>' % adj if adj is not None else ""
    ln = ('<a:ln w="%d">%s</a:ln>' % (int(lw * 12700), clr(line))) if line else "<a:ln><a:noFill/></a:ln>"
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
            '<a:prstGeom prst="%s"><a:avLst>%s</a:avLst></a:prstGeom>%s%s</p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>'
            % (ids.next(), name, e(x), e(y), e(w), e(h), geom, av, clr(fill) if fill else "<a:noFill/>", ln))


def hline(ids, name, x1, y, x2, color, lw=1.0):
    return ('<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="%d" name="%s"/><p:cNvCxnSpPr/><p:nvPr userDrawn="1"/></p:nvCxnSpPr>'
            '<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="0"/></a:xfrm>'
            '<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="%d">%s</a:ln></p:spPr></p:cxnSp>'
            % (ids.next(), name, e(x1), e(y), e(x2 - x1), int(lw * 12700), clr(color)))


def pageno(ids, x, y, w, h, color, size=9, align="r", bold=False):
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="页码"/><p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr>'
            '<a:lstStyle/><a:p><a:pPr algn="%s"/><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">'
            '<a:rPr lang="zh-CN" sz="%d" b="%d">%s<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:rPr><a:t>‹#›</a:t></a:fld>'
            '</a:p></p:txBody></p:sp>' % (ids.next(), e(x), e(y), e(w), e(h), align, size * 100, 1 if bold else 0, clr(color)))


def ph(ids, name, kind, idx, x, y, w, h, size, color, bold=False, anchor="t", align="l", prompt=None,
       fill=None, inset=(0, 0, 0, 0), geom="rect", adj=None, major=False, spc=None, grad=None, font=None):
    """占位符。kind：title、ctrTitle、subTitle、body。idx 只对 body 和 subTitle 有意义。"""
    attr = 'type="%s"' % kind + (' idx="%d"' % idx if idx else "")
    fam = "mj" if major else "mn"
    av = '<a:gd name="adj" fmla="val %d"/>' % adj if adj is not None else ""
    if grad:
        fx = ('<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="%s"/></a:gs>'
              '<a:gs pos="100000"><a:schemeClr val="%s"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>' % grad)
    else:
        fx = clr(fill) if fill else "<a:noFill/>"
    latin = font or ("+%s-lt" % fam)
    sp = ' spc="%d"' % spc if spc else ""
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph %s/></p:nvPr></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
            '<a:prstGeom prst="%s"><a:avLst>%s</a:avLst></a:prstGeom>%s<a:ln><a:noFill/></a:ln></p:spPr>'
            '<p:txBody><a:bodyPr wrap="square" lIns="%s" tIns="%s" rIns="%s" bIns="%s" anchor="%s"><a:noAutofit/></a:bodyPr>'
            '<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="%s"><a:lnSpc><a:spcPct val="110000"/></a:lnSpc>'
            '<a:spcBef><a:spcPts val="0"/></a:spcBef><a:buNone/>'
            '<a:defRPr sz="%d" b="%d" kern="1200"%s>%s<a:latin typeface="%s"/><a:ea typeface="+%s-ea"/></a:defRPr>'
            '</a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="zh-CN"/><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
            % (ids.next(), name, attr, e(x), e(y), e(w), e(h), geom, av, fx,
               e(inset[0]), e(inset[1]), e(inset[2]), e(inset[3]), {"t": "t", "m": "ctr", "b": "b"}[anchor],
               {"l": "l", "c": "ctr", "r": "r"}[align], int(size * 100), 1 if bold else 0, sp, clr(color), latin, fam,
               prompt or name))


def pic(name, file, x, y, w, h):
    return ("pic", name, file, x, y, w, h)


def pic_xml(pid, name, rid, x, y, w, h):
    return ('<p:pic><p:nvPicPr><p:cNvPr id="%d" name="%s"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
            '<p:nvPr userDrawn="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            '<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>' % (pid, name, rid, e(x), e(y), e(w), e(h)))


def layout_xml(name, shapes, bg=None):
    bgx = '<p:bg><p:bgPr>%s<a:effectLst/></p:bgPr></p:bg>' % clr(bg) if bg else ""
    return ('<p:sldLayout %s preserve="1" userDrawn="1"><p:cSld name="%s">%s<p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '%s</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
            % (NS, name, bgx, "".join(shapes)))


def master_csld():
    ids = Ids()
    shapes = [ph(ids, "标题", "title", 0, 0.6, 0.68, 12.13, 0.5, 24, "tx1", True, major=True),
              ph(ids, "正文", "body", 1, 0.6, 1.75, 12.13, 4.5, 14, "tx1")]
    return ('<p:cSld %s><p:bg><p:bgPr>%s<a:effectLst/></p:bgPr></p:bg><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '%s</p:spTree></p:cSld>' % (NS, clr("bg1"), "".join(shapes)))


def tx_styles():
    def lvl(tag, size, fam, bold):
        return ('<a:%s marL="0" indent="0" algn="l"><a:lnSpc><a:spcPct val="110000"/></a:lnSpc><a:buNone/>'
                '<a:defRPr sz="%d" b="%d" kern="1200">%s<a:latin typeface="+%s-lt"/><a:ea typeface="+%s-ea"/></a:defRPr></a:%s>'
                % (tag, size * 100, bold, clr("tx1"), fam, fam, tag))
    body = "".join(lvl("lvl%dpPr" % i, 14, "mn", 0) for i in range(1, 6))
    other = "".join(lvl("lvl%dpPr" % i, 12, "mn", 0) for i in range(1, 6))
    return ('<p:txStyles %s><p:titleStyle>%s</p:titleStyle><p:bodyStyle>%s</p:bodyStyle>'
            '<p:otherStyle>%s</p:otherStyle></p:txStyles>' % (NS, lvl("lvl1pPr", 24, "mj", 1), body, other))


def theme_xml(xml, name, colors):
    order = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
    cs = "".join('<a:%s><a:srgbClr val="%s"/></a:%s>' % (k, colors[k], k) for k in order)
    cs += '<a:hlink><a:srgbClr val="%s"/></a:hlink><a:folHlink><a:srgbClr val="%s"/></a:folHlink>' % (colors["accent1"], colors["accent5"])
    xml = re.sub(r"<a:clrScheme.*?</a:clrScheme>", '<a:clrScheme name="%s">%s</a:clrScheme>' % (name, cs), xml, flags=re.S)
    font = ('<a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface=""/>'
            '<a:font script="Hans" typeface="%s"/>' % (LATIN, EA, EA))
    fs = '<a:fontScheme name="%s"><a:majorFont>%s</a:majorFont><a:minorFont>%s</a:minorFont></a:fontScheme>' % (name, font, font)
    xml = re.sub(r"<a:fontScheme.*?</a:fontScheme>", fs, xml, flags=re.S)
    return re.sub(r'(<a:theme [^>]*name=")[^"]*"', r'\g<1>%s"' % name, xml, count=1)


def replace_children(el, xml):
    new = parse_xml(xml.encode("utf-8"))
    for c in list(el):
        el.remove(c)
    el.attrib.clear()
    for k, v in new.attrib.items():
        el.set(k, v)
    for c in list(new):
        el.append(c)


def write(t):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    master = prs.slide_master
    th = master.part.part_related_by(RT.THEME)
    th._blob = theme_xml(th.blob.decode("utf-8"), t["name"], t["colors"]).encode("utf-8")
    m = master._element
    old = m.find("{http://schemas.openxmlformats.org/presentationml/2006/main}cSld")
    m.replace(old, parse_xml(master_csld().encode("utf-8")))
    old = m.find("{http://schemas.openxmlformats.org/presentationml/2006/main}txStyles")
    m.replace(old, parse_xml(tx_styles().encode("utf-8")))
    lays = list(prs.slide_layouts)
    names = ["封面", "目录", "章首页", "正文", "结尾"]
    for lay in lays[len(names):]:
        prs.slide_layouts.remove(lay)
    for lay, name in zip(lays, names):
        ids = Ids()
        shapes, bg = t["layouts"][name](ids)
        out = []
        for sh in shapes:
            if isinstance(sh, tuple):
                _, rid = lay.part.get_or_add_image_part(os.path.join(HERE, "art", sh[2]))
                out.append(pic_xml(ids.next(), sh[1], rid, *sh[3:]))
            else:
                out.append(sh)
        replace_children(lay._element, layout_xml(name, out, bg))
    path = os.path.join(HERE, t["name"] + ".pptx")
    prs.save(path)
    with open(os.path.join(HERE, t["name"] + ".json"), "w", encoding="utf-8") as f:
        json.dump(t["style"], f, ensure_ascii=False, indent=2)
    return path


def roles(body, agenda_body, conclusion):
    return {
        "cover": {"layout": "封面", "fields": {"title": {"ph": 0}, "subtitle": {"ph": 1}, "kicker": {"ph": 13}, "meta": {"ph": 14}}},
        "agenda": {"layout": "目录", "fields": {"title": {"ph": 0}}, "body": agenda_body},
        "section": {"layout": "章首页", "fields": {"title": {"ph": 0}, "number": {"ph": 13}, "lead": {"ph": 14}}},
        "content": {"layout": "正文", "fields": {"title": {"ph": 0}, "lead": {"ph": 13}, "conclusion": conclusion,
                                                "chapter": {"ph": 15}, "source": {"ph": 16}}, "body": body,
                    "lead_h": 0.45, "conclusion_h": 0.68},
        "closing": {"layout": "结尾", "fields": {"title": {"ph": 0}, "subtitle": {"ph": 1}}},
    }


def dot(ids, x, y, d, color):
    return shape(ids, "标记", x, y, d, d, color, geom="ellipse")


# ------------------------------------------------------------------ 航线：深蓝底上的航线网络图
def routes():
    def cover(ids):
        return [pic("底图", "routes-cover.jpg", 0, 0, W, H),
                dot(ids, 0.8, 1.62, 0.13, "accent2"),
                ph(ids, "眉题", "body", 13, 1.05, 1.53, 8, 0.3, 11, "accent3", True, "m", spc=300),
                ph(ids, "标题", "ctrTitle", 0, 0.8, 1.95, 8.4, 2.35, 46, "bg1", True, "b", major=True),
                ph(ids, "副标题", "subTitle", 1, 0.8, 4.55, 7.6, 0.9, 18, "accent4"),
                hline(ids, "分隔线", 0.8, 6.38, 3.2, "bg1@35", 0.75),
                ph(ids, "单位和日期", "body", 14, 0.8, 6.5, 7, 0.35, 12, "accent4", anchor="m")], "tx1"

    def agenda(ids):
        return [pic("侧图", "routes-side.jpg", 0, 0, 4.6, H),
                dot(ids, 0.7, 0.78, 0.13, "accent2"),
                ph(ids, "标题", "title", 0, 0.7, 1.05, 3.4, 1.4, 34, "bg1", True, major=True),
                pageno(ids, 11.58, 7.0, 1.0, 0.28, "accent5", 11, bold=True)], "bg1"

    def section(ids):
        return [pic("底图", "routes-section.jpg", 0, 0, W, H),
                ph(ids, "章号", "body", 13, 0.72, 0.7, 6, 2.1, 120, "bg1@24", True, "b"),
                dot(ids, 0.8, 3.08, 0.13, "accent2"),
                ph(ids, "标题", "title", 0, 0.8, 3.35, 8.2, 0.9, 40, "bg1", True, major=True),
                ph(ids, "章首预告", "body", 14, 0.8, 4.45, 7.0, 1.5, 16, "accent4")], "tx1"

    def content(ids):
        return [pic("角图", "routes-corner.png", W - 5.26, 0, 5.26, 2.2),
                dot(ids, 0.75, 0.47, 0.12, "accent2"),
                ph(ids, "章节标签", "body", 15, 0.98, 0.4, 8, 0.26, 10.5, "accent1", True, "m", spc=100),
                ph(ids, "标题", "title", 0, 0.75, 0.74, 11.83, 0.52, 24, "tx1", True, major=True),
                ph(ids, "导语", "body", 13, 0.75, 1.3, 11.83, 0.3, 13, "accent5"),
                ph(ids, "出处", "body", 16, 0.75, 7.0, 9.5, 0.28, 9, "accent5", anchor="m"),
                pageno(ids, 11.58, 7.0, 1.0, 0.28, "accent5", 11, bold=True)], "bg1"

    def closing(ids):
        return [pic("底图", "routes-closing.jpg", 0, 0, W, H),
                dot(ids, 0.8, 1.55, 0.13, "accent2"),
                ph(ids, "标题", "title", 0, 0.8, 1.9, 8.6, 2.0, 44, "bg1", True, "b", major=True),
                ph(ids, "副标题", "subTitle", 1, 0.8, 4.1, 7.6, 1.0, 17, "accent4")], "tx1"

    return {
        "name": "航线",
        "colors": dict(dk1="0B1220", lt1="FFFFFF", dk2="0A2A5E", lt2="F1F5FB", accent1="1463FF", accent2="FF6B35",
                       accent3="22D3EE", accent4="A9BCD8", accent5="5B6B85", accent6="DCE3EE"),
        "layouts": {"封面": cover, "目录": agenda, "章首页": section, "正文": content, "结尾": closing},
        "style": {
            "name": "航线", "用途": "向上汇报、决策请示、方案评审。深蓝封面配航线网络图，正文白底，密度可以做到中高",
            "block": "line", "shadow": False,
            "colors": {"primary": "accent1", "accent": "accent2", "text": "tx1", "muted": "accent5", "line": "accent6",
                       "tint": "bg2", "surface": "bg1", "on_primary": "bg1", "dark": "tx2"},
            "roles": roles([0.75, 1.85, 11.83, 4.3], [5.5, 1.3, 7.0, 5.2],
                           {"box": [0.75, 6.26, 11.83, 0.56], "size": 13.5, "bold": True, "color": "on_primary", "fill": "dark",
                            "radius": 0.1, "inset": [0.3, 0, 0.3, 0]}),
        },
    }


# ------------------------------------------------------------------ 极光：靛蓝到珊瑚的渐变光斑
def aurora():
    def cover(ids):
        return [pic("底图", "mesh-cover.jpg", 0, 0, W, H),
                ph(ids, "眉题", "body", 13, 0.8, 1.5, 8, 0.3, 11, "bg1@85", True, "m", spc=300),
                ph(ids, "标题", "ctrTitle", 0, 0.8, 1.95, 9.0, 2.35, 48, "bg1", True, "b", major=True),
                ph(ids, "副标题", "subTitle", 1, 0.8, 4.55, 8.0, 0.9, 18, "bg1@88"),
                ph(ids, "单位和日期", "body", 14, 0.8, 6.5, 7, 0.35, 12, "bg1@80", anchor="m")], "tx2"

    def agenda(ids):
        return [pic("底图", "mesh-light.jpg", 0, 0, W, H),
                ph(ids, "标题", "title", 0, 0.8, 0.9, 8, 0.75, 34, "tx1", True, major=True),
                pageno(ids, 11.53, 7.0, 1.0, 0.28, "accent5", 11, bold=True)], "bg2"

    def section(ids):
        return [pic("底图", "mesh-section.jpg", 0, 0, W, H),
                ph(ids, "章号", "body", 13, 0.72, 0.6, 6, 2.2, 130, "bg1@62", True, "b"),
                ph(ids, "标题", "title", 0, 0.8, 3.3, 9.0, 0.9, 40, "bg1", True, major=True),
                ph(ids, "章首预告", "body", 14, 0.8, 4.4, 7.6, 1.5, 16, "bg1@90")], "tx2"

    def content(ids):
        return [pic("底图", "mesh-light.jpg", 0, 0, W, H),
                ph(ids, "章节标签", "body", 15, 0.8, 0.4, 8, 0.26, 10.5, "accent1", True, "m", spc=100),
                ph(ids, "标题", "title", 0, 0.8, 0.74, 11.73, 0.52, 24, "tx1", True, major=True),
                ph(ids, "导语", "body", 13, 0.8, 1.3, 11.73, 0.3, 13, "accent5"),
                ph(ids, "出处", "body", 16, 0.8, 7.0, 9.5, 0.28, 9, "accent5", anchor="m"),
                pageno(ids, 11.53, 7.0, 1.0, 0.28, "accent5", 11, bold=True)], "bg2"

    def closing(ids):
        return [pic("底图", "mesh-section.jpg", 0, 0, W, H),
                ph(ids, "标题", "title", 0, 0.8, 1.9, 9.5, 2.0, 46, "bg1", True, "b", major=True),
                ph(ids, "副标题", "subTitle", 1, 0.8, 4.1, 8.5, 1.0, 17, "bg1@90")], "tx2"

    return {
        "name": "极光",
        "colors": dict(dk1="1B1B3A", lt1="FFFFFF", dk2="2B2A6B", lt2="F6F7FB", accent1="4F46E5", accent2="7C3AED",
                       accent3="FF7A59", accent4="22D3EE", accent5="6B7089", accent6="E2E4F0"),
        "layouts": {"封面": cover, "目录": agenda, "章首页": section, "正文": content, "结尾": closing},
        "style": {
            "name": "极光", "用途": "客户方案交流、培训、现场讲。渐变光斑底图，白色卡片带柔和阴影",
            "block": "card", "shadow": True,
            "colors": {"primary": "accent1", "accent": "accent3", "text": "tx1", "muted": "accent5", "line": "accent6",
                       "tint": "accent1+0.9", "surface": "bg1", "on_primary": "bg1", "dark": "tx2"},
            "roles": roles([0.8, 1.85, 11.73, 4.3], [0.8, 2.05, 8.6, 4.7],
                           {"box": [0.8, 6.26, 11.73, 0.56], "size": 13.5, "bold": True, "color": "on_primary", "fill": "primary",
                            "radius": 0.28, "inset": [0.35, 0, 0.35, 0]}),
        },
    }


# ------------------------------------------------------------------ 朱砂：米白、墨黑和一个朱红圆
def vermilion():
    def cover(ids):
        return [shape(ids, "红圆", 8.0, -0.9, 6.6, 6.6, "accent1", geom="ellipse"),
                shape(ids, "环", 9.3, 2.3, 5.0, 5.0, None, "tx1", 2.25, "ellipse"),
                shape(ids, "黑点", 12.05, 6.15, 0.5, 0.5, "tx1", geom="ellipse"),
                shape(ids, "标记", 0.8, 0.62, 0.9, 0.09, "tx1"),
                ph(ids, "眉题", "body", 13, 0.8, 0.86, 6, 0.3, 11, "tx1", True, "m", spc=300),
                ph(ids, "标题", "ctrTitle", 0, 0.8, 2.35, 7.6, 2.8, 50, "tx1", True, "b", major=True),
                ph(ids, "副标题", "subTitle", 1, 0.8, 5.4, 6.8, 0.8, 18, "accent5"),
                ph(ids, "单位和日期", "body", 14, 0.8, 6.55, 7, 0.35, 12, "accent5", anchor="m")], "bg2"

    def agenda(ids):
        return [shape(ids, "红圆", 10.5, 4.6, 4.8, 4.8, "accent1", geom="ellipse"),
                shape(ids, "环", 9.5, 3.7, 3.3, 3.3, None, "tx1", 1.75, "ellipse"),
                shape(ids, "标记", 0.8, 0.62, 0.9, 0.09, "tx1"),
                ph(ids, "标题", "title", 0, 0.8, 0.95, 8, 0.75, 34, "tx1", True, major=True),
                pageno(ids, 0.8, 7.0, 1.0, 0.28, "tx1", 11, align="l", bold=True)], "bg2"

    def section(ids):
        return [shape(ids, "环", 9.2, 3.2, 5.8, 5.8, None, "tx1", 2.25, "ellipse"),
                shape(ids, "白点", 11.95, 0.75, 0.5, 0.5, "bg1", geom="ellipse"),
                ph(ids, "章号", "body", 13, 0.62, 0.3, 8, 3.4, 220, "tx1", True, "b"),
                ph(ids, "标题", "title", 0, 0.8, 3.95, 8.2, 0.95, 42, "bg1", True, major=True),
                ph(ids, "章首预告", "body", 14, 0.8, 5.05, 7.8, 1.3, 16, "bg1@92")], "accent1"

    def content(ids):
        return [shape(ids, "标记", 0.8, 0.44, 0.5, 0.07, "tx1"),
                shape(ids, "红点", 12.2, 0.36, 0.33, 0.33, "accent1", geom="ellipse"),
                ph(ids, "章节标签", "body", 15, 1.42, 0.34, 8, 0.26, 10.5, "tx1", True, "m", spc=100),
                ph(ids, "标题", "title", 0, 0.8, 0.74, 11.2, 0.52, 24, "tx1", True, major=True),
                ph(ids, "导语", "body", 13, 0.8, 1.3, 11.73, 0.3, 13, "accent5"),
                ph(ids, "出处", "body", 16, 0.8, 7.0, 9.5, 0.28, 9, "accent5", anchor="m"),
                pageno(ids, 11.53, 7.0, 1.0, 0.28, "tx1", 11, bold=True)], "bg1"

    def closing(ids):
        return [shape(ids, "红圆", 7.6, 0.9, 6.6, 6.6, "accent1", geom="ellipse"),
                shape(ids, "环", 6.4, 3.4, 3.6, 3.6, None, "bg1", 2.0, "ellipse"),
                shape(ids, "标记", 0.8, 1.3, 0.9, 0.09, "bg1"),
                ph(ids, "标题", "title", 0, 0.8, 1.7, 6.4, 2.2, 44, "bg1", True, "b", major=True),
                ph(ids, "副标题", "subTitle", 1, 0.8, 4.1, 5.4, 1.2, 17, "accent4")], "tx1"

    return {
        "name": "朱砂",
        "colors": dict(dk1="141414", lt1="FFFFFF", dk2="141414", lt2="F5F2EC", accent1="D7261E", accent2="141414",
                       accent3="8A8A8A", accent4="BDBDBD", accent5="6E6E6E", accent6="DAD6CE"),
        "layouts": {"封面": cover, "目录": agenda, "章首页": section, "正文": content, "结尾": closing},
        "style": {
            "name": "朱砂", "用途": "分析调研、观点鲜明的汇报。米白、墨黑加一个朱红圆，全部是可编辑的矢量图形",
            "block": "line", "shadow": False,
            "colors": {"primary": "accent1", "accent": "accent1", "text": "tx1", "muted": "accent5", "line": "accent6",
                       "tint": "bg2", "surface": "bg1", "on_primary": "bg1", "dark": "tx1"},
            "roles": roles([0.8, 1.85, 11.73, 4.3], [0.8, 2.1, 8.2, 4.6],
                           {"box": [0.8, 6.26, 11.73, 0.56], "size": 13.5, "bold": True, "color": "on_primary", "fill": "dark",
                            "inset": [0.25, 0, 0.25, 0]}),
        },
    }


TEMPLATES = [routes, aurora, vermilion]

if __name__ == "__main__":
    want = sys.argv[1:]
    for fn in TEMPLATES:
        t = fn()
        if want and t["name"] not in want:
            continue
        print(write(t))
