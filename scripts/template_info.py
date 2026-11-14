"""读一份用户给的模板或旧稿，列出照着它做页面需要知道的事实。

    python template_info.py 模板.pptx

输出：页面尺寸、主题色和主题字体、每个版式的占位符和固定元素、每张示例页上的文字
（位置、字号、颜色、字体），以及全稿用得最多的颜色、字体、字号。
写风格文件时以这里的数字为准，不凭印象。
"""
import re
import sys
from collections import Counter

from pptx import Presentation
from pptx.oxml.ns import qn

E = 914400.0
DEFAULT_ACCENT1 = ("4472C4", "5B9BD5", "4F81BD", "156082")


def inch(v):
    return round((v or 0) / E, 2)


def rect(s):
    return "(%.2f, %.2f, %.2f, %.2f)" % (inch(s.left), inch(s.top), inch(s.width), inch(s.height))


def fill_of(el):
    sp = el.find(qn("p:spPr"))
    if sp is None:
        return ""
    f = sp.find(qn("a:solidFill"))
    if f is None:
        return "图片填充" if sp.find(qn("a:blipFill")) is not None else ("渐变" if sp.find(qn("a:gradFill")) is not None else "")
    c = f[0]
    val = c.get("val", "")
    alpha = c.find(qn("a:alpha"))
    return ("主题色 " + val if c.tag == qn("a:schemeClr") else val) + ("，不透明度 %d%%" % (int(alpha.get("val")) // 1000) if alpha is not None else "")


def runs_of(el):
    out = []
    for r in el.iter(qn("a:r")):
        t = r.find(qn("a:t"))
        if t is None or not (t.text or "").strip():
            continue
        rPr = r.find(qn("a:rPr"))
        size = color = font = None
        bold = False
        if rPr is not None:
            size = int(rPr.get("sz")) / 100 if rPr.get("sz") else None
            bold = rPr.get("b") == "1"
            sf = rPr.find(qn("a:solidFill"))
            if sf is not None and len(sf):
                color = ("主题色 " if sf[0].tag == qn("a:schemeClr") else "") + sf[0].get("val", "")
            ea = rPr.find(qn("a:ea"))
            la = rPr.find(qn("a:latin"))
            font = (ea.get("typeface") if ea is not None else None) or (la.get("typeface") if la is not None else None)
        out.append((t.text, size, bold, color, font))
    return out


def kind(s):
    return str(s.shape_type).split(" ")[0].replace("None", "形状") if s.shape_type is not None else "形状"


def describe_text(s, stats):
    rs = runs_of(s._element)
    if not rs:
        return None
    text = "".join(r[0] for r in rs).replace("\n", " ")
    size = next((r[1] for r in rs if r[1]), None)
    color = next((r[3] for r in rs if r[3]), None)
    font = next((r[4] for r in rs if r[4]), None)
    bold = any(r[2] for r in rs)
    for r in rs:
        n = len(r[0])
        if r[1]:
            stats["size"][r[1]] += n
        if r[3]:
            stats["color"][r[3]] += n
        if r[4]:
            stats["font"][r[4]] += n
    attrs = [a for a in ("%spt" % size if size else "字号继承", "加粗" if bold else "", color or "颜色继承", font or "字体继承") if a]
    return "“%s”　%s" % (text[:28] + ("…" if len(text) > 28 else ""), "，".join(attrs))


def walk(shapes, depth, stats, lines):
    for s in shapes:
        pad = "    " + "  " * depth
        if s.shape_type == 6:
            lines.append("%s组合 %s %s" % (pad, s.name, rect(s)))
            walk(s.shapes, depth + 1, stats, lines)
            continue
        f = fill_of(s._element)
        if f and not f.startswith("图片") and not f.startswith("渐变"):
            stats["fill"][f.split("，")[0]] += 1
        ph = ""
        if s.is_placeholder:
            pf = s.placeholder_format
            ph = "占位符 编号%d 类型%s　" % (pf.idx, str(pf.type).split(" ")[0])
        t = describe_text(s, stats) if s.has_text_frame else None
        lines.append("%s%s%s　%s　位置%s%s%s" % (pad, ph, kind(s), s.name, rect(s), "　填充 " + f if f else "", "　" + t if t else ""))


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    prs = Presentation(sys.argv[1])
    Wd, Ht = inch(prs.slide_width), inch(prs.slide_height)
    print("== 页面\n宽 %.2f 高 %.2f 英寸；示例页 %d 张；母版 %d 个" % (Wd, Ht, len(prs.slides), len(prs.slide_masters)))

    theme = None
    for rel in prs.slide_master.part.rels.values():
        if rel.reltype.endswith("/theme"):
            theme = rel.target_part.blob.decode("utf-8", "ignore")
    print("\n== 主题")
    if theme:
        cols = re.findall(r"<a:(dk1|lt1|dk2|lt2|accent\d)>\s*<a:(?:srgbClr|sysClr)[^>]*?(?:val|lastClr)=\"(\w+)\"", theme)
        print("主题色：" + "　".join("%s=%s" % c for c in cols))
        mj = re.search(r"<a:majorFont>(.*?)</a:majorFont>", theme, re.S)
        mn = re.search(r"<a:minorFont>(.*?)</a:minorFont>", theme, re.S)
        for label, m in (("标题字体", mj), ("正文字体", mn)):
            if m:
                la = re.search(r'<a:latin typeface="([^"]*)"', m.group(1))
                ea = re.search(r'<a:ea typeface="([^"]*)"', m.group(1))
                hans = re.search(r'script="Hans" typeface="([^"]*)"', m.group(1))
                print("%s：西文 %s，中文 %s" % (label, la.group(1) if la else "", (ea.group(1) if ea and ea.group(1) else (hans.group(1) if hans else ""))))
        if dict(cols).get("accent1") in DEFAULT_ACCENT1:
            print("注意：主题色是 Office 的默认值，这份模板没有定制主题。品牌色要从下面“用得最多的颜色”里取，写成色值。")

    stats = {"fill": Counter(), "color": Counter(), "font": Counter(), "size": Counter()}
    used = Counter(s.slide_layout.name for s in prs.slides)
    print("\n== 版式（固定元素会出现在用这个版式的每一页上）")
    for m in prs.slide_masters:
        if len(m.shapes):
            print("  母版上的元素：")
            lines = []
            walk(m.shapes, 0, stats, lines)
            print("\n".join(lines))
        for lay in m.slide_layouts:
            n_ph = sum(1 for s in lay.shapes if s.is_placeholder)
            print("  版式“%s”　占位符 %d 个　被 %d 张示例页使用" % (lay.name, n_ph, used.get(lay.name, 0)))
            lines = []
            walk(lay.shapes, 0, stats, lines)
            print("\n".join(lines))

    print("\n== 示例页（没有占位符的模板，文字就放在示例页上；要复用就复制这一页再换字）")
    for i, s in enumerate(prs.slides, 1):
        print("  第 %d 页　版式“%s”" % (i, s.slide_layout.name))
        lines = []
        walk(s.shapes, 0, stats, lines)
        print("\n".join(lines) if lines else "    （页面上没有自己的元素，内容全部来自版式）")

    print("\n== 用得最多的")
    print("形状填充色：" + "　".join("%s×%d" % kv for kv in stats["fill"].most_common(8)))
    print("文字颜色（按字数）：" + "　".join("%s×%d" % kv for kv in stats["color"].most_common(6)))
    print("字体（按字数）：" + "　".join("%s×%d" % kv for kv in stats["font"].most_common(5)))
    print("字号（按字数）：" + "　".join("%spt×%d" % kv for kv in stats["size"].most_common(8)))
    if not any(s.is_placeholder for lay in prs.slide_layouts for s in lay.shapes):
        print("\n这份模板的版式里没有占位符。封面、章首页这类页按“示例页”复用；正文页用版式加固定位置的文本框，"
              "位置写进风格文件。做法见 references/render.md 第三节。")


if __name__ == "__main__":
    main()
