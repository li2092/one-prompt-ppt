"""把 pptx 渲染成逐页图片和一张总览图，用来检查版面。

    python preview.py deck.pptx [输出目录] [--font 字体名] [--dpi 80]

需要本机有 LibreOffice（soffice）和 pdftoppm。装了 Pillow 时另出一张“总览.png”。
本机没有模板所用的中文字体时，预览会退回宋体，看到的字宽和实际不一样。--font 指定一个
本机有的中文字体，只用于预览，不改原文件。macOS 上默认用 Heiti SC。
"""
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile


def find_soffice():
    for c in ("soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice", "/usr/bin/soffice"):
        if shutil.which(c) or os.path.exists(c):
            return c
    sys.exit("找不到 LibreOffice。请安装后重试，或用其他办法把 pptx 转成图片。")


def with_font(src, dst, font):
    """复制一份，把主题和页面里的中文字体换成 font。只影响预览。"""
    zin = zipfile.ZipFile(src)
    names = set()
    for n in zin.namelist():
        if re.match(r"ppt/theme/theme\d+\.xml", n):
            x = zin.read(n).decode("utf-8")
            names.update(re.findall(r'<a:ea typeface="([^"]+)"', x))
            names.update(re.findall(r'script="Hans" typeface="([^"]+)"', x))
    names.discard(font)
    latin_ok = ("Arial", "Calibri", "Helvetica", "Times New Roman", "Georgia", "Courier New", "Cambria", "Verdana", "Tahoma")

    def swap(m):
        tag, face = m.group(1), m.group(2)
        if face.startswith("+") or face == "":
            return m.group(0)
        if tag == "latin" and (face.startswith(latin_ok) or all(ord(c) < 128 for c in face) and face not in names
                               and "Noto" not in face and "Source Han" not in face):
            return m.group(0)
        return '<a:%s typeface="%s"' % (tag, font)

    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.endswith(".xml"):
            x = data.decode("utf-8")
            for nm in names:
                x = x.replace('typeface="%s"' % nm, 'typeface="%s"' % font)
            x = re.sub(r'<a:(ea|latin) typeface="([^"]*)"', swap, x)
            data = x.encode("utf-8")
        zout.writestr(item, data)
    zout.close()


def main():
    args = [a for a in sys.argv[1:]]
    font = "Heiti SC" if sys.platform == "darwin" else None
    dpi = 80
    for flag in ("--font", "--dpi"):
        if flag in args:
            i = args.index(flag)
            val = args[i + 1]
            del args[i:i + 2]
            if flag == "--font":
                font = None if val == "none" else val
            else:
                dpi = int(val)
    if not args:
        sys.exit(__doc__)
    src = os.path.abspath(args[0])
    out = os.path.abspath(args[1]) if len(args) > 1 else os.path.join(os.path.dirname(src), "preview")
    os.makedirs(out, exist_ok=True)
    stem = os.path.splitext(os.path.basename(src))[0]
    tmp = tempfile.mkdtemp()
    work = os.path.join(tmp, "deck.pptx")
    if font:
        with_font(src, work, font)
    else:
        shutil.copy(src, work)
    subprocess.run([find_soffice(), "--headless", "--convert-to", "pdf", "--outdir", tmp, work],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
    for old in glob.glob(os.path.join(out, stem + "-*.png")):
        os.remove(old)
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", os.path.join(tmp, "deck.pdf"), os.path.join(out, stem)], check=True)
    pages = sorted(glob.glob(os.path.join(out, stem + "-*.png")))
    print("共 %d 页，图片在 %s" % (len(pages), out))
    try:
        from PIL import Image
    except ImportError:
        return
    ims = [Image.open(p) for p in pages]
    w, h = ims[0].size
    cols = 3 if len(ims) > 4 else 2
    rows = (len(ims) + cols - 1) // cols
    scale = min(1.0, 2400.0 / (w * cols))
    tw, th, gap = int(w * scale), int(h * scale), 12
    sheet = Image.new("RGB", (tw * cols + gap * (cols - 1), th * rows + gap * (rows - 1)), (170, 170, 170))
    for i, im in enumerate(ims):
        sheet.paste(im.resize((tw, th)), ((i % cols) * (tw + gap), (i // cols) * (th + gap)))
    path = os.path.join(out, stem + "-总览.png")
    sheet.save(path)
    print("总览图：%s" % path)
    shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
