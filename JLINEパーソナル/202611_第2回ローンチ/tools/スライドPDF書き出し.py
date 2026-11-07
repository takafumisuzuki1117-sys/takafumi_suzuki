"""セミナースライド（pptx）を、参加者に渡すPDFに書き出す。

使い方:
    python3 スライドPDF書き出し.py 入力.pptx 出力.pdf

事前に入れておくもの（このスライドで使っているフォント）:
    - Noto Sans JP（Google Fonts の NotoSansJP[wght].ttf から Regular / Bold を作って ~/.fonts へ）
    - Oswald（同じく Oswald[wght].ttf から Regular / Bold）
    - fc-cache -f

LibreOffice でそのまま書き出すと、PowerPoint と比べて
  1. 日本語と英数字の間にすき間が入り（「サッカー IQ 爆上げ」）、タイトルなどの改行がずれる
  2. フォント指定のない文字（テーマの「ＭＳ Ｐゴシック」など）が中国語用フォントで細く出る
ため、書き出す前に次の2つを直す（元の pptx は変えない）。
  1. 段落の「アジア文字と英数字の間隔」を切り、句読点のぶら下げを入れる
  2. アジア文字のフォントが LibreOffice の既定（Noto Sans CJK SC）か Arial になっている所を Noto Sans JP にする
"""
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

FIX_ASIAN_FONTS = {"Noto Sans CJK SC", "Arial"}


def pv(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def connect():
    proc = subprocess.Popen([
        "soffice", "--headless", "--invisible", "--norestore",
        "--accept=socket,host=localhost,port=2002;urp;",
    ])
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(60):
        try:
            ctx = resolver.resolve(
                "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            return proc, ctx
        except Exception:
            time.sleep(1)
    raise RuntimeError("LibreOffice に接続できませんでした")


def fix_shapes(shapes, stats):
    for i in range(shapes.getCount()):
        shape = shapes.getByIndex(i)
        if shape.supportsService("com.sun.star.drawing.GroupShape"):
            fix_shapes(shape, stats)
            continue
        try:
            text = shape.getText()
        except Exception:
            continue
        if text is None:
            continue
        paragraphs = text.createEnumeration()
        while paragraphs.hasMoreElements():
            para = paragraphs.nextElement()
            try:
                para.setPropertyValue("ParaIsCharacterDistance", False)
                para.setPropertyValue("ParaIsHangingPunctuation", True)
                stats["paragraphs"] += 1
            except Exception:
                pass
            portions = para.createEnumeration()
            while portions.hasMoreElements():
                portion = portions.nextElement()
                try:
                    if portion.getPropertyValue("CharFontNameAsian") in FIX_ASIAN_FONTS:
                        portion.setPropertyValue("CharFontNameAsian", "Noto Sans JP")
                        stats["fonts"] += 1
                except Exception:
                    pass


def main(src, dst):
    proc, ctx = connect()
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(src), "_blank", 0, (pv("Hidden", True),))
    stats = {"paragraphs": 0, "fonts": 0}
    pages = doc.getDrawPages()
    for p in range(pages.getCount()):
        fix_shapes(pages.getByIndex(p), stats)
    print(f"段落 {stats['paragraphs']} 件・フォント {stats['fonts']} か所を直しました")
    doc.storeToURL(uno.systemPathToFileUrl(dst), (pv("FilterName", "impress_pdf_Export"),))
    doc.close(True)
    try:
        desktop.terminate()
    except Exception:
        pass
    proc.wait(timeout=30)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
