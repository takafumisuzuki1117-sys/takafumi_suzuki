"""JLINEメソッド 教材スライドを作るスクリプト。

デザインは JLINE PERSONAL の提案資料（生成り色の背景・黒・赤）に合わせています。
使い方: python3 build_slides.py <出力フォルダ>
章を増やすときは、content_*.py に DECKS を書き足して、ここで読み込みます。
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from content_ch0 import DECKS  # noqa: E402

# ---- デザインの基本 -------------------------------------------------------
BG = RGBColor(0xFA, 0xFA, 0xF8)      # 生成り色の背景
INK = RGBColor(0x14, 0x14, 0x14)     # 黒
RED = RGBColor(0xED, 0x14, 0x2C)     # JLINE の赤
CARD = RGBColor(0xEF, 0xEF, 0xEC)    # カードの地
LINE = RGBColor(0xD2, 0xD2, 0xCD)    # カードの枠
GRAY = RGBColor(0x6E, 0x6E, 0x6E)    # 補足の文字
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

JP = "Yu Gothic"   # 日本語（Windows・Mac の Office どちらにも入っている）
EN = "Oswald"      # 英字ラベル・数字（提案資料と同じ）

W, H = 13.333, 7.5
LEFT = 0.83
RIGHT = W - 0.83
CONTENT_W = RIGHT - LEFT


def _set_fonts(run, latin, east_asian=JP):
    run.font.name = latin
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", east_asian)


def text(slide, x, y, w, h, value, size=14, bold=False, color=INK, font=JP,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.25):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    for i, line in enumerate(value.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        _set_fonts(r, font if font == EN else JP)
    return box


def rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(0.75)
    s.shadow.inherit = False
    return s


def background(slide):
    rect(slide, 0, 0, W, H, BG)


def header(slide, label, page):
    text(slide, LEFT, 0.32, 8, 0.3, f"JLINE METHOD / {label}", 11, True, GRAY, EN)
    text(slide, RIGHT - 1.5, 0.32, 1.5, 0.3, f"{page:03d}", 11, True, GRAY, EN, PP_ALIGN.RIGHT)


def headline(slide, value, sub=None):
    lines = value.count("\n") + 1
    text(slide, LEFT, 0.85, CONTENT_W, 0.6 * lines, value, 26, True)
    y = 0.85 + 0.58 * lines
    if sub:
        text(slide, LEFT, y + 0.05, CONTENT_W, 0.4, sub, 13, False, GRAY)
        y += 0.5
    return y + 0.3  # 本文を置き始める位置


def band(slide, y, value, h=0.72):
    rect(slide, LEFT, y, CONTENT_W, h, INK)
    text(slide, LEFT + 0.35, y, CONTENT_W - 0.7, h, value, 17, True, WHITE,
         anchor=MSO_ANCHOR.MIDDLE)


def card(slide, x, y, w, h, title, body="", num=None, dark=False, title_size=17, body_size=13):
    rect(slide, x, y, w, h, INK if dark else CARD, None if dark else LINE)
    pad = 0.3
    ty = y + 0.28
    if num:
        text(slide, x + pad, ty, w - 2 * pad, 0.25, num, 11, True, RED, EN)
        ty += 0.36
    n_title = title.count("\n") + 1
    title_h = n_title * title_size * 1.3 / 72
    text(slide, x + pad, ty, w - 2 * pad, title_h, title, title_size, True,
         WHITE if dark else INK, spacing=1.2)
    if body:
        by = ty + title_h + 0.16
        text(slide, x + pad, by, w - 2 * pad, max(0.3, y + h - by - 0.15), body, body_size, False,
             RGBColor(0xDD, 0xDD, 0xDB) if dark else GRAY, spacing=1.4)


def sizes(cols):
    if cols >= 5:
        return 15, 12
    if cols == 4:
        return 16, 12.5
    return 18, 14


def grid(slide, top, items, cols, bottom=6.9, gap=0.25, max_h=None):
    rows = (len(items) + cols - 1) // cols
    cw = (CONTENT_W - gap * (cols - 1)) / cols
    ch = (bottom - top - gap * (rows - 1)) / rows
    if max_h:
        ch = min(ch, max_h)
    ts, bs = sizes(cols)
    for i, it in enumerate(items):
        r, c = divmod(i, cols)
        card(slide, LEFT + c * (cw + gap), top + r * (ch + gap), cw, ch,
             it["title"], it.get("body", ""), it.get("num"), dark=it.get("dark", False),
             title_size=ts, body_size=bs)
    return top + rows * ch + (rows - 1) * gap


# ---- スライドの種類 -------------------------------------------------------
def s_cover(slide, d, deck):
    background(slide)
    rect(slide, 0, 0, 0.09, H, RED)
    text(slide, LEFT, 1.35, 10, 0.4, f"CHAPTER {deck['chapter']}  /  {deck['no']}", 14, True, RED, EN)
    text(slide, LEFT, 1.75, 11, 1.0, "JLINE METHOD", 54, True, INK, EN)
    text(slide, LEFT, 3.05, 11.5, 0.8, deck["title"], 30, True)
    rect(slide, LEFT, 3.95, 0.9, 0.05, RED)
    text(slide, LEFT, 4.2, 11, 0.5, deck["subtitle"], 15, False, GRAY)
    text(slide, RIGHT - 6, 6.75, 6, 0.3, "株式会社アスリートバリュー　鈴木崇文", 11, False, GRAY,
         align=PP_ALIGN.RIGHT)


def s_cards(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub"))
    cols = d.get("cols", len(d["items"]))
    has_body = any(it.get("body") for it in d["items"])
    rows = (len(d["items"]) + cols - 1) // cols
    max_h = (2.7 if has_body else 1.9) if rows == 1 else (2.1 if has_body else 1.6)
    bottom = 5.95 if d.get("band") else 6.9
    used = grid(slide, top, d["items"], cols, bottom, max_h=max_h)
    if d.get("band"):
        band(slide, min(6.18, used + 0.4), d["band"])


def s_statement(slide, d, deck):
    rect(slide, LEFT, 2.0, 0.09, 2.2, RED)
    text(slide, LEFT + 0.45, 1.95, CONTENT_W - 0.6, 2.4, d["statement"], 32, True,
         anchor=MSO_ANCHOR.MIDDLE, spacing=1.35)
    if d.get("sub"):
        text(slide, LEFT + 0.45, 4.75, CONTENT_W - 0.6, 1.2, d["sub"], 15, False, GRAY, spacing=1.5)


def s_message(slide, d, deck):
    if d.get("items"):
        top = headline(slide, d["headline"], d.get("sub"))
        cols = d.get("cols", len(d["items"]))
        ts, bs = sizes(cols)
        used = grid(slide, top, d["items"], cols, 5.6, max_h=2.3)
        band(slide, min(6.18, used + 0.4), d["band"])
        return
    lines = d["headline"].count("\n") + 1
    rect(slide, LEFT, 1.55, 0.09, 0.62 * lines + 0.2, RED)
    text(slide, LEFT + 0.45, 1.5, CONTENT_W - 0.6, 0.7 * lines, d["headline"], 32, True, spacing=1.3)
    y = 1.5 + 0.68 * lines + 0.35
    if d.get("sub"):
        text(slide, LEFT + 0.45, y, CONTENT_W - 0.6, 0.5, d["sub"], 15, False, GRAY)
    band(slide, 5.2, d["band"], h=0.95)


def s_flow(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub"))
    items = d["items"]
    per_row = d.get("per_row", len(items))
    rows = (len(items) + per_row - 1) // per_row
    gap = 0.32
    bw = (CONTENT_W - gap * (per_row - 1)) / per_row
    bottom = 5.95 if d.get("band") else 6.9
    bh = min(2.4 if rows == 1 else 1.75, (bottom - top - 0.35 * (rows - 1)) / rows)
    for i, it in enumerate(items):
        r, c = divmod(i, per_row)
        x = LEFT + c * (bw + gap)
        y = top + r * (bh + 0.35)
        dark = it.get("dark", False)
        rect(slide, x, y, bw, bh, INK if dark else CARD, None if dark else LINE)
        text(slide, x + 0.2, y + 0.2, bw - 0.4, 0.25, f"{i + 1:02d}", 11, True, RED, EN)
        text(slide, x + 0.2, y + 0.5, bw - 0.4, 0.8, it["title"], 15 if per_row <= 4 else 14, True,
             WHITE if dark else INK, spacing=1.2)
        if it.get("body"):
            text(slide, x + 0.2, y + bh - 0.62, bw - 0.4, 0.5, it["body"], 12 if per_row <= 4 else 11, False,
                 RGBColor(0xDD, 0xDD, 0xDB) if dark else GRAY, spacing=1.25)
        if c < per_row - 1 and i < len(items) - 1:
            text(slide, x + bw, y + bh / 2 - 0.17, gap, 0.3, "→", 14, True, RED, align=PP_ALIGN.CENTER)
    if d.get("band"):
        last = top + rows * bh + (rows - 1) * 0.35
        band(slide, min(6.18, last + 0.45), d["band"])


def s_numbers(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub"))
    items = d["items"]
    n = len(items)
    gap = 0.3
    arrows = d.get("arrows", False)
    cw = (CONTENT_W - gap * (n - 1)) / n
    ch = 2.5
    for i, it in enumerate(items):
        x = LEFT + i * (cw + gap)
        rect(slide, x, top, cw, ch, CARD, LINE)
        text(slide, x + 0.28, top + 0.28, cw - 0.56, 0.3, it["label"], 13, True, INK)
        big = 48 if len(it["value"]) <= 5 else (38 if len(it["value"]) <= 8 else 30)
        text(slide, x + 0.28, top + 0.72, cw - 0.56, 1.0, it["value"], big, True, RED, EN)
        if it.get("note"):
            text(slide, x + 0.28, top + ch - 0.6, cw - 0.56, 0.45, it["note"], 12, False, GRAY)
        if arrows and i < n - 1:
            text(slide, x + cw, top + ch / 2 - 0.17, gap, 0.3, "→", 14, True, RED, align=PP_ALIGN.CENTER)
    y = top + ch + 0.3
    if d.get("note"):
        text(slide, LEFT, y, CONTENT_W, 0.5, d["note"], 13, False, GRAY, spacing=1.4)
        y += 0.55
    if d.get("band"):
        band(slide, min(6.18, y + 0.15), d["band"])


def s_compare(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub"))
    gap = 0.3
    cw = (CONTENT_W - gap) / 2
    rows = d["rows"]
    rect(slide, LEFT, top, cw, 0.6, CARD, LINE)
    text(slide, LEFT + 0.3, top, cw - 0.6, 0.6, d["left"], 15, True, GRAY, anchor=MSO_ANCHOR.MIDDLE)
    rect(slide, LEFT + cw + gap, top, cw, 0.6, INK)
    text(slide, LEFT + cw + gap + 0.3, top, cw - 0.6, 0.6, d["right"], 15, True, WHITE,
         anchor=MSO_ANCHOR.MIDDLE)
    rh = (6.9 - top - 0.75) / len(rows)
    for i, (l, r) in enumerate(rows):
        y = top + 0.75 + i * rh
        rect(slide, LEFT, y + rh - 0.02, CONTENT_W, 0.015, LINE)
        text(slide, LEFT + 0.3, y, cw - 0.6, rh, l, 14, False, GRAY, anchor=MSO_ANCHOR.MIDDLE)
        text(slide, LEFT + cw + gap + 0.3, y, cw - 0.6, rh, r, 14, True, INK, anchor=MSO_ANCHOR.MIDDLE)
        text(slide, LEFT + cw - 0.1, y, gap + 0.2, rh, "→", 14, True, RED, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)


def s_timeline(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub"))
    items = d["items"]
    rh = (6.9 - top) / len(items)
    rect(slide, LEFT + 1.95, top + 0.12, 0.03, rh * (len(items) - 1) + 0.05, LINE)
    for i, it in enumerate(items):
        y = top + i * rh
        text(slide, LEFT, y, 1.8, 0.35, it["date"], 14, True, INK, EN if it["date"].isascii() else JP)
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(LEFT + 1.87), Inches(y + 0.06),
                                     Inches(0.19), Inches(0.19))
        dot.fill.solid()
        dot.fill.fore_color.rgb = RED if it.get("accent") else INK
        dot.line.fill.background()
        text(slide, LEFT + 2.35, y - 0.02, CONTENT_W - 2.35, 0.35, it["title"], 15, True)
        if it.get("body"):
            text(slide, LEFT + 2.35, y + 0.34, CONTENT_W - 2.35, 0.35, it["body"], 11.5, False, GRAY)


def s_work(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub"))
    text(slide, LEFT, top - 0.05, 2, 0.3, "WORK", 12, True, RED, EN)
    rect(slide, LEFT, top + 0.35, CONTENT_W, 6.9 - top - 0.35, WHITE, LINE)
    y = top + 0.7
    for line in d["lines"]:
        text(slide, LEFT + 0.4, y, CONTENT_W - 0.8, 0.35, line, 14, True)
        y += 0.42
        rect(slide, LEFT + 0.4, y + 0.25, CONTENT_W - 0.8, 0.012, LINE)
        y += d.get("line_gap", 0.55)
    if d.get("hint"):
        text(slide, LEFT + 0.4, 6.35, CONTENT_W - 0.8, 0.35, d["hint"], 11.5, False, GRAY)


def s_summary(slide, d, deck):
    top = headline(slide, d.get("headline", "まとめ"))
    y = top
    for item in d["checks"]:
        rect(slide, LEFT, y + 0.07, 0.22, 0.22, WHITE, INK)
        text(slide, LEFT + 0.4, y, CONTENT_W - 0.4, 0.4, item, 16, True)
        y += 0.62
    band(slide, 6.0, f"次の回　{d['next']}", h=0.9)


RENDER = {
    "cover": s_cover, "cards": s_cards, "statement": s_statement, "message": s_message,
    "flow": s_flow, "numbers": s_numbers, "compare": s_compare, "timeline": s_timeline,
    "work": s_work, "summary": s_summary,
}


def build(deck, out_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    blank = prs.slide_layouts[6]
    for i, d in enumerate(deck["slides"], 1):
        slide = prs.slides.add_slide(blank)
        if d["type"] != "cover":
            background(slide)
            header(slide, f"{deck['no']} {deck['label']}", i)
        RENDER[d["type"]](slide, d, deck)
    path = Path(out_dir) / f"{deck['no']}_{deck['title']}.pptx"
    prs.save(path)
    return path


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    out.mkdir(parents=True, exist_ok=True)
    for deck in DECKS:
        print(build(deck, out))
