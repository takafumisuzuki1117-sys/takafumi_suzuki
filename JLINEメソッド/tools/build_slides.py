"""JLINEメソッド 教材スライドを作るスクリプト。

デザインは JLINE PERSONAL の提案資料（生成り色の背景・黒・赤）に合わせています。
使い方: python3 build_slides.py <出力フォルダ> [章番号 ...]
章を増やすときは、content_chN.py（スライド）と notes_chN.py（台本）を作り、CHAPTERS に足します。
"""
import math
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
import importlib  # noqa: E402


def _load(ch):
    """content_chN（スライド）と notes_chN（台本。まだなければ空）を読み込む。"""
    content = importlib.import_module(f"content_ch{ch}")
    try:
        notes = importlib.import_module(f"notes_ch{ch}").NOTES
    except ImportError:
        notes = {}
    return content.DECKS, notes


# 章ごとの中身と台本
CHAPTERS = {ch: _load(ch) for ch in ("0", "1", "2")}

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
    # 日本語として扱う指定。これがないと、フォントがない環境で中国語の字形に置き換わることがある
    rpr.set("lang", "ja-JP")
    rpr.set("altLang", "en-US")
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


# ---- 文字の大きさの見積もり ----------------------------------------------
def _char_w(ch, size):
    """1文字の幅（インチ）。全角は1em、半角は0.55em で見積もる。"""
    return size / 72 * (0.58 if ord(ch) < 0x2E80 else 1.0)


def n_lines(value, size, width):
    n = 0
    for line in value.split("\n"):
        w = sum(_char_w(c, size) for c in line)
        n += max(1, math.ceil(w / (width * 0.94) - 0.02))
    return n


def text_h(value, size, width, spacing=1.25):
    """折り返しを含めた文字の高さ（インチ）。"""
    return n_lines(value, size, width) * size * spacing * 1.2 / 72


def fit_size(value, width, start, minimum):
    """幅に収まる最大の文字サイズ（1行で収めたいとき）。太字の数字は幅を広めにみる。"""
    size = start
    while size > minimum:
        w = sum(size / 72 * (0.7 if ord(c) < 0x2E80 else 1.0) for c in value)
        if w <= width * 0.92:
            break
        size -= 1
    return size


# ---- 共通の部品 -----------------------------------------------------------
def header(slide, label, page):
    text(slide, LEFT, 0.34, 8, 0.3, f"JLINE METHOD / {label}", 12, True, GRAY, EN)
    text(slide, RIGHT - 1.5, 0.34, 1.5, 0.3, f"{page:03d}", 12, True, GRAY, EN, PP_ALIGN.RIGHT)


HEAD = 30


def headline(slide, value, sub=None):
    lines = value.count("\n") + 1
    pitch = HEAD * 1.2 * 1.2 / 72
    text(slide, LEFT, 0.8, CONTENT_W, pitch * lines, value, HEAD, True, spacing=1.2)
    y = 0.8 + pitch * lines
    if sub:
        text(slide, LEFT, y + 0.12, CONTENT_W, 0.4, sub, 16, False, GRAY)
        y += 0.5
    return y + 0.38  # 本文を置き始める位置


def band(slide, y, value, h=0.86, size=20):
    rect(slide, LEFT, y, CONTENT_W, h, INK)
    text(slide, LEFT + 0.4, y, CONTENT_W - 0.8, h, value, size, True, WHITE,
         anchor=MSO_ANCHOR.MIDDLE)
    return y + h


def place_band(slide, used, value):
    """本文のすぐ下に帯を置く。下にはみ出す場合は下端にそろえる。"""
    band(slide, min(H - 0.55 - 0.86, used + 0.42), value)


PAD = 0.34


def card_need(w, title, body, num, ts, bs):
    inner = w - 2 * PAD
    h = PAD * 2
    if num:
        h += 0.42
    h += text_h(title, ts, inner, 1.2)
    if body:
        h += 0.16 + text_h(body, bs, inner, 1.4)
    return h


def card(slide, x, y, w, h, title, body="", num=None, dark=False, ts=20, bs=16):
    rect(slide, x, y, w, h, INK if dark else CARD, None if dark else LINE)
    inner = w - 2 * PAD
    ty = y + PAD
    if num:
        text(slide, x + PAD, ty, inner, 0.3, num, 13, True, RED, EN)
        ty += 0.42
    th = text_h(title, ts, inner, 1.2)
    text(slide, x + PAD, ty, inner, th, title, ts, True, WHITE if dark else INK, spacing=1.2)
    if body:
        by = ty + th + 0.16
        text(slide, x + PAD, by, inner, max(0.3, y + h - by - 0.1), body, bs, False,
             RGBColor(0xDD, 0xDD, 0xDB) if dark else GRAY, spacing=1.4)


def sizes(cols):
    if cols >= 5:
        return 18, 14
    if cols == 4:
        return 19, 15
    if cols == 3:
        return 21, 16
    return 22, 17


def grid(slide, top, items, cols, bottom=6.95, gap=0.26, min_h=1.2):
    """カードを並べる。高さは中身に合わせ、全カードでそろえる。使った下端を返す。"""
    rows = (len(items) + cols - 1) // cols
    cw = (CONTENT_W - gap * (cols - 1)) / cols
    room = (bottom - top - gap * (rows - 1)) / rows
    ts, bs = sizes(cols)
    while True:
        need = max(card_need(cw, it["title"], it.get("body", ""), it.get("num"), ts, bs)
                   for it in items)
        if need <= room or bs <= 12:
            break
        ts, bs = ts - 1, bs - 1
    ch = min(room, max(need, min_h))
    for i, it in enumerate(items):
        r, c = divmod(i, cols)
        card(slide, LEFT + c * (cw + gap), top + r * (ch + gap), cw, ch,
             it["title"], it.get("body", ""), it.get("num"), dark=it.get("dark", False),
             ts=ts, bs=bs)
    return top + rows * ch + (rows - 1) * gap


def rows_list(slide, top, items, bottom=6.95, gap=0.16, max_h=1.3):
    """横長の行を縦に並べる。左に番号、中央に見出し、右に説明。"""
    n = len(items)
    ts, bs = (22, 18) if n <= 4 else (21, 17)
    items = [dict(it) for it in items]
    num_w = 0.95
    has_body = any(it.get("body") for it in items)
    title_w = max(sum(_char_w(c, ts) for c in it["title"].replace("\n", "")) for it in items) * 1.08 + 0.55
    title_w = min(max(title_w, 3.0), 5.6) if has_body else CONTENT_W - num_w - 2 * PAD
    body_x = LEFT + PAD + num_w + title_w
    body_w = RIGHT - PAD - body_x
    for it in items:  # 1行に収まる説明は、改行を外す
        if it.get("body") and n_lines(it["body"].replace("\n", ""), bs, body_w) == 1:
            it["body"] = it["body"].replace("\n", "")
    room = (bottom - top - gap * (n - 1)) / n
    while True:
        need = max(max(text_h(it["title"].replace("\n", ""), ts, title_w - 0.3, 1.2),
                       text_h(it.get("body", ""), bs, body_w, 1.35) if it.get("body") else 0)
                   for it in items) + 0.16
        if need <= room or bs <= 13:
            break
        ts, bs = ts - 1, bs - 1
    rh = min(max_h if n > 2 else 1.5, room)
    for i, it in enumerate(items):
        y = top + i * (rh + gap)
        dark = it.get("dark", False)
        rect(slide, LEFT, y, CONTENT_W, rh, INK if dark else CARD, None if dark else LINE)
        if it.get("num"):
            text(slide, LEFT + PAD, y, num_w, rh, it["num"], 20, True, RED, EN,
                 anchor=MSO_ANCHOR.MIDDLE)
        text(slide, LEFT + PAD + num_w, y, title_w - 0.3, rh, it["title"].replace("\n", ""), ts, True,
             WHITE if dark else INK, anchor=MSO_ANCHOR.MIDDLE, spacing=1.2)
        if it.get("body"):
            rect(slide, body_x - 0.2, y + 0.24, 0.015, rh - 0.48, LINE)
            text(slide, body_x, y, body_w, rh, it["body"], bs, False, GRAY,
                 anchor=MSO_ANCHOR.MIDDLE, spacing=1.35)
    return top + n * rh + (n - 1) * gap


# ---- スライドの種類 -------------------------------------------------------
def s_cover(slide, d, deck):
    background(slide)
    rect(slide, 0, 0, 0.09, H, RED)
    text(slide, LEFT, 1.3, 10, 0.4, f"CHAPTER {deck['chapter']}  /  {deck['no']}", 16, True, RED, EN)
    text(slide, LEFT, 1.75, 11, 1.1, "JLINE METHOD", 60, True, INK, EN)
    text(slide, LEFT, 3.15, 11.5, 0.9, deck["title"], 38, True)
    rect(slide, LEFT, 4.2, 0.9, 0.06, RED)
    text(slide, LEFT, 4.5, 11, 0.5, deck["subtitle"], 19, False, GRAY)
    text(slide, RIGHT - 6, 6.7, 6, 0.3, "株式会社アスリートバリュー　鈴木崇文", 13, False, GRAY,
         align=PP_ALIGN.RIGHT)


def s_cards(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub"))
    bottom = 5.75 if d.get("band") else 6.95
    if d.get("layout") == "list":
        used = rows_list(slide, top, d["items"], bottom)
    else:
        cols = d.get("cols", len(d["items"]))
        used = grid(slide, top, d["items"], cols, bottom)
    if d.get("band"):
        place_band(slide, used, d["band"])


def s_statement(slide, d, deck):
    lines = d["statement"].count("\n") + 1
    sh = lines * 38 * 1.35 * 1.2 / 72
    y = 1.9
    rect(slide, LEFT, y + 0.05, 0.1, sh - 0.1, RED)
    text(slide, LEFT + 0.5, y, CONTENT_W - 0.6, sh, d["statement"], 38, True, spacing=1.35)
    if d.get("sub"):
        text(slide, LEFT + 0.5, y + sh + 0.55, CONTENT_W - 0.6, 1.3, d["sub"], 19, False, GRAY, spacing=1.55)


def s_message(slide, d, deck):
    if d.get("items"):
        top = headline(slide, d["headline"], d.get("sub"))
        cols = d.get("cols", len(d["items"]))
        used = grid(slide, top, d["items"], cols, 5.75)
        place_band(slide, used, d["band"])
        return
    lines = d["headline"].count("\n") + 1
    hh = lines * 38 * 1.3 * 1.2 / 72
    y = 1.55
    rect(slide, LEFT, y + 0.05, 0.1, hh - 0.1, RED)
    text(slide, LEFT + 0.5, y, CONTENT_W - 0.6, hh, d["headline"], 38, True, spacing=1.3)
    y += hh + 0.35
    if d.get("sub"):
        text(slide, LEFT + 0.5, y, CONTENT_W - 0.6, 0.5, d["sub"], 19, False, GRAY)
        y += 0.5
    band(slide, max(4.9, y + 0.6), d["band"], h=1.05, size=23)


def s_flow(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub"))
    items = d["items"]
    per_row = d.get("per_row", len(items))
    rows = (len(items) + per_row - 1) // per_row
    gap = 0.2 if per_row > 4 else 0.34
    bw = (CONTENT_W - gap * (per_row - 1)) / per_row
    pad = 0.15 if per_row > 4 else 0.3
    inner = bw - 2 * pad
    ts = 18 if per_row > 4 else 20
    bs = 14 if per_row > 4 else 15
    t_need = max(text_h(it["title"], ts, inner, 1.2) for it in items)
    b_need = max((text_h(it["body"], bs, inner, 1.3) for it in items if it.get("body")), default=0)
    need = pad + 0.42 + t_need + (0.3 + b_need if b_need else 0) + pad
    bottom = 5.75 if d.get("band") else 6.95
    bh = min(max(need, 1.6), (bottom - top - gap * (rows - 1)) / rows)
    for i, it in enumerate(items):
        r, c = divmod(i, per_row)
        x = LEFT + c * (bw + gap)
        y = top + r * (bh + gap)
        dark = it.get("dark", False)
        rect(slide, x, y, bw, bh, INK if dark else CARD, None if dark else LINE)
        text(slide, x + pad, y + pad, inner, 0.3, f"{i + 1:02d}", 13, True, RED, EN)
        text(slide, x + pad, y + pad + 0.42, inner, t_need, it["title"], ts, True,
             WHITE if dark else INK, spacing=1.2)
        if it.get("body"):
            text(slide, x + pad, y + bh - pad - b_need, inner, b_need, it["body"], bs, False,
                 RGBColor(0xDD, 0xDD, 0xDB) if dark else GRAY, spacing=1.3)
        if c < per_row - 1 and i < len(items) - 1:
            text(slide, x + bw, y + bh / 2 - 0.2, gap, 0.4, "→", 16, True, RED, align=PP_ALIGN.CENTER)
    if d.get("band"):
        place_band(slide, top + rows * bh + (rows - 1) * gap, d["band"])


def s_numbers(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub"))
    items = d["items"]
    n = len(items)
    gap = 0.34
    arrows = d.get("arrows", False)
    cw = (CONTENT_W - gap * (n - 1)) / n
    pad = 0.28
    inner = cw - 2 * pad
    big = min(fit_size(it["value"], inner, 60, 28) for it in items)
    note_h = max((text_h(it["note"], 14, inner, 1.3) for it in items if it.get("note")), default=0)
    ch = pad + 0.45 + big * 1.25 / 72 + 0.15 + note_h + pad
    for i, it in enumerate(items):
        x = LEFT + i * (cw + gap)
        rect(slide, x, top, cw, ch, CARD, LINE)
        text(slide, x + pad, top + pad, inner, 0.35, it["label"], 16, True, INK)
        text(slide, x + pad, top + pad + 0.45, inner, big * 1.25 / 72, it["value"], big, True, RED, EN,
             spacing=1.0)
        if it.get("note"):
            text(slide, x + pad, top + ch - pad - note_h, inner, note_h, it["note"], 14, False, GRAY,
                 spacing=1.3)
        if arrows and i < n - 1:
            text(slide, x + cw, top + ch / 2 - 0.2, gap, 0.4, "→", 16, True, RED, align=PP_ALIGN.CENTER)
    y = top + ch
    if d.get("note"):
        text(slide, LEFT, y + 0.3, CONTENT_W, 0.5, d["note"], 16, False, INK, spacing=1.4)
        y += 0.75
    if d.get("band"):
        place_band(slide, y, d["band"])


def s_compare(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub"))
    gap = 0.34
    cw = (CONTENT_W - gap) / 2
    rows = d["rows"]
    hh = 0.72
    rect(slide, LEFT, top, cw, hh, CARD, LINE)
    text(slide, LEFT + 0.35, top, cw - 0.7, hh, d["left"], 18, True, GRAY, anchor=MSO_ANCHOR.MIDDLE)
    rect(slide, LEFT + cw + gap, top, cw, hh, INK)
    text(slide, LEFT + cw + gap + 0.35, top, cw - 0.7, hh, d["right"], 18, True, WHITE,
         anchor=MSO_ANCHOR.MIDDLE)
    rh = min(1.1, (6.95 - top - hh - 0.1) / len(rows))
    for i, (l, r) in enumerate(rows):
        y = top + hh + 0.1 + i * rh
        rect(slide, LEFT, y + rh - 0.02, CONTENT_W, 0.015, LINE)
        text(slide, LEFT + 0.35, y, cw - 0.7, rh, l, 18, False, GRAY, anchor=MSO_ANCHOR.MIDDLE)
        text(slide, LEFT + cw + gap + 0.35, y, cw - 0.7, rh, r, 19, True, INK, anchor=MSO_ANCHOR.MIDDLE)
        text(slide, LEFT + cw - 0.1, y, gap + 0.2, rh, "→", 18, True, RED, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)


def s_timeline(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub")) + 0.05
    items = d["items"]
    rh = min(0.88, (6.95 - top) / len(items))
    dx = LEFT + 2.25
    rect(slide, dx + 0.1, top + 0.2, 0.03, rh * (len(items) - 1), LINE)
    for i, it in enumerate(items):
        y = top + i * rh
        text(slide, LEFT, y, 2.0, 0.4, it["date"], 17, True, INK, EN if it["date"].isascii() else JP)
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(dx), Inches(y + 0.1),
                                     Inches(0.23), Inches(0.23))
        dot.fill.solid()
        dot.fill.fore_color.rgb = RED if it.get("accent") else INK
        dot.line.fill.background()
        text(slide, dx + 0.55, y, RIGHT - dx - 0.55, 0.42, it["title"], 19, True)
        if it.get("body"):
            text(slide, dx + 0.55, y + 0.42, RIGHT - dx - 0.55, 0.35, it["body"], 14, False, GRAY)


def s_work(slide, d, deck):
    top = headline(slide, d["headline"], d.get("sub"))
    text(slide, LEFT, top - 0.08, 2, 0.32, "WORK", 14, True, RED, EN)
    box_top = top + 0.36
    hint_h = 0.6 if d.get("hint") else 0
    rect(slide, LEFT, box_top, CONTENT_W, 6.95 - box_top, WHITE, LINE)
    n = len(d["lines"])
    pitch = min(1.15, (6.95 - box_top - 0.5 - hint_h) / n)
    y = box_top + 0.45
    for line in d["lines"]:
        text(slide, LEFT + 0.5, y, CONTENT_W - 1.0, 0.45, line, 20, True)
        rect(slide, LEFT + 0.5, y + 0.6, CONTENT_W - 1.0, 0.012, LINE)
        y += pitch
    if d.get("hint"):
        text(slide, LEFT + 0.5, 6.95 - 0.3 - 0.35, CONTENT_W - 1.0, 0.35, d["hint"], 14, False, GRAY)


def s_summary(slide, d, deck):
    top = headline(slide, d.get("headline", "まとめ"))
    y = top
    rh, gap = 0.9, 0.18
    for item in d["checks"]:
        rect(slide, LEFT, y, CONTENT_W, rh, CARD, LINE)
        rect(slide, LEFT + 0.38, y + rh / 2 - 0.15, 0.3, 0.3, WHITE, INK)
        text(slide, LEFT + 0.42, y + rh / 2 - 0.2, 0.25, 0.4, "✓", 16, True, RED, align=PP_ALIGN.CENTER)
        text(slide, LEFT + 0.95, y, CONTENT_W - 1.3, rh, item, 21, True, anchor=MSO_ANCHOR.MIDDLE)
        y += rh + gap
    band(slide, y + 0.32, f"次の回　{d['next']}", h=0.95, size=21)


def s_report(slide, d, deck):
    """章末レポート。左に質問、右に今の状態チェック（ある場合）。"""
    top = headline(slide, d["headline"], d.get("sub"))
    has_check = bool(d.get("checks"))
    gap = 0.34
    left_w = CONTENT_W * 0.58 if (has_check or d.get("fields")) else CONTENT_W
    text(slide, LEFT, top - 0.08, 3, 0.32, "REPORT", 14, True, RED, EN)
    y0 = top + 0.36
    bottom = 6.95
    n = len(d["questions"])
    qgap = 0.16
    rh = min(1.35, (bottom - y0 - qgap * (n - 1)) / n)
    for i, q in enumerate(d["questions"]):
        y = y0 + i * (rh + qgap)
        rect(slide, LEFT, y, left_w, rh, CARD, LINE)
        text(slide, LEFT + PAD, y, 0.8, rh, f"Q{i + 1}", 22, True, RED, EN, anchor=MSO_ANCHOR.MIDDLE)
        text(slide, LEFT + PAD + 0.8, y, left_w - 2 * PAD - 0.8, rh, q, 20, True,
             anchor=MSO_ANCHOR.MIDDLE)
    if d.get("fields"):
        x = LEFT + left_w + gap
        w = CONTENT_W - left_w - gap
        box_h = n * rh + (n - 1) * qgap
        rect(slide, x, y0, w, box_h, INK)
        text(slide, x + 0.35, y0 + 0.3, w - 0.7, 0.4, d["fields_title"], 18, True, WHITE)
        y = y0 + 1.0
        for f in d["fields"]:
            text(slide, x + 0.35, y, w - 0.7, 0.35, f, 15, True, RGBColor(0xDD, 0xDD, 0xDB))
            rect(slide, x + 0.35, y + 0.72, w - 0.7, 0.015, RGBColor(0x88, 0x88, 0x88))
            y += 0.85
        if d.get("fields_note"):
            text(slide, x + 0.35, y0 + box_h - 0.25 - 0.62, w - 0.7, 0.62, d["fields_note"], 13, False,
                 RGBColor(0xBB, 0xBB, 0xB8), spacing=1.35)
    if has_check:
        x = LEFT + left_w + gap
        w = CONTENT_W - left_w - gap
        box_h = n * rh + (n - 1) * qgap
        rect(slide, x, y0, w, box_h, INK)
        text(slide, x + 0.35, y0 + 0.3, w - 0.7, 0.4, d["check_title"], 18, True, WHITE)
        y = y0 + 0.88
        for c in d["checks"]:
            rect(slide, x + 0.35, y + 0.07, 0.24, 0.24, INK, WHITE)
            text(slide, x + 0.75, y, w - 1.1, 0.4, c, 16, False, WHITE)
            y += 0.5
        if d.get("check_note"):
            text(slide, x + 0.35, y0 + box_h - 0.25 - 0.62, w - 0.7, 0.62, d["check_note"], 13, False,
                 RGBColor(0xBB, 0xBB, 0xB8), spacing=1.35)


RENDER = {
    "cover": s_cover, "cards": s_cards, "statement": s_statement, "message": s_message,
    "flow": s_flow, "numbers": s_numbers, "compare": s_compare, "timeline": s_timeline,
    "work": s_work, "summary": s_summary, "report": s_report,
}


def slide_title(d, deck):
    if d["type"] == "cover":
        return f"表紙　{deck['title']}"
    if d["type"] == "statement":
        return d["statement"].replace("\n", "")
    if d["type"] == "summary":
        return "まとめ"
    return d["headline"].replace("\n", "")


def build(deck, out_dir, notes_map):
    notes = notes_map.get(deck["no"])
    if notes is not None and len(notes) != len(deck["slides"]):
        raise ValueError(f"{deck['no']}: 台本 {len(notes)} 本とスライド {len(deck['slides'])} 枚が合いません")
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    blank = prs.slide_layouts[6]
    for i, d in enumerate(deck["slides"], 1):
        slide = prs.slides.add_slide(blank)
        if d["type"] != "cover":
            background(slide)
            header(slide, f"{deck['no']} {deck['label']}", i)
        RENDER[d["type"]](slide, d, deck)
        if notes:
            slide.notes_slide.notes_text_frame.text = notes[i - 1]
    path = Path(out_dir) / f"{deck['no']}_{deck['title']}.pptx"
    prs.save(path)
    return path


def write_script(decks, notes_map, out_dir, chapter):
    """スマホで見られる台本（Markdown）を書き出す。"""
    lines = [f"# JLINEメソッド {chapter}章 台本", "",
             "1枚あたり約1分。【ここは自分の言葉で：…】は、自分の体験で話すところ。", ""]
    for deck in decks:
        notes = notes_map.get(deck["no"])
        if not notes:
            continue
        lines += [f"## {deck['no']} {deck['title']}", ""]
        for i, (d, n) in enumerate(zip(deck["slides"], notes), 1):
            lines += [f"### {i}. {slide_title(d, deck)}", ""]
            lines += [ln + "  " for ln in n.split("\n")] + [""]
    path = Path(out_dir) / f"{chapter}章_台本.md"
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


if __name__ == "__main__":
    # 使い方: python3 build_slides.py <出力フォルダ> [章番号 ...]（章を省くと全章）
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    chapters = sys.argv[2:] or list(CHAPTERS)
    for ch in chapters:
        decks, notes_map = CHAPTERS[ch]
        ch_out = out / f"{ch}章" if len(chapters) > 1 else out
        ch_out.mkdir(parents=True, exist_ok=True)
        for deck in decks:
            print(build(deck, ch_out, notes_map))
        if notes_map:
            print(write_script(decks, notes_map, ch_out, ch))
