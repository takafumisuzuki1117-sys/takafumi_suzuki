"""黒田様の提案資料（10/23版）をもとに、11月からの価格の提案資料（名前なし）を作る。
使い方: python3 make_nov.py kuroda.pdf out.pdf [お名前]
"""
import sys
import pymupdf

SRC, OUT = sys.argv[1], sys.argv[2]
NAME = sys.argv[3] if len(sys.argv) > 3 else ''
IPAG = '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'
CARLITO = '/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf'

doc = pymupdf.open(SRC)
oswald_buf = doc.extract_font(6)[3]
F_IPAG = pymupdf.Font(fontfile=IPAG)
F_CARL = pymupdf.Font(fontfile=CARLITO)
F_OSW = pymupdf.Font(fontbuffer=oswald_buf)


def rgb(h):
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def spans(page):
    out = []
    for b in page.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            for s in l['spans']:
                if s['text'].strip():
                    out.append(s)
    return out


def find(page, text, y=None, x=None):
    hits = [s for s in spans(page) if s['text'].strip() == text
            and (y is None or abs(s['bbox'][1] - y) < 3) and (x is None or abs(s['bbox'][0] - x) < 3)]
    assert len(hits) == 1, (page.number + 1, text, y, len(hits))
    return hits[0]


def redact_rect(page, rect):
    r = pymupdf.Rect(rect)
    page.add_redact_annot(pymupdf.Rect(r.x0 - 0.5, r.y0 + 1, r.x1 + 0.5, r.y1 - 1))


def redact_band(page, x0, y0, x1, y1):
    """この範囲にかかる文字を全部消す"""
    for s in spans(page):
        bx0, by0, bx1, by1 = s['bbox']
        if bx0 >= x0 - 1 and bx1 <= x1 + 1 and by0 >= y0 - 1 and by1 <= y1 + 1:
            redact_rect(page, s['bbox'])


def setup_fonts(page):
    page.insert_font(fontname='ipg11', fontfile=IPAG)
    page.insert_font(fontname='crl11', fontfile=CARLITO)
    page.insert_font(fontname='osw11', fontbuffer=oswald_buf)


def runs(text):
    """英数字はCarlito、それ以外はIPAゴシック"""
    out, cur, kind = [], '', None
    for ch in text:
        k = 'carl' if (ch.isascii() and ch not in '()') else 'ipag'
        if kind is None or k == kind:
            cur += ch
        else:
            out.append((kind, cur))
            cur = ch
        kind = k
    if cur:
        out.append((kind, cur))
    return out


FONTS = {'ipag': F_IPAG, 'carl': F_CARL, 'osw': F_OSW}
PDFNAME = {'ipag': 'ipg11', 'carl': 'crl11', 'osw': 'osw11'}


def width(text, size, font=None):
    if font:
        return FONTS[font].text_length(text, size)
    return sum(FONTS[k].text_length(t, size) for k, t in runs(text))


def write(page, x, y, text, size, color, font=None, bold=False):
    """(x, y) はベースライン。bold は細い縁取りで太く見せる"""
    kw = dict(fontsize=size, color=rgb(color))
    if bold:
        kw.update(render_mode=2, fill=rgb(color), border_width=0.03)
    setup_fonts(page)
    parts = [(font, text)] if font else runs(text)
    for f, t in parts:
        page.insert_text((x, y), t, fontname=PDFNAME[f], **kw)
        x += FONTS[f].text_length(t, size)
    return x


def replace_span(page, span, new, font, align='left'):
    """span を同じ大きさ・色・ベースラインで new に置き換える"""
    redact_rect(page, span['bbox'])
    page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE, graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
    ox, oy = span['origin']
    w = width(new, span['size'], font)
    x = ox if align == 'left' else span['bbox'][2] - w
    write(page, x, oy, new, span['size'], f"{span['color']:06x}", font)


def apply(page):
    page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE, graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)


# ---------- 表紙：お名前 ----------
p = doc[0]
setup_fonts(p)
name_span = find(p, '黒田')
sama = find(p, '様')
redact_rect(p, name_span['bbox'])
apply(p)
if NAME:
    # 枠（92.5〜612.4）の中央に「お名前 様」を置き直す
    redact_rect(p, sama['bbox'])
    apply(p)
    label = f'{NAME} 様'
    w = width(label, sama['size'], 'ipag')
    write(p, (92.5 + 612.4 - w) / 2, sama['origin'][1], label, sama['size'], '6e6e6e', 'ipag')

# ---------- 006 STEP：プレミアムの注記を消す ----------
p = doc[6]
redact_band(p, 88, 692, 660, 712)
apply(p)

# ---------- 012 PLAN ----------
p = doc[12]
setup_fonts(p)
sub = find(p, 'どちらも')
foot = find(p, '※表示価格は税別です（')
box = find(p, '追加オプションの単価で換算すると、スタンダードは')
sub_y, foot_y, box_y = sub['origin'][1], foot['origin'][1], box['origin'][1]
price = find(p, '120,000', x=129.4)
m1, m2 = find(p, '20,000', x=184.5), find(p, '22,000', x=272.9)
redact_band(p, 88, 175, 880, 199)            # 見出し下の説明
redact_band(p, 170, 248, 260, 266)           # 「｜おすすめ」
redact_band(p, 750, 240, 1345, 510)          # プレミアムの中身（枠は残す）
redact_band(p, 120, 560, 880, 582)           # 相当額の行
redact_band(p, 88, 706, 830, 725)            # 一番下の注記
for s in (price, m1, m2):
    redact_rect(p, s['bbox'])
apply(p)

write(p, 90.1, sub_y, '6ヶ月で150,000円（税込165,000円）です。月あたり25,000円（税込27,500円）になります。', 18.0, '6e6e6e')
write(p, price['origin'][0], price['origin'][1], '150,000', price['size'], 'ed142c', 'osw')
write(p, m1['origin'][0], m1['origin'][1], '25,000', m1['size'], 'c8c8c8', 'carl')
write(p, m2['origin'][0], m2['origin'][1], '27,500', m2['size'], 'c8c8c8', 'carl')
write(p, 125.5, box_y, '追加オプションの単価で換算すると、180,000円相当の内容です。', 15.8, '141414')
write(p, 90.1, foot_y, '※表示価格は税別です（6ヶ月総額 税込165,000円）。　※この料金は2026年11月1日以降にお申し込みの方に適用されます。', 13.0, '6e6e6e')

# 右の枠：お支払い方法（プレミアムの「月あたり」の下にあった短い線を消す）
p.draw_rect(pymupdf.Rect(775, 430, 860, 446), color=None, fill=(0.98, 0.98, 0.973), width=0)
X = 783.4
write(p, X, 262.0, 'お支払い方法', 13.7, '6e6e6e', 'ipag')
write(p, X, 307.0, '一括でも、分割でも', 28.1, '141414', 'ipag', bold=True)
rows = [('一括', '165,000', '円'), ('2回', '82,500', '円 × 2回'), ('3回', '55,000', '円 × 3回'), ('6回', '27,500', '円 × 6回')]
y = 352.0
for label, num, unit in rows:
    write(p, X, y, label, 15.8, '444444')
    nx = X + 70
    nw = width(num, 22.0, 'osw')
    write(p, nx + 92 - nw, y + 1, num, 22.0, '141414', 'osw')   # 右そろえ
    write(p, nx + 96, y, unit, 15.8, '444444', 'ipag')
    y += 31
write(p, X, 487.7 + 12, '金額は税込。分割はクレジットカードのみ、銀行振込は一括のみです。', 13.0, '6e6e6e')

# ---------- 014 FOR FAMILY ----------
p = doc[14]
setup_fonts(p)
replace_span(p, find(p, '120,000', y=373.7), '150,000', 'carl')
replace_span(p, find(p, '132,000', y=373.7), '165,000', 'carl')

# ---------- 015 OPTION ----------
p = doc[15]
setup_fonts(p)
for old, y, new in (('25,000', 255.1, '30,000'), ('15,000', 354.4, '18,000'), ('15,000', 453.7, '18,000'), ('10,000', 553.0, '12,000')):
    replace_span(p, find(p, old, y=y), new, 'osw', align='right')

doc.subset_fonts()
doc.save(OUT, garbage=3, deflate=True)
print('saved', OUT)
