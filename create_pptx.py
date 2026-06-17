from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Color scheme
NAVY = RGBColor(0x1B, 0x3A, 0x5C)
DARK_BLUE = RGBColor(0x2C, 0x5F, 0x8A)
ACCENT_BLUE = RGBColor(0x3A, 0x7C, 0xBD)
LIGHT_BLUE = RGBColor(0xD6, 0xEA, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x33, 0x33, 0x33)
GRAY = RGBColor(0x66, 0x66, 0x66)
RED = RGBColor(0xC0, 0x39, 0x2B)
ORANGE = RGBColor(0xE6, 0x7E, 0x22)
GREEN = RGBColor(0x27, 0xAE, 0x60)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
YELLOW_BG = RGBColor(0xFF, 0xF3, 0xCD)


def add_bg_rect(slide, color=NAVY):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_text_box(slide, left, top, width, height, text, font_size=18, color=BLACK, bold=False, alignment=PP_ALIGN.LEFT, font_name="Meiryo"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def add_multi_text(slide, left, top, width, height, lines, default_size=16, default_color=BLACK):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, (text, size, color, bold) in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = text
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.bold = bold
        p.font.name = "Meiryo"
        p.space_after = Pt(4)
    return txBox


def add_bullet_list(slide, left, top, width, height, items, font_size=15, color=BLACK, spacing=6):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        if isinstance(item, tuple):
            text, is_bold, item_color = item
        else:
            text, is_bold, item_color = item, False, color
        p.text = text
        p.font.size = Pt(font_size)
        p.font.color.rgb = item_color
        p.font.bold = is_bold
        p.font.name = "Meiryo"
        p.space_after = Pt(spacing)
    return txBox


# ===== SLIDE 1: Title =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg_rect(slide, NAVY)
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.08), ACCENT_BLUE)
add_rect(slide, Inches(0), Inches(7.42), prs.slide_width, Inches(0.08), ACCENT_BLUE)

add_text_box(slide, Inches(1), Inches(1.5), Inches(11), Inches(1.2),
             "インターホン・自動火災報知設備", 42, WHITE, True, PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(2.7), Inches(11), Inches(1),
             "設置マニュアル", 52, WHITE, True, PP_ALIGN.CENTER)

add_rect(slide, Inches(5), Inches(3.9), Inches(3.3), Inches(0.04), ACCENT_BLUE)

add_text_box(slide, Inches(1), Inches(4.3), Inches(11), Inches(0.6),
             "対象機器：Panasonic シンプルP-1シリーズ", 22, RGBColor(0xAA,0xCC,0xEE), False, PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(4.9), Inches(11), Inches(0.6),
             "対象読者：新入社員（初めて設備工事を行う方）", 20, RGBColor(0x88,0xAA,0xCC), False, PP_ALIGN.CENTER)

add_text_box(slide, Inches(8), Inches(6.5), Inches(4.5), Inches(0.5),
             "作成日：2026年6月17日", 14, RGBColor(0x88,0xAA,0xCC), False, PP_ALIGN.RIGHT)


# ===== SLIDE 2: TOC =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "目次", 36, WHITE, True, PP_ALIGN.LEFT)

toc_items = [
    ("1", "概要", "何のための作業か"),
    ("2", "事前準備・前提条件", "必要な資格・工具・部材"),
    ("3", "全体の大まかな流れ", "作業フロー"),
    ("4", "具体的な手順", "ステップバイステップ（10ステップ）"),
    ("5", "トラブルシューティング", "よくある質問・注意点"),
]
for i, (num, title, desc) in enumerate(toc_items):
    y = Inches(1.8) + Inches(i * 1.0)
    add_rect(slide, Inches(1.2), y, Inches(0.7), Inches(0.7), ACCENT_BLUE)
    add_text_box(slide, Inches(1.2), y + Pt(6), Inches(0.7), Inches(0.6), num, 28, WHITE, True, PP_ALIGN.CENTER)
    add_text_box(slide, Inches(2.2), y + Pt(2), Inches(5), Inches(0.4), title, 24, NAVY, True)
    add_text_box(slide, Inches(2.2), y + Pt(30), Inches(8), Inches(0.4), desc, 16, GRAY, False)


# ===== SLIDE 3: Overview =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "1. 概要 ― 何のための作業か", 32, WHITE, True)

add_text_box(slide, Inches(0.8), Inches(1.6), Inches(11.5), Inches(0.8),
             "この作業は、建物にインターホン（来客対応用の通話設備）と\n自動火災報知設備（火災を自動で検知し、警報を出す装置）を設置するものです。",
             20, BLACK, False)

add_rect(slide, Inches(0.8), Inches(3.0), Inches(5.5), Inches(3.5), LIGHT_BLUE)
add_text_box(slide, Inches(1.1), Inches(3.1), Inches(5), Inches(0.5), "シンプルP-1シリーズとは", 20, NAVY, True)
add_bullet_list(slide, Inches(1.1), Inches(3.6), Inches(5), Inches(2.8), [
    "- インターホン機能と自動火災報知機能を一体化した受信機（親機）",
    "- 各階・各部屋に設置する感知器（火災センサー）",
    "- ベル（警報音を鳴らす装置）",
    "- スピーカー（放送用）",
    "- ドアホン子機（玄関に取り付ける通話装置）",
    "  などを組み合わせたシステム",
], 15, BLACK, 4)

add_rect(slide, Inches(7.0), Inches(3.0), Inches(5.5), Inches(1.5), RGBColor(0xFD,0xED,0xEC))
add_text_box(slide, Inches(7.3), Inches(3.1), Inches(5), Inches(0.4), "この作業を行う理由", 20, RED, True)
add_bullet_list(slide, Inches(7.3), Inches(3.6), Inches(5), Inches(0.8), [
    "- 建物の居住者・利用者の安全を守るため（消防法に基づく義務設置）",
    "- 来客時の通話・解錠を行うため",
], 15, BLACK, 4)

add_rect(slide, Inches(7.0), Inches(4.8), Inches(5.5), Inches(1.7), YELLOW_BG)
add_text_box(slide, Inches(7.3), Inches(4.9), Inches(5), Inches(0.4), "重要ポイント", 20, ORANGE, True)
add_bullet_list(slide, Inches(7.3), Inches(5.4), Inches(5), Inches(1.0), [
    ("- 安全に関わる設備のため、手順を正確に守ってください", True, RED),
    "- 不明点は必ず上長に確認してください",
], 15, BLACK, 4)


# ===== SLIDE 4: Prerequisites - Qualifications & Tools =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "2. 事前準備 ― 必要な資格・工具", 32, WHITE, True)

# Qualifications
add_rect(slide, Inches(0.8), Inches(1.6), Inches(5.5), Inches(2.2), LIGHT_BLUE)
add_text_box(slide, Inches(1.1), Inches(1.7), Inches(5), Inches(0.5), "必要な資格・許可", 22, NAVY, True)
add_bullet_list(slide, Inches(1.1), Inches(2.3), Inches(5), Inches(1.4), [
    "- 工事担任者資格 または 電気工事士資格",
    "  （有資格者の監督のもとで作業してください）",
    "- 消防設備士の資格が必要な場合あり",
    "  （管轄の消防署に確認してください）",
], 15, BLACK, 3)

# Tools table
add_text_box(slide, Inches(0.8), Inches(4.1), Inches(5.5), Inches(0.5), "必要な工具一覧", 22, NAVY, True)

tools = [
    ("工具名", "用途"),
    ("プラスドライバー", "ねじ止め全般"),
    ("マイナスドライバー", "端子台への配線"),
    ("ワイヤーストリッパー", "電線の被覆をむく"),
    ("電工ナイフ", "ケーブル外装をむく"),
    ("テスター（回路計）", "配線の導通確認"),
    ("水平器", "機器を水平に取り付け"),
    ("ドリル", "壁面への穴あけ"),
    ("脚立・はしご", "高所作業用"),
]

for i, (tool, use) in enumerate(tools):
    y = Inches(4.6) + Inches(i * 0.3)
    bg_color = NAVY if i == 0 else (LIGHT_GRAY if i % 2 == 0 else WHITE)
    txt_color = WHITE if i == 0 else BLACK
    add_rect(slide, Inches(0.8), y, Inches(2.5), Inches(0.3), bg_color)
    add_rect(slide, Inches(3.3), y, Inches(3.2), Inches(0.3), bg_color)
    add_text_box(slide, Inches(0.9), y + Pt(1), Inches(2.3), Inches(0.28), tool, 12, txt_color, i == 0)
    add_text_box(slide, Inches(3.4), y + Pt(1), Inches(3), Inches(0.28), use, 12, txt_color, i == 0)

# Materials
add_rect(slide, Inches(7.0), Inches(1.6), Inches(5.5), Inches(5.5), RGBColor(0xE8,0xF5,0xE9))
add_text_box(slide, Inches(7.3), Inches(1.7), Inches(5), Inches(0.5), "必要な部材（施工図面で確認）", 22, GREEN, True)
add_bullet_list(slide, Inches(7.3), Inches(2.3), Inches(5), Inches(4.5), [
    ("- 受信機（親機）", True, BLACK),
    "  火災受信・インターホン制御を行う本体",
    ("- 感知器", True, BLACK),
    "  煙や熱を検知するセンサー（天井に取り付け）",
    ("- 発信機", True, BLACK),
    "  手動で火災信号を送る押しボタン",
    ("- ベル・ブザー", True, BLACK),
    "  火災時に警報音を鳴らす装置",
    ("- スピーカー（SP）", True, BLACK),
    "  放送用",
    ("- ドアホン子機", True, BLACK),
    "  玄関等に設置する通話装置",
    ("- 住戸用受信機", True, BLACK),
    "  各住戸内の通話・報知装置",
    ("- 配線用ケーブル", True, BLACK),
    "  施工図面で指定された種類・本数",
    ("- 取付ねじ・アンカー", True, BLACK),
    "  壁面固定用",
], 13, GRAY, 1)


# ===== SLIDE 5: Prerequisites - Pre-check =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "2. 事前準備 ― 事前確認事項", 32, WHITE, True)

add_rect(slide, Inches(0.8), Inches(1.6), Inches(11.5), Inches(5.2), YELLOW_BG)
add_text_box(slide, Inches(1.2), Inches(1.8), Inches(10), Inches(0.5), "作業開始前に必ず確認してください", 24, ORANGE, True)

checks = [
    ("施工図面を必ず入手し、内容を理解してから作業を開始してください", True),
    ("設置場所の壁の材質（コンクリート・石膏ボード等）を確認する", False),
    ("電源の位置と容量を確認する（AC100V電源が必要）", False),
    ("配線ルートに障害物がないか事前に現地を確認する", False),
    ("各機器の型番と数量が図面と一致しているか確認する", False),
    ("作業に必要な工具が全て揃っているか確認する", False),
    ("作業当日の天候を確認する（屋外作業がある場合）", False),
]

for i, (text, is_important) in enumerate(checks):
    y = Inches(2.6) + Inches(i * 0.6)
    color = RED if is_important else BLACK
    mark = "!!" if is_important else str(i + 1)
    mark_bg = RED if is_important else ACCENT_BLUE
    add_rect(slide, Inches(1.5), y, Inches(0.45), Inches(0.45), mark_bg)
    add_text_box(slide, Inches(1.5), y + Pt(3), Inches(0.45), Inches(0.4), mark, 14, WHITE, True, PP_ALIGN.CENTER)
    add_text_box(slide, Inches(2.2), y + Pt(3), Inches(9.5), Inches(0.4),
                 text, 17 if is_important else 16, color, is_important)


# ===== SLIDE 6: Overall Flow =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "3. 全体の大まかな流れ", 32, WHITE, True)

steps = [
    ("1", "施工図面の確認\n現地調査"),
    ("2", "受信機の設置場所\n準備・取付金具固定"),
    ("3", "各機器の取付\n(感知器・ベル等)"),
    ("4", "配線工事\n(ケーブル接続)"),
    ("5", "受信機本体の\n取付・配線接続"),
    ("6", "蓄電池\nの接続"),
    ("7", "電源投入\n動作確認テスト"),
    ("8", "銘板記入\n仕上げ"),
    ("9", "消防検査\nの立会い"),
]

for i, (num, label) in enumerate(steps):
    x = Inches(0.5) + Inches(i * 1.4)
    y = Inches(2.5)
    # Circle-like rounded rectangle
    shape = add_rect(slide, x, y, Inches(1.15), Inches(1.15), ACCENT_BLUE if i < 8 else GREEN)
    add_text_box(slide, x, y + Pt(4), Inches(1.15), Inches(0.35), num, 22, WHITE, True, PP_ALIGN.CENTER)
    add_text_box(slide, x - Inches(0.15), y + Inches(1.3), Inches(1.45), Inches(0.8), label, 12, BLACK, False, PP_ALIGN.CENTER)

    # Arrow
    if i < 8:
        arrow_x = x + Inches(1.18)
        add_text_box(slide, arrow_x, y + Pt(14), Inches(0.25), Inches(0.5), ">", 24, ACCENT_BLUE, True, PP_ALIGN.CENTER)

# Important note at bottom
add_rect(slide, Inches(0.8), Inches(5.0), Inches(11.5), Inches(1.8), YELLOW_BG)
add_text_box(slide, Inches(1.2), Inches(5.1), Inches(10), Inches(0.5), "作業の重要ポイント", 20, ORANGE, True)
add_bullet_list(slide, Inches(1.2), Inches(5.6), Inches(10.5), Inches(1.2), [
    ("- ステップ5「受信機への配線接続」が最も重要な工程です", True, RED),
    "- 各ステップ完了後、次に進む前に必ず確認を行ってください",
    "- 不明点は自己判断せず、必ず上長に確認してください",
    "- 電源に関わる作業は必ず電源OFFの状態で行ってください",
], 15, BLACK, 3)


# ===== SLIDE 7: Step 1-2 =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "4. 具体的な手順 ― ステップ1・2", 32, WHITE, True)

# Step 1
add_rect(slide, Inches(0.5), Inches(1.5), Inches(6), Inches(5.5), LIGHT_BLUE)
add_rect(slide, Inches(0.5), Inches(1.5), Inches(6), Inches(0.6), ACCENT_BLUE)
add_text_box(slide, Inches(0.7), Inches(1.55), Inches(5.5), Inches(0.5), "ステップ1：施工図面の確認・現地調査", 20, WHITE, True)
add_bullet_list(slide, Inches(0.8), Inches(2.3), Inches(5.5), Inches(4.5), [
    "- 施工図面で以下を確認する：",
    "  ・受信機の設置場所（通常は1階の管理室や玄関付近）",
    "  ・各階の感知器の数と位置",
    "  ・ベル・発信機の設置位置",
    "  ・ドアホン子機の設置位置",
    "  ・配線ルート",
    "",
    "- 現地で壁の材質・スペースを確認する",
    "",
    ("- 図面と現地が異なる場合は、必ず上長に報告", True, RED),
], 14, BLACK, 2)

# Step 2
add_rect(slide, Inches(6.8), Inches(1.5), Inches(6), Inches(5.5), RGBColor(0xE8,0xF5,0xE9))
add_rect(slide, Inches(6.8), Inches(1.5), Inches(6), Inches(0.6), GREEN)
add_text_box(slide, Inches(7.0), Inches(1.55), Inches(5.5), Inches(0.5), "ステップ2：受信機（親機）の取付準備", 20, WHITE, True)
add_bullet_list(slide, Inches(7.1), Inches(2.3), Inches(5.5), Inches(4.5), [
    "- 壁面に取付用の穴をあける",
    "  本体外形：横370mm x 縦500mm程度",
    "  取付穴間隔：横310mm x 縦450mm",
    "  （機種により異なるため図面で確認）",
    "",
    "- 壁面にアンカー（壁に固定するための部品）を打ち込む",
    "",
    "- 取付板（受信機の裏板）をねじで固定する",
    "",
    ("- 取付位置は床から約1.5mの高さに", True, RED),
    ("- 水平器で水平を確認してから固定する", True, RED),
], 14, BLACK, 2)


# ===== SLIDE 8: Step 3-4 =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "4. 具体的な手順 ― ステップ3・4", 32, WHITE, True)

# Step 3
add_rect(slide, Inches(0.5), Inches(1.5), Inches(6), Inches(5.5), LIGHT_BLUE)
add_rect(slide, Inches(0.5), Inches(1.5), Inches(6), Inches(0.6), ACCENT_BLUE)
add_text_box(slide, Inches(0.7), Inches(1.55), Inches(5.5), Inches(0.5), "ステップ3：配線の引き込み", 20, WHITE, True)
add_bullet_list(slide, Inches(0.8), Inches(2.3), Inches(5.5), Inches(4.5), [
    "- 施工図面に従い、受信機設置位置まで",
    "  ケーブルを配線する",
    "",
    "- ケーブルは受信機の下部または側面の",
    "  ノックアウト穴（打ち抜き穴）から引き込む",
    "",
    ("- ケーブルの被覆むき長さ：約12mm", True, RED),
    ("- 外装むき長さ：約10mm", True, RED),
    "",
    "- ケーブルをまとめて整理し、",
    "  余長（予備の長さ）を適切に取る",
], 14, BLACK, 2)

# Step 4
add_rect(slide, Inches(6.8), Inches(1.5), Inches(6), Inches(5.5), RGBColor(0xE8,0xF5,0xE9))
add_rect(slide, Inches(6.8), Inches(1.5), Inches(6), Inches(0.6), GREEN)
add_text_box(slide, Inches(7.0), Inches(1.55), Inches(5.5), Inches(0.5), "ステップ4：感知器の取付", 20, WHITE, True)
add_bullet_list(slide, Inches(7.1), Inches(2.3), Inches(5.5), Inches(4.8), [
    "感知器＝煙や熱を感知して火災信号を",
    "受信機に送る装置。天井面に取り付けます。",
    "",
    "1. 天井に感知器用の台座をねじで固定",
    "2. 台座に配線を接続",
    "   ・L端子（ライン）：信号線（＋側）",
    "   ・C端子（コモン）：共通線（−側）",
    "3. 感知器本体を台座にはめ込む",
    "   （右に回してロック）",
    "4. 最後の感知器に終端抵抗器（10kΩ）を取付",
    "",
    ("- 壁から60cm以上離して設置", True, RED),
    ("- エアコン吹出口の近くには設置不可", True, RED),
    "  （誤報の原因になります）",
], 14, BLACK, 2)


# ===== SLIDE 9: Step 5-6 =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "4. 具体的な手順 ― ステップ5・6", 32, WHITE, True)

# Step 5
add_rect(slide, Inches(0.5), Inches(1.5), Inches(6), Inches(5.5), LIGHT_BLUE)
add_rect(slide, Inches(0.5), Inches(1.5), Inches(6), Inches(0.6), ACCENT_BLUE)
add_text_box(slide, Inches(0.7), Inches(1.55), Inches(5.5), Inches(0.5), "ステップ5：ベル・発信機の取付", 20, WHITE, True)
add_bullet_list(slide, Inches(0.8), Inches(2.3), Inches(5.5), Inches(4.5), [
    "【ベル（警報音装置）】",
    "- 指定の位置に取り付ける",
    "- BL端子・BC端子に配線を接続",
    "",
    "【発信機（手動の火災報知ボタン）】",
    "- 廊下等の指定位置に取り付ける",
    ("- 設置高さ：床から0.8m〜1.5m", True, RED),
    "",
    "【ドアホン子機】",
    "- 玄関等に取り付ける",
    "",
    "【住戸用受信機】",
    "- 各部屋のインターホン親機を取り付ける",
], 14, BLACK, 2)

# Step 6
add_rect(slide, Inches(6.8), Inches(1.5), Inches(6), Inches(5.5), RGBColor(0xFD,0xED,0xEC))
add_rect(slide, Inches(6.8), Inches(1.5), Inches(6), Inches(0.6), RED)
add_text_box(slide, Inches(7.0), Inches(1.55), Inches(5.5), Inches(0.5), "ステップ6：受信機への配線接続 [最重要]", 18, WHITE, True)
add_bullet_list(slide, Inches(7.1), Inches(2.3), Inches(5.5), Inches(4.8), [
    ("この工程が最も重要です！", True, RED),
    "",
    "接続手順：",
    "1. 電源スイッチがOFFになっていることを確認",
    "2. 端子台のねじをゆるめる",
    "3. ケーブルの芯線を端子台に差し込む",
    "4. ねじをしっかり締める",
    "5. 軽く引っ張って抜けないことを確認",
    "6. 全配線完了後、結線表を作成",
    "",
    ("- 端子記号を間違えると機器が動作しません", True, RED),
    ("- 1本ずつ確認しながら接続してください", True, RED),
    ("- 電源線（P1, P2）は最後に接続", True, RED),
], 14, BLACK, 2)


# ===== SLIDE 10: Terminal Reference =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "4. 端子記号 一覧表（配線時に参照）", 32, WHITE, True)

terminals = [
    ("端子記号", "意味", "接続先"),
    ("P1, P2", "電源端子", "AC100V電源"),
    ("L1, L2, Ln", "感知器回線", "各階の感知器"),
    ("C", "コモン（共通線）", "感知器の共通線"),
    ("A, T", "アドレス・電話", "インターホン回線"),
    ("B, BC", "ベル回線", "ベル・ブザー"),
    ("BL", "ベルライン", "ベル信号線"),
    ("EA, EL", "戸外表示灯", "住戸用表示灯"),
    ("EF, EC", "戸外通話", "住戸用通話"),
    ("EB", "戸外ブザー", "住戸用ブザー"),
    ("SP1, SP2", "スピーカー", "スピーカー"),
    ("Fc, Fb, Fa", "防排煙連動", "防火戸・排煙装置"),
    ("D, DC", "戸開放回線", "自動ドア連動"),
    ("H, HL", "放送回線", "放送設備"),
    ("U, UL", "非常電源", "非常電源装置"),
    ("N, NC", "地区音響", "地区ベル"),
]

col_widths = [Inches(2.2), Inches(3.0), Inches(3.5)]
x_starts = [Inches(2.0), Inches(4.2), Inches(7.2)]

for i, (sym, meaning, dest) in enumerate(terminals):
    y = Inches(1.5) + Inches(i * 0.35)
    bg = NAVY if i == 0 else (LIGHT_GRAY if i % 2 == 0 else WHITE)
    txt = WHITE if i == 0 else BLACK
    for j, (x, w, t) in enumerate(zip(x_starts, col_widths, [sym, meaning, dest])):
        add_rect(slide, x, y, w, Inches(0.35), bg)
        add_text_box(slide, x + Pt(6), y + Pt(2), w - Pt(12), Inches(0.3), t, 13, txt, i == 0 or j == 0)


# ===== SLIDE 11: Step 7-8 =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "4. 具体的な手順 ― ステップ7・8", 32, WHITE, True)

# Step 7
add_rect(slide, Inches(0.5), Inches(1.5), Inches(6), Inches(5.5), RGBColor(0xFD,0xED,0xEC))
add_rect(slide, Inches(0.5), Inches(1.5), Inches(6), Inches(0.6), RED)
add_text_box(slide, Inches(0.7), Inches(1.55), Inches(5.5), Inches(0.5), "ステップ7：蓄電池（バッテリー）の接続", 20, WHITE, True)
add_bullet_list(slide, Inches(0.8), Inches(2.3), Inches(5.5), Inches(4.5), [
    "蓄電池＝停電時にも設備を動かすための予備電源",
    "",
    "1. 受信機内部の蓄電池収納スペースに",
    "   蓄電池を設置する",
    "",
    "2. プラス（+）とマイナス（-）を正しく接続",
    "   ・赤い線 → プラス（+）",
    "   ・黒い線 → マイナス（-）",
    "",
    ("極性（+と-）を逆に接続すると", True, RED),
    ("機器が故障します。必ず確認！", True, RED),
    "",
    "- 蓄電池はNi-Cd（ニッケルカドミウム）タイプ",
    "- 廃棄時はリサイクルに出してください",
], 14, BLACK, 2)

# Step 8
add_rect(slide, Inches(6.8), Inches(1.5), Inches(6), Inches(5.5), RGBColor(0xE8,0xF5,0xE9))
add_rect(slide, Inches(6.8), Inches(1.5), Inches(6), Inches(0.6), GREEN)
add_text_box(slide, Inches(7.0), Inches(1.55), Inches(5.5), Inches(0.5), "ステップ8：電源投入・動作確認テスト", 20, WHITE, True)
add_bullet_list(slide, Inches(7.1), Inches(2.3), Inches(5.5), Inches(4.8), [
    "1. 全ての配線が完了していることを再確認",
    "2. 受信機の扉を閉める",
    "3. AC100V電源を投入（ブレーカーON）",
    "4. 受信機の電源スイッチをONにする",
    "",
    "【火災報知系の確認】",
    "- 各感知器回線が「正常」と表示されるか",
    "- 加煙試験器で感知器を作動させ「火災」表示確認",
    "- ベルが鳴動するか確認",
    "- 発信機ボタンで受信機が反応するか確認",
    "",
    "【インターホン系の確認】",
    "- ドアホン子機の呼出→住戸機が鳴動するか",
    "- 通話ができるか",
    "- 解錠機能の動作確認",
    "",
    ("- 異常時は電源を切ってから配線確認", True, RED),
], 14, BLACK, 2)


# ===== SLIDE 12: Step 9-10 =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "4. 具体的な手順 ― ステップ9・10", 32, WHITE, True)

# Step 9
add_rect(slide, Inches(0.5), Inches(1.5), Inches(6), Inches(5.0), LIGHT_BLUE)
add_rect(slide, Inches(0.5), Inches(1.5), Inches(6), Inches(0.6), ACCENT_BLUE)
add_text_box(slide, Inches(0.7), Inches(1.55), Inches(5.5), Inches(0.5), "ステップ9：銘板の記入・仕上げ", 20, WHITE, True)
add_bullet_list(slide, Inches(0.8), Inches(2.3), Inches(5.5), Inches(4.0), [
    "- 受信機前面の銘板に以下を記入：",
    "  ・設置場所",
    "  ・設置年月日",
    "  ・施工業者名",
    "",
    "- DIPスイッチ（機器の設定スイッチ）を",
    "  施工図面の指示通りに設定（OP1〜OP4）",
    "",
    "- 受信機内部を清掃し、扉を閉めて施錠",
    "",
    "- 施工完了報告書を作成する",
], 14, BLACK, 2)

# Step 10
add_rect(slide, Inches(6.8), Inches(1.5), Inches(6), Inches(5.0), RGBColor(0xE8,0xF5,0xE9))
add_rect(slide, Inches(6.8), Inches(1.5), Inches(6), Inches(0.6), GREEN)
add_text_box(slide, Inches(7.0), Inches(1.55), Inches(5.5), Inches(0.5), "ステップ10：消防検査の立会い", 20, WHITE, True)
add_bullet_list(slide, Inches(7.1), Inches(2.3), Inches(5.5), Inches(4.0), [
    "- 管轄の消防署に検査申請を行う",
    "",
    "- 消防検査員の立会いのもと、",
    "  全ての機器の動作確認を行う",
    "",
    "- 検査で指摘された事項があれば",
    "  速やかに是正する",
    "",
    "- 検査合格後、検査済証を受領する",
    "",
    ("- 消防検査に合格するまで", True, RED),
    ("  建物の使用は開始できません", True, RED),
], 14, BLACK, 2)


# ===== SLIDE 13: Troubleshooting =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "5. トラブルシューティング", 32, WHITE, True)

troubles = [
    ("症状", "考えられる原因", "対処法"),
    ("電源が入らない", "電源ケーブルの接続不良\nブレーカーがOFF", "電源配線を確認\nブレーカーをONにする"),
    ("「断線」の表示", "配線が切れている\n端子接続が外れている", "該当回線の配線を\n1本ずつ確認"),
    ("感知器が反応しない", "台座にしっかり\nはまっていない/配線ミス", "感知器を付け直す\nL・C端子の配線確認"),
    ("誤報が出る", "感知器周辺の環境\n終端抵抗の未接続", "設置環境を確認\n終端抵抗器(10kΩ)確認"),
    ("ベルが鳴らない", "ベル回線の配線ミス\nベル本体の故障", "BL・BC端子の接続確認\nベル本体を交換して試す"),
    ("通話ができない", "通話回線の配線ミス", "EA・EL・EF・EC\n端子の接続確認"),
    ("充電ランプ不点灯", "蓄電池の接続不良\n極性間違い", "+/-の接続を確認"),
]

col_w = [Inches(2.5), Inches(3.5), Inches(4.0)]
col_x = [Inches(1.2), Inches(3.7), Inches(7.2)]

for i, (s, c, f) in enumerate(troubles):
    y = Inches(1.5) + Inches(i * 0.72)
    bg = NAVY if i == 0 else (RGBColor(0xFD,0xED,0xEC) if i % 2 == 1 else LIGHT_GRAY)
    txt = WHITE if i == 0 else BLACK
    for j, (x, w, t) in enumerate(zip(col_x, col_w, [s, c, f])):
        add_rect(slide, x, y, w, Inches(0.7), bg)
        add_text_box(slide, x + Pt(6), y + Pt(3), w - Pt(12), Inches(0.65), t, 12, txt, i == 0)


# ===== SLIDE 14: FAQ =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), NAVY)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "5. よくある質問（FAQ）", 32, WHITE, True)

faqs = [
    ("Q1. 配線のケーブル種類がわかりません。",
     "A1. 施工図面に記載されています。一般的に、感知器回線にはHP（耐熱ペア線）、\nインターホン回線にはFCPV（ビニル被覆ペア線）等を使用します。不明な場合は上長に確認。"),
    ("Q2. 終端抵抗器はどこに付けますか？",
     "A2. 各回線の最後の機器（一番遠い感知器等）に取り付けます。\n受信機から見て最も末端の機器です。抵抗値は10kΩです。"),
    ("Q3. 感知器の種類の使い分けは？",
     "A3. 煙感知器は居室・廊下に、熱感知器はキッチン・浴室付近に使用します。\n施工図面の指示に従ってください。"),
    ("Q4. 配線を間違えた場合は？",
     "A4. まず電源を切ってください。その後、結線表と照らし合わせて正しい端子に接続し直してください。"),
    ("Q5. 作業中に分からないことが出たら？",
     "A5. 自己判断で作業を進めないでください。必ず上長または先輩社員に相談してください。\n安全に関わる設備ですので、確認を怠ると人命に関わります。"),
]

for i, (q, a) in enumerate(faqs):
    y = Inches(1.5) + Inches(i * 1.15)
    add_rect(slide, Inches(0.8), y, Inches(11.5), Inches(1.1), LIGHT_BLUE if i % 2 == 0 else LIGHT_GRAY)
    add_text_box(slide, Inches(1.1), y + Pt(4), Inches(11), Inches(0.35), q, 15, NAVY, True)
    add_text_box(slide, Inches(1.1), y + Pt(26), Inches(11), Inches(0.7), a, 13, BLACK, False)


# ===== SLIDE 15: Safety =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.2), RED)
add_text_box(slide, Inches(0.8), Inches(0.25), Inches(11), Inches(0.8), "安全上の重要注意事項", 36, WHITE, True)

safety_items = [
    "感電に注意してください。配線作業は必ず電源を切った状態で行ってください",
    "濡れた手で作業しないでください",
    "本体を分解しないでください。修理はメーカーまたは有資格者が行います",
    "必ずアース線（接地線）を接続してください。漏電による感電を防ぎます",
    "高所作業時は必ず安全帯を使用してください",
    "作業後は必ず消防検査を受けてください",
]

for i, item in enumerate(safety_items):
    y = Inches(1.8) + Inches(i * 0.9)
    add_rect(slide, Inches(1.0), y, Inches(11.3), Inches(0.75), RGBColor(0xFD,0xED,0xEC))
    add_rect(slide, Inches(1.0), y, Inches(0.08), Inches(0.75), RED)
    add_rect(slide, Inches(1.3), y + Pt(8), Inches(0.45), Inches(0.45), RED)
    add_text_box(slide, Inches(1.3), y + Pt(11), Inches(0.45), Inches(0.4), str(i+1), 16, WHITE, True, PP_ALIGN.CENTER)
    add_text_box(slide, Inches(2.0), y + Pt(10), Inches(10), Inches(0.5), item, 18, RED, True)


# ===== SLIDE 16: End =====
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg_rect(slide, NAVY)
add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.08), ACCENT_BLUE)
add_rect(slide, Inches(0), Inches(7.42), prs.slide_width, Inches(0.08), ACCENT_BLUE)

add_text_box(slide, Inches(1), Inches(2.0), Inches(11), Inches(1),
             "マニュアルは以上です", 44, WHITE, True, PP_ALIGN.CENTER)

add_rect(slide, Inches(5), Inches(3.2), Inches(3.3), Inches(0.04), ACCENT_BLUE)

add_text_box(slide, Inches(1), Inches(3.8), Inches(11), Inches(0.6),
             "不明点は自己判断せず、必ず上長・先輩社員に相談してください", 22, RGBColor(0xFF,0xAA,0xAA), True, PP_ALIGN.CENTER)

add_text_box(slide, Inches(1), Inches(5.0), Inches(11), Inches(0.5),
             "対象機器：Panasonic シンプルP-1シリーズ", 18, RGBColor(0xAA,0xCC,0xEE), False, PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(5.5), Inches(11), Inches(0.5),
             "施工説明書 8A3 P02 00003 / S0615-20423AB", 14, RGBColor(0x88,0xAA,0xCC), False, PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(6.2), Inches(11), Inches(0.5),
             "作成日：2026年6月17日", 16, RGBColor(0x88,0xAA,0xCC), False, PP_ALIGN.CENTER)


# Save
output_path = "/home/user/claude-project/インターホン・自動火災報知設備_設置マニュアル.pptx"
prs.save(output_path)
print(f"Saved: {output_path}")
