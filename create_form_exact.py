#!/usr/bin/env python3
"""
消防設備工事用書類（Word様式）をExcelに変換し、入力シートから自動反映させる統合ファイル
書式のレイアウト・マス幅を忠実に再現
"""
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# Styles
F_TITLE = Font(name='ＭＳ ゴシック', bold=True, size=14)
F_HEADER = Font(name='ＭＳ ゴシック', bold=True, size=9)
F_NORMAL = Font(name='ＭＳ ゴシック', size=9)
F_SMALL = Font(name='ＭＳ ゴシック', size=7.5)
F_SECTION = Font(name='ＭＳ ゴシック', bold=True, size=11)

THIN = Border(
    left=Side('thin'), right=Side('thin'),
    top=Side('thin'), bottom=Side('thin'))
THIN_L = Border(left=Side('thin'), top=Side('thin'), bottom=Side('thin'))
THIN_R = Border(right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
THIN_T = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'))
THIN_B = Border(left=Side('thin'), right=Side('thin'), bottom=Side('thin'))
THICK = Border(
    left=Side('medium'), right=Side('medium'),
    top=Side('medium'), bottom=Side('medium'))

FILL_INPUT = PatternFill('solid', fgColor='FFFFCC')
FILL_AUTO = PatternFill('solid', fgColor='E0FFE0')
FILL_HEADER = PatternFill('solid', fgColor='D9E1F2')
FILL_SECTION = PatternFill('solid', fgColor='B4C6E7')
FILL_WARN = PatternFill('solid', fgColor='FFE0E0')

AL_C = Alignment(horizontal='center', vertical='center', wrap_text=True)
AL_L = Alignment(horizontal='left', vertical='center', wrap_text=True)
AL_R = Alignment(horizontal='right', vertical='center', wrap_text=True)
AL_TL = Alignment(horizontal='left', vertical='top', wrap_text=True)

def sc(ws, r, c, val=None, font=F_NORMAL, al=AL_C, border=THIN, fill=None):
    """Set cell with styling"""
    cell = ws.cell(row=r, column=c, value=val)
    if font: cell.font = font
    if al: cell.alignment = al
    if border: cell.border = border
    if fill: cell.fill = fill
    return cell

def border_range(ws, r1, r2, c1, c2, border=THIN):
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            ws.cell(row=r, column=c).border = border

def input_cell(ws, r, c, val=None):
    return sc(ws, r, c, val, fill=FILL_INPUT)

def auto_cell(ws, r, c, val=None):
    return sc(ws, r, c, val, fill=FILL_AUTO)

def header_cell(ws, r, c, val=None):
    return sc(ws, r, c, val, font=F_HEADER, fill=FILL_HEADER)

def merge_and_set(ws, r1, c1, r2, c2, val=None, font=F_NORMAL, al=AL_C, border=THIN, fill=None):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    sc(ws, r1, c1, val, font, al, border, fill)
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            ws.cell(row=r, column=c).border = border

# ============================================================
# SHEET: 入力シート
# ============================================================
ws = wb.active
ws.title = '入力シート'
ws.sheet_properties.tabColor = 'FF0000'

for c in range(1, 8):
    ws.column_dimensions[get_column_letter(c)].width = 28

merge_and_set(ws, 1, 1, 1, 7, '消防設備工事 統合入力シート', F_TITLE, AL_C)
sc(ws, 2, 1, '※ 黄色セルに入力 → 各書式に自動反映されます', F_SMALL, AL_L)

# --- 基本情報 ---
r = 4
merge_and_set(ws, r, 1, r, 4, '【基本情報】', F_SECTION, AL_L, fill=FILL_SECTION)

fields = [
    # (label, cell_name_for_ref, default, row)
    ('担当消防署（例：○○消防署長　殿）', None, None),
    ('届出日', None, None),
    ('届出者 住所', None, None),
    ('届出者 氏名（管理組合名等）', None, None),
    ('工事場所（住所）', None, None),
    ('防火対象物の名称', None, None),
    ('設備の種類', None, '自動火災報知設備'),
    ('用途（例：5項ロ）', None, None),
    ('構造', None, '耐火構造'),
    ('地上階数', None, None),
    ('地下階数', None, None),
    ('延べ面積（㎡）', None, None),
    ('着工予定日', None, None),
    ('完成予定日', None, None),
    ('理事長名（設置届用）', None, None),
    ('全住戸数', None, None),
]

for i, (label, _, default) in enumerate(fields):
    rr = 5 + i
    header_cell(ws, rr, 1, label)
    input_cell(ws, rr, 2, default)

# Dropdowns
dv = DataValidation(type="list", formula1='"自動火災報知設備,共同住宅用自動火災報知設備,住戸用自動火災報知設備"', allow_blank=True)
ws.add_data_validation(dv)
dv.add(ws['B11'])

dv2 = DataValidation(type="list", formula1='"耐火構造,準耐火構造,その他"', allow_blank=True)
ws.add_data_validation(dv2)
dv2.add(ws['B13'])

# --- 施工者情報 ---
r = 22
merge_and_set(ws, r, 1, r, 4, '【施工者情報】', F_SECTION, AL_L, fill=FILL_SECTION)
for i, (label, default) in enumerate([
    ('施工会社名', '株式会社　TSCアクセス・プロ'),
    ('代表者名', '代表取締役　中村　勇'),
    ('施工会社住所', '名古屋市中区栄４丁目１６番３号　山光堂ビル５階'),
    ('電話番号', '052-269-9100'),
]):
    header_cell(ws, 23+i, 1, label)
    input_cell(ws, 23+i, 2, default)

# --- 消防設備士マスタ ---
r = 28
merge_and_set(ws, r, 1, r, 7, '【消防設備士マスタ】', F_SECTION, AL_L, fill=FILL_SECTION)
mst_headers = ['氏名','住所','免状種類','交付年月日','交付番号','講習年月','交付知事']
for i, h in enumerate(mst_headers):
    header_cell(ws, 29, 1+i, h)

staff = [
    ['河合　佑樹','愛知県尾張旭市晴丘町東236番地4','甲4','R5年7月21日','第00324号','','愛知'],
    ['戸田　健仁','愛知県長久手市西原山1-1 ｾﾝﾄｱｰｽ502','甲4','R5年11月7日','第00448号','','愛知'],
    ['棚町　征','岐阜県土岐市泉町大富195番地の249','甲4','H14年10月30日','第00009号','令和元年10月','愛知'],
]
for i, s in enumerate(staff):
    for j, v in enumerate(s):
        sc(ws, 30+i, 1+j, v, border=THIN)

# 担当者選択
r = 34
header_cell(ws, r, 1, '着工届 届出者（選択）')
input_cell(ws, r, 2, '戸田　健仁')
header_cell(ws, r+1, 1, '試験実施者（選択）')
input_cell(ws, r+1, 2)
header_cell(ws, r+2, 1, '配線試験実施者（選択）')
input_cell(ws, r+2, 2, '戸田　健仁')

for rr in range(34, 37):
    dv_s = DataValidation(type="list", formula1='"河合　佑樹,戸田　健仁,棚町　征"', allow_blank=True)
    ws.add_data_validation(dv_s)
    dv_s.add(ws.cell(row=rr, column=2))

# --- 受信機情報 ---
r = 38
merge_and_set(ws, r, 1, r, 4, '【受信機・機器情報（共用部）】', F_SECTION, AL_L, fill=FILL_SECTION)
for i, (label, default) in enumerate([
    ('受信機 型級（P・GP型 ○級）', '1'),
    ('回線数（例：10/10）', '10/10'),
    ('予備電源 DC（V）', '24'),
    ('予備電源（AH）', '0.45'),
    ('受信機設置場所', '管理人室'),
    ('受信機メーカー', 'パナソニック㈱'),
    ('型式番号（受第○～○号）', ''),
]):
    header_cell(ws, 39+i, 1, label)
    input_cell(ws, 39+i, 2, default)

r = 47
merge_and_set(ws, r, 1, r, 4, '【受信機・機器情報（住戸部）】', F_SECTION, AL_L, fill=FILL_SECTION)
for i, (label, default) in enumerate([
    ('受信機 型級（P・GP型 ○級）', '3'),
    ('回線数（例：1/1）', '1/1'),
    ('受信機設置場所', '各住戸リビング'),
    ('受信機メーカー', 'パナソニック㈱'),
    ('型式番号（受第○～○号）', ''),
    ('表示器台数', ''),
]):
    header_cell(ws, 48+i, 1, label)
    input_cell(ws, 48+i, 2, default)

# --- 音響・発信機 ---
r = 55
merge_and_set(ws, r, 1, r, 4, '【地区音響装置・発信機・表示灯】', F_SECTION, AL_L, fill=FILL_SECTION)
for i, (label, default) in enumerate([
    ('地区音響 認評音第（号）', '13-5'),
    ('地区音響 メーカー', 'パナソニック㈱'),
    ('地区音響 電圧DC（V）', '24'),
    ('地区音響 電流（mA）', '10'),
    ('地区音響 個数', ''),
    ('地区音響 音量（dB）', '95'),
    ('地区音響 鐘径（mm）', '150'),
    ('発信機 型（P等）', 'P'),
    ('発信機 級', '1'),
    ('発信機 個数（屋外型）', ''),
    ('発信機 型式番号（発第）', ''),
    ('発信機 メーカー', 'パナソニック㈱'),
    ('表示灯 電圧DC（V）', '24'),
    ('表示灯 電流（mA）', '9'),
    ('表示灯 個数', ''),
]):
    header_cell(ws, 56+i, 1, label)
    input_cell(ws, 56+i, 2, default)

# --- 配線試験値 ---
r = 72
merge_and_set(ws, r, 1, r, 4, '【配線試験 絶縁抵抗値】', F_SECTION, AL_L, fill=FILL_SECTION)
for i, (label, default) in enumerate([
    ('電源回路 試験電圧（V）', '250'),
    ('電源回路 絶縁抵抗（MΩ）', '150'),
    ('操作回路 試験電圧（V）', '50'),
    ('操作回路 絶縁抵抗（MΩ）', '20'),
    ('表示灯回路 試験電圧（V）', '50'),
    ('表示灯回路 絶縁抵抗（MΩ）', '20'),
    ('警報回路 試験電圧（V）', '50'),
    ('警報回路 絶縁抵抗（MΩ）', '20'),
    ('感知器回路 試験電圧（V）', '50'),
    ('感知器回路 絶縁抵抗（MΩ）', '20'),
]):
    header_cell(ws, 73+i, 1, label)
    input_cell(ws, 73+i, 2, default)

# --- 工事種別 ---
r = 84
merge_and_set(ws, r, 1, r, 4, '【工事関連】', F_SECTION, AL_L, fill=FILL_SECTION)
header_cell(ws, 85, 1, '工事種別')
input_cell(ws, 85, 2, '取替え')
dv_w = DataValidation(type="list", formula1='"新設,増設,移設,取替え,改造,その他"', allow_blank=True)
ws.add_data_validation(dv_w)
dv_w.add(ws['B85'])

header_cell(ws, 86, 1, '試験実施日')
input_cell(ws, 86, 2)
header_cell(ws, 87, 1, '完成年月日')
input_cell(ws, 87, 2)

header_cell(ws, 88, 1, '関連設備')
input_cell(ws, 88, 2, 'インターホン設備')

header_cell(ws, 89, 1, '鳴動方式')
input_cell(ws, 89, 2, '一斉鳴動')
dv_ring = DataValidation(type="list", formula1='"一斉鳴動,区分鳴動"', allow_blank=True)
ws.add_data_validation(dv_ring)
dv_ring.add(ws['B89'])

header_cell(ws, 90, 1, '放送設備連動')
input_cell(ws, 90, 2, '無')
dv_yn = DataValidation(type="list", formula1='"有,無"', allow_blank=True)
ws.add_data_validation(dv_yn)
dv_yn.add(ws['B90'])

header_cell(ws, 91, 1, '中継器設置場所')
input_cell(ws, 91, 2, '戸外表示機組込')

header_cell(ws, 92, 1, '受信機操作部 高さ（m）')
input_cell(ws, 92, 2, '1.4')

# ============================================================
# Helper function: create a form sheet by converting Word table structure
# ============================================================

def create_着工届(wb):
    """別記様式第1号の7 着工届出書"""
    ws = wb.create_sheet('着工届出書')
    ws.sheet_properties.tabColor = '4472C4'
    
    # Column widths to match original
    widths = [2.5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i+1)].width = w
    
    sc(ws, 1, 1, '別記様式第１号の７（第３３条の１８関係）', F_SMALL, AL_L, border=None)
    
    merge_and_set(ws, 3, 1, 3, 18, '工事整備対象設備等着工届出書', F_TITLE, AL_C, border=None)
    
    # Date
    merge_and_set(ws, 5, 1, 5, 9, None, F_NORMAL, AL_C, border=None)
    auto_cell(ws, 5, 1).value = '=入力シート!B6'
    
    # 殿
    merge_and_set(ws, 7, 1, 7, 9, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 7, 1).value = '=入力シート!B5'
    
    # 届出者
    sc(ws, 9, 8, '届 出 者', F_HEADER, AL_R, border=None)
    
    sc(ws, 10, 8, '住 所', F_NORMAL, AL_R, border=None)
    merge_and_set(ws, 10, 9, 10, 18, None, F_NORMAL, AL_L, border=None)
    # Use VLOOKUP to get address from staff master based on selected name
    auto_cell(ws, 10, 9).value = '=VLOOKUP(入力シート!B34,入力シート!A30:G32,2,FALSE)'
    
    sc(ws, 11, 8, '氏 名', F_NORMAL, AL_R, border=None)
    merge_and_set(ws, 11, 9, 11, 18, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 11, 9).value = '=入力シート!B34'
    
    # 工事場所等 with borders
    r = 13
    merge_and_set(ws, r, 1, r, 7, '工事の場所', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 8, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 8).value = '=入力シート!B9'
    
    r = 14
    merge_and_set(ws, r, 1, r, 7, '工事を行う防火対象物の名称', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 8, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 8).value = '=入力シート!B10'
    
    r = 15
    merge_and_set(ws, r, 1, r, 7, '工事整備対象設備等の種類', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 8, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 8).value = '=入力シート!B11'
    
    # 工事施工者
    r = 16
    merge_and_set(ws, r, 1, r+1, 2, '工事整備対象設備\n等の工事施工者', F_SMALL, AL_C, THIN)
    sc(ws, r, 3, '住所', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 4, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 4).value = '=入力シート!B25'
    
    r = 17
    sc(ws, r, 3, '氏名', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 4, r, 9, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 4).value = '=入力シート!B23&"  "&入力シート!B24'
    merge_and_set(ws, r, 10, r, 18, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r, 10).value = '="電話　"&入力シート!B26'
    
    # 消防設備士
    r = 18
    merge_and_set(ws, r, 1, r+3, 2, '消防設備士', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 3, r+1, 4, '免状の\n種類及び指定区分', F_SMALL, AL_C, THIN)
    sc(ws, r, 5, '種類等', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 6, r, 7, '交付知事', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 8, r, 12, '交付年月日', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 13, r, 18, '講習受講状況', F_SMALL, AL_C, THIN)
    
    r = 19
    sc(ws, r, 5, None)
    merge_and_set(ws, r, 6, r, 7, None, F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 8, r, 12, '交付番号', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 13, r, 15, '受講地', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 16, r, 18, '受講年月', F_SMALL, AL_C, THIN)
    
    r = 20
    sc(ws, r, 3, '甲・乙', F_SMALL, AL_C, THIN)
    sc(ws, r, 4, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r, 4).value = '="種4類"'
    sc(ws, r, 5, None, F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 6, r, 7, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r, 6).value = '=VLOOKUP(入力シート!B34,入力シート!A30:G32,7,FALSE)'
    merge_and_set(ws, r, 8, r, 12, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r, 8).value = '=VLOOKUP(入力シート!B34,入力シート!A30:G32,4,FALSE)'
    
    r = 21
    sc(ws, r, 3, '甲・乙', F_SMALL, AL_C, THIN)
    sc(ws, r, 4, None, F_SMALL, AL_C, THIN)
    sc(ws, r, 5, None, F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 6, r, 7, None, F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 8, r, 12, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r, 8).value = '=VLOOKUP(入力シート!B34,入力シート!A30:G32,5,FALSE)'
    merge_and_set(ws, r, 13, r, 15, None, F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 16, r, 18, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r, 16).value = '=VLOOKUP(入力シート!B34,入力シート!A30:G32,6,FALSE)'
    
    # 工事種別
    r = 22
    merge_and_set(ws, r, 1, r, 7, '工事の種別', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 8, r, 18, '１ 新設　２ 増設　３ 移設　４ 取替え　５ 改造　６ その他', F_NORMAL, AL_L, THIN)
    
    # 着工・完成予定日
    r = 23
    merge_and_set(ws, r, 1, r, 7, '着工予定日', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 8, r, 12, None, F_NORMAL, AL_C, THIN)
    auto_cell(ws, r, 8).value = '=入力シート!B17'
    merge_and_set(ws, r, 13, r, 15, '完成予定日', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 16, r, 18, None, F_NORMAL, AL_C, THIN)
    auto_cell(ws, r, 16).value = '=入力シート!B18'
    
    # 受付欄・経過欄
    r = 24
    merge_and_set(ws, r, 1, r+3, 9, '※ 受　付　欄', F_NORMAL, AL_TL, THIN)
    merge_and_set(ws, r, 10, r+3, 18, '※ 経　過　欄', F_NORMAL, AL_TL, THIN)
    
    # 備考
    r = 28
    sc(ws, r, 1, '備考', F_SMALL, AL_L, border=None)
    sc(ws, r+1, 2, '１ この用紙の大きさは、日本産業規格Ａ４とすること。', F_SMALL, AL_L, border=None)
    sc(ws, r+2, 2, '２ 工事の種別の欄は、該当する事項を○印で囲むこと。', F_SMALL, AL_L, border=None)
    sc(ws, r+3, 2, '３ ※印の欄には、記入しないこと。', F_SMALL, AL_L, border=None)
    
    ws.print_area = 'A1:R31'
    return ws

def create_設置届(wb):
    """別記様式第1号の2の3 設置届出書"""
    ws = wb.create_sheet('設置届出書')
    ws.sheet_properties.tabColor = '548235'
    
    widths = [2.5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i+1)].width = w
    
    sc(ws, 1, 1, '別記様式第１号の２の３（第３１条の３関係）', F_SMALL, AL_L, border=None)
    merge_and_set(ws, 3, 1, 3, 18, '消防用設備等（特殊消防用設備等）設置届出書', F_TITLE, AL_C, border=None)
    
    # Date (input)
    merge_and_set(ws, 5, 12, 5, 18, None, F_NORMAL, AL_L, border=None)
    input_cell(ws, 5, 12).value = '令和　　年　　月　　日'
    
    # 殿
    merge_and_set(ws, 7, 1, 7, 9, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 7, 1).value = '=入力シート!B5'
    
    sc(ws, 9, 8, '届出者', F_HEADER, AL_R, border=None)
    sc(ws, 10, 8, '住 所', F_NORMAL, AL_R, border=None)
    merge_and_set(ws, 10, 9, 10, 18, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 10, 9).value = '=入力シート!B7'
    
    sc(ws, 11, 8, '氏 名', F_NORMAL, AL_R, border=None)
    merge_and_set(ws, 11, 9, 11, 18, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 11, 9).value = '=入力シート!B8&"  "&入力シート!B19'
    
    sc(ws, 13, 1, '記', F_HEADER, AL_C, border=None)
    
    # 設置者
    r = 14
    merge_and_set(ws, r, 1, r+1, 4, '設置者', F_NORMAL, AL_C, THIN)
    sc(ws, r, 5, '住所', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 6, r, 18, '届出者に同じ', F_NORMAL, AL_L, THIN)
    sc(ws, r+1, 5, '氏名', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r+1, 6, r+1, 18, '届出者に同じ', F_NORMAL, AL_L, THIN)
    
    # 防火対象物
    r = 16
    merge_and_set(ws, r, 1, r+3, 4, '防火対象物', F_NORMAL, AL_C, THIN)
    sc(ws, r, 5, '所在地', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 6, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 6).value = '=入力シート!B9'
    
    sc(ws, r+1, 5, '名称', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r+1, 6, r+1, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+1, 6).value = '=入力シート!B10'
    
    sc(ws, r+2, 5, '用途', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r+2, 6, r+2, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+2, 6).value = '="令別表第一（"&入力シート!B12&"）  "&入力シート!B13'
    
    sc(ws, r+3, 5, '構造・規模', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r+3, 6, r+3, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+3, 6).value = '=入力シート!B13&"  地上 "&入力シート!B14&"階  延べ面積 "&入力シート!B16&"㎡"'
    
    # 消防用設備等の種類
    r = 20
    merge_and_set(ws, r, 1, r, 4, '消防用設備等の種類', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 5, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 5).value = '=入力シート!B11'
    
    # 工事
    r = 21
    merge_and_set(ws, r, 1, r+5, 4, '工　事', F_NORMAL, AL_C, THIN)
    sc(ws, r, 5, '種別', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 6, r, 18, '新設、増設、移設、取替え、改造、その他（　）', F_NORMAL, AL_L, THIN)
    
    merge_and_set(ws, r+1, 5, r+2, 5, '施工者', F_SMALL, AL_C, THIN)
    sc(ws, r+1, 6, '住所', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r+1, 7, r+1, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+1, 7).value = '=入力シート!B25'
    sc(ws, r+2, 6, '氏名', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r+2, 7, r+2, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+2, 7).value = '=入力シート!B23&"　"&入力シート!B24'
    
    merge_and_set(ws, r+3, 5, r+5, 5, '消防\n設備士', F_SMALL, AL_C, THIN)
    sc(ws, r+3, 6, '住所', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r+3, 7, r+3, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+3, 7).value = '=VLOOKUP(入力シート!B34,入力シート!A30:G32,2,FALSE)'
    sc(ws, r+4, 6, '氏名', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r+4, 7, r+4, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+4, 7).value = '=入力シート!B34'
    
    # 免状
    sc(ws, r+5, 6, '免状', F_NORMAL, AL_C, THIN)
    sc(ws, r+5, 7, '甲・乙', F_SMALL, AL_C, THIN)
    sc(ws, r+5, 8, '種4類', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+5, 9, r+5, 11, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r+5, 9).value = '=VLOOKUP(入力シート!B34,入力シート!A30:G32,7,FALSE)&"　"&VLOOKUP(入力シート!B34,入力シート!A30:G32,4,FALSE)'
    merge_and_set(ws, r+5, 12, r+5, 14, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r+5, 12).value = '=VLOOKUP(入力シート!B34,入力シート!A30:G32,5,FALSE)'
    merge_and_set(ws, r+5, 15, r+5, 18, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r+5, 15).value = '=VLOOKUP(入力シート!B34,入力シート!A30:G32,6,FALSE)'
    
    # 完成年月日
    r = 27
    merge_and_set(ws, r, 1, r, 4, '完成年月日', F_NORMAL, AL_C, THIN)
    merge_and_set(ws, r, 5, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 5).value = '=入力シート!B87'
    
    # 受付欄等
    r = 28
    merge_and_set(ws, r, 1, r+2, 6, '※ 受付欄', F_NORMAL, AL_TL, THIN)
    merge_and_set(ws, r, 7, r+2, 12, '※ 決裁欄', F_NORMAL, AL_TL, THIN)
    merge_and_set(ws, r, 13, r+2, 18, '※ 備考', F_NORMAL, AL_TL, THIN)
    
    # 備考
    r = 31
    sc(ws, r, 1, '備考　１　この用紙の大きさは、日本産業規格Ａ４とすること。', F_SMALL, AL_L, border=None)
    sc(ws, r+1, 1, '　　　２　消防用設備等設計図書は、種類ごとにそれぞれ添付すること。', F_SMALL, AL_L, border=None)
    sc(ws, r+2, 1, '　　　３　※欄には、記入しないこと。', F_SMALL, AL_L, border=None)
    
    return ws

def create_概要表その1(wb):
    """別記様式第5 自動火災報知設備の概要表（その1）- 感知器"""
    ws = wb.create_sheet('概要表（その1）')
    ws.sheet_properties.tabColor = 'ED7D31'
    
    widths = [3, 8, 8, 3, 3, 3, 3, 8, 3, 3, 3, 3, 5]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i+1)].width = w
    
    sc(ws, 1, 1, '別記様式５', F_SMALL, AL_L, border=None)
    merge_and_set(ws, 2, 1, 2, 13, '自 動 火 災 報 知 設 備 の 概 要 表　　　（その１）', F_HEADER, AL_C, border=None)
    
    # Headers
    r = 4
    merge_and_set(ws, r, 1, r+17, 1, '感\n知\n器', F_HEADER, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 8, '機　種', F_HEADER, AL_C, THIN)
    sc(ws, r, 9, '蓄積', F_SMALL, AL_C, THIN)
    sc(ws, r, 10, '自動', F_SMALL, AL_C, THIN)
    sc(ws, r, 11, '遠隔', F_SMALL, AL_C, THIN)
    sc(ws, r, 12, '種別', F_SMALL, AL_C, THIN)
    sc(ws, r, 13, '個数', F_SMALL, AL_C, THIN)
    
    # 10 sensor entry rows (2 rows each: type + serial)
    dv_type = DataValidation(type="list", formula1='"差動式,定温式,光電式,補償式,熱アナログ式,光電アナログ式,イオン化式"', allow_blank=True)
    ws.add_data_validation(dv_type)
    dv_grade = DataValidation(type="list", formula1='"特種,1種,2種,3種"', allow_blank=True)
    ws.add_data_validation(dv_grade)
    dv_yesno = DataValidation(type="list", formula1='"〇,"', allow_blank=True)
    ws.add_data_validation(dv_yesno)
    dv_maker = DataValidation(type="list", formula1='"パナソニック㈱,ニッタン(株),能美防災㈱,ホーチキ㈱"', allow_blank=True)
    ws.add_data_validation(dv_maker)
    
    for i in range(10):
        rr = 5 + i*2
        # Type row
        merge_and_set(ws, rr, 2, rr, 6, None, F_NORMAL, AL_L, THIN)
        input_cell(ws, rr, 2)
        dv_type.add(ws.cell(row=rr, column=2))
        sc(ws, rr, 7, None, F_NORMAL, AL_C, THIN)
        input_cell(ws, rr, 7)  # 型名詳細
        sc(ws, rr, 8, None, F_NORMAL, AL_C, THIN)
        input_cell(ws, rr, 8)  # メーカー
        dv_maker.add(ws.cell(row=rr, column=8))
        input_cell(ws, rr, 9); dv_yesno.add(ws.cell(row=rr, column=9))
        input_cell(ws, rr, 10); dv_yesno.add(ws.cell(row=rr, column=10))
        input_cell(ws, rr, 11); dv_yesno.add(ws.cell(row=rr, column=11))
        input_cell(ws, rr, 12); dv_grade.add(ws.cell(row=rr, column=12))
        input_cell(ws, rr, 13)  # 個数
        
        # Serial row
        rr2 = rr + 1
        merge_and_set(ws, rr2, 2, rr2, 4, '型式番号　感第', F_SMALL, AL_L, THIN)
        merge_and_set(ws, rr2, 5, rr2, 6, None, F_NORMAL, AL_C, THIN)
        input_cell(ws, rr2, 5)  # 型式番号
        sc(ws, rr2, 7, '号', F_SMALL, AL_L, THIN)
        sc(ws, rr2, 8, '製造会社名', F_SMALL, AL_C, THIN)
        merge_and_set(ws, rr2, 9, rr2, 13, None, F_NORMAL, AL_L, THIN)
        auto_cell(ws, rr2, 9).value = f'=H{rr}'
    
    # 発信機
    r = 25
    merge_and_set(ws, r, 1, r+1, 1, '発信機', F_SMALL, AL_C, THIN)
    sc(ws, r, 2, '屋内型', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 3, r, 6, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 3).value = '=入力シート!B63&"型 "&入力シート!B64&"級"'
    merge_and_set(ws, r, 7, r, 10, None, F_SMALL, AL_L, THIN)
    merge_and_set(ws, r, 11, r, 13, None, F_NORMAL, AL_L, THIN)
    
    sc(ws, r+1, 2, '屋外型', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+1, 3, r+1, 6, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+1, 3).value = '=入力シート!B63&"型 "&入力シート!B64&"級 "&入力シート!B65&"個"'
    merge_and_set(ws, r+1, 7, r+1, 10, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r+1, 7).value = '="型式番号　発第"&入力シート!B66&"号"'
    merge_and_set(ws, r+1, 11, r+1, 13, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+1, 11).value = '=入力シート!B67'
    
    # 表示灯
    r = 27
    sc(ws, r, 1, '表示灯', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 13, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 2).value = '="DC "&入力シート!B68&"V  "&入力シート!B69&"mA  "&入力シート!B70&"個"'
    
    # 中継器
    r = 28
    merge_and_set(ws, r, 1, r+10, 1, '中継器', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 6, '種　別', F_SMALL, AL_C, THIN)
    sc(ws, r, 7, '回線数', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 8, r, 12, '電源供給方式', F_SMALL, AL_C, THIN)
    sc(ws, r, 13, '設置台数', F_SMALL, AL_C, THIN)
    
    for i in range(10):
        rr = 29 + i
        merge_and_set(ws, rr, 2, rr, 6, '自動･遠隔･アナログ･その他（型式番号　中第　　号）', F_SMALL, AL_L, THIN)
        input_cell(ws, rr, 2)
        input_cell(ws, rr, 7)
        merge_and_set(ws, rr, 8, rr, 12, '専用（予備電源　Ｖ　AH）･受信機･その他', F_SMALL, AL_L, THIN)
        input_cell(ws, rr, 13)
    
    # 備考
    r = 39
    sc(ws, r, 1, '備考　１　この用紙の大きさは、日本産業規格Ａ４とすること。', F_SMALL, AL_L, border=None)
    sc(ws, r+1, 1, '　　　２　選択肢の併記してある欄は、該当事項を○印で囲むこと。', F_SMALL, AL_L, border=None)
    
    return ws

def create_概要表その2(wb):
    """別記様式第5 自動火災報知設備の概要表（その2）- 受信機・音響等"""
    ws = wb.create_sheet('概要表（その2）')
    ws.sheet_properties.tabColor = 'ED7D31'
    
    widths = [3, 6, 6, 3, 5, 5, 5, 5, 5, 5, 5, 5, 5]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i+1)].width = w
    
    sc(ws, 1, 1, '別記様式５', F_SMALL, AL_L, border=None)
    merge_and_set(ws, 2, 1, 2, 13, '自 動 火 災 報 知 設 備 の 概 要 表　　　（その２）', F_HEADER, AL_C, border=None)
    
    # 受信機
    r = 4
    merge_and_set(ws, r, 1, r+3, 1, '受信機', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 13, '蓄積式・二信号式・アナログ式・自動試験機能付き・遠隔試験機能付き・その他（　　）', F_SMALL, AL_L, THIN)
    
    merge_and_set(ws, r+1, 2, r+1, 13, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+1, 2).value = '="Ｐ・GP型"&入力シート!B39&"級  "&入力シート!B40&"回線  Ｒ・GR型  自火報点数　点  予備点数　点"'
    
    merge_and_set(ws, r+2, 2, r+2, 5, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+2, 2).value = '="予備電源（DC "&入力シート!B41&"V  "&入力シート!B42&"AH）"'
    sc(ws, r+2, 6, '設置場所', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 7, r+2, 13, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+2, 7).value = '=入力シート!B43'
    
    merge_and_set(ws, r+3, 2, r+3, 5, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+3, 2).value = '="型式番号　受第"&入力シート!B45&"号"'
    sc(ws, r+3, 6, '製造会社名', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+3, 7, r+3, 13, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+3, 7).value = '=入力シート!B44'
    
    # 表示器
    r = 8
    merge_and_set(ws, r, 1, r+1, 1, '表示器', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 13, '　／　　回線　　台　自火報点数　点　その他点数　点　予備点数　点', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+1, 2, r+1, 13, '　／　　回線　　台　自火報点数　点　その他点数　点　予備点数　点', F_SMALL, AL_L, THIN)
    
    # 電源
    r = 10
    merge_and_set(ws, r, 1, r+3, 1, '電源', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r+1, 2, '常用電源', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 3, r, 13, '単相・三相　AC　100Ｖ　非常電源専用受電設備回路・電灯回路・動力回路', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+1, 3, r+1, 13, 'DC　　Ｖ　　AH　充電方式（トリクル・浮動）　使用別（専用・共用）', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+2, 2, r+3, 2, '非常電源', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 3, r+2, 13, '非常電源専用受電設備　単相・三相　AC　　Ｖ', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+3, 3, r+3, 13, '蓄電池設備　DC　Ｖ　AH　充電方式（トリクル・浮動）', F_SMALL, AL_L, THIN)
    
    # 音響装置
    r = 14
    merge_and_set(ws, r, 1, r+5, 1, '音響装置', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r+2, 3, '主音響装置\n（内蔵されている\nものは除く）', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 4, r, 13, 'ベル・サイレン・電子ブザー・音声合成・その他（　　）', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+1, 4, r+1, 6, None, F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+1, 7, r+1, 13, None, F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+2, 4, r+2, 6, None, F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+2, 7, r+2, 13, None, F_SMALL, AL_L, THIN)
    
    merge_and_set(ws, r+3, 2, r+3, 3, '地区音響装置', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+3, 4, r+3, 6, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r+3, 4).value = '="型式番号（認評音第"&入力シート!B56&"号）"'
    sc(ws, r+3, 7, '製造会社名', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+3, 8, r+3, 13, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+3, 8).value = '=入力シート!B57'
    
    merge_and_set(ws, r+4, 2, r+4, 3, '', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+4, 4, r+4, 13, 'ベル・サイレン・電子ブザー・スピーカー・その他', F_SMALL, AL_L, THIN)
    
    merge_and_set(ws, r+5, 2, r+5, 3, '', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+5, 4, r+5, 6, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r+5, 4).value = '="鐘径 "&入力シート!B62&"mm"'
    merge_and_set(ws, r+5, 7, r+5, 13, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r+5, 7).value = '="定格DC "&入力シート!B58&"V  "&入力シート!B59&"mA  "&入力シート!B60&"個  "&入力シート!B61&"dB"'
    
    # 音声切替装置
    r = 20
    merge_and_set(ws, r, 1, r+2, 1, '', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 3, '音声切替装置', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 4, r, 7, '型式番号（　　号）DC　Ｖ', F_SMALL, AL_L, THIN)
    sc(ws, r, 8, '製造会社名', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 9, r, 13, None, F_SMALL, AL_L, THIN)
    
    merge_and_set(ws, r+1, 2, r+1, 3, '', F_SMALL, AL_C, THIN)
    sc(ws, r+1, 4, '常用電源', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+1, 5, r+1, 13, '単相　AC　Ｖ　非常電源専用受電設備回路・電灯回路', F_SMALL, AL_L, THIN)
    
    merge_and_set(ws, r+2, 2, r+2, 3, '', F_SMALL, AL_C, THIN)
    sc(ws, r+2, 4, '非常電源', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 5, r+2, 13, '蓄電池設備　DC　Ｖ　AH　充電方式（トリクル・浮動）', F_SMALL, AL_L, THIN)
    
    # 配線
    r = 23
    merge_and_set(ws, r, 1, r+3, 1, '配線', F_SMALL, AL_C, THIN)
    sc(ws, r, 2, '常用電源回路', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 3, r, 13, 'ケーブル露出・電線管露出・電線管埋設・その他', F_SMALL, AL_L, THIN)
    sc(ws, r+1, 2, '非常電源回路', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+1, 3, r+1, 13, '耐火電線・電線管露出・電線管埋設・その他', F_SMALL, AL_L, THIN)
    sc(ws, r+2, 2, '警報回路', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 3, r+2, 13, '耐火電線・電線管露出・電線管埋設・その他', F_SMALL, AL_L, THIN)
    sc(ws, r+3, 2, 'その他回路', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+3, 3, r+3, 13, 'IV電線・ケーブル露出・電線管露出・電線管埋設・その他', F_SMALL, AL_L, THIN)
    
    # 関連設備
    r = 27
    merge_and_set(ws, r, 1, r+1, 1, '関連設備', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 13, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r, 2).value = '=入力シート!B88'
    merge_and_set(ws, r+1, 2, r+1, 13, None, F_SMALL, AL_L, THIN)
    
    # 工事者区分
    r = 29
    merge_and_set(ws, r, 1, r+4, 1, '工事者区分', F_SMALL, AL_C, THIN)
    sc(ws, r, 2, '電源工事', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 3, r, 13, None, F_SMALL, AL_L, THIN)
    sc(ws, r+1, 2, '配線工事', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+1, 3, r+1, 13, None, F_SMALL, AL_L, THIN)
    sc(ws, r+2, 2, '配線工事', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 3, r+2, 13, None, F_SMALL, AL_L, THIN)
    sc(ws, r+3, 2, '配線工事', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+3, 3, r+3, 13, None, F_SMALL, AL_L, THIN)
    sc(ws, r+4, 2, '機器の取付工事', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+4, 3, r+4, 13, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r+4, 3).value = '=入力シート!B23'
    
    # その他
    r = 34
    sc(ws, r, 1, 'その他', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r+2, 13, None, F_SMALL, AL_L, THIN)
    
    r = 37
    sc(ws, r, 1, '備考　１　この用紙の大きさは、日本産業規格Ａ４とすること。', F_SMALL, AL_L, border=None)
    sc(ws, r+1, 1, '　　　２　選択肢の併記してある欄は、該当事項を○印で囲むこと。', F_SMALL, AL_L, border=None)
    sc(ws, r+2, 1, '　　　３　感知器記入欄の（　）内は、その機能又は性能を記入すること。', F_SMALL, AL_L, border=None)
    sc(ws, r+3, 1, '　　　４　関連設備の消火設備（　）内は、その設備等の種類を記入すること。', F_SMALL, AL_L, border=None)
    
    return ws

def create_試験結果11(wb):
    """別記様式第11 自動火災報知設備 試験結果報告書（その1）"""
    ws = wb.create_sheet('試験結果(様式11)')
    ws.sheet_properties.tabColor = '7030A0'
    
    widths = [3, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 3, 3, 5, 5, 5, 5, 5]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i+1)].width = w
    
    sc(ws, 1, 1, '別記様式第11', F_SMALL, AL_L, border=None)
    sc(ws, 1, 18, '（その１）①', F_SMALL, AL_R, border=None)
    merge_and_set(ws, 3, 1, 3, 18, '自 動 火 災 報 知 設 備 試 験 結 果 報 告 書', F_HEADER, AL_C, border=None)
    
    merge_and_set(ws, 5, 1, 5, 18, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 5, 1).value = '="試験実施日　"&入力シート!B86'
    
    sc(ws, 7, 8, '試験実施者', F_HEADER, AL_R, border=None)
    sc(ws, 9, 8, '住　所', F_NORMAL, AL_R, border=None)
    merge_and_set(ws, 9, 9, 9, 18, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 9, 9).value = '=VLOOKUP(入力シート!B35,入力シート!A30:G32,2,FALSE)'
    sc(ws, 10, 8, '氏　名', F_NORMAL, AL_R, border=None)
    merge_and_set(ws, 10, 9, 10, 18, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 10, 9).value = '=入力シート!B35'
    
    # 基本情報
    r = 12
    sc(ws, r, 1, '用途', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 2).value = '="（"&入力シート!B12&"）項"'
    
    r = 13
    sc(ws, r, 1, '延べ面積', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 8, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 2).value = '=入力シート!B16&"㎡"'
    sc(ws, r, 9, '階数', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 10, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 10).value = '="地上 "&入力シート!B14&"階"'
    
    # 受信機
    r = 14
    merge_and_set(ws, r, 1, r+3, 2, '受信機', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 3, r, 18, '蓄積式・二信号式・アナログ式・自動試験機能付き・遠隔試験機能付き・無線式・その他', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+1, 3, r+1, 9, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r+1, 3).value = '="Ｐ・ＧＰ型"&入力シート!B39&"級　回線数  "&入力シート!B40'
    merge_and_set(ws, r+1, 10, r+1, 18, 'Ｒ・GR型 自火報点数　点・その他点数　点・予備点数　点', F_SMALL, AL_L, THIN)
    
    merge_and_set(ws, r+2, 3, r+2, 4, '定格電圧', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 5, r+2, 18, 'ＡＣ　100　Ｖ　・　ＤＣ　　Ｖ', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+3, 3, r+3, 4, '予備電源', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+3, 5, r+3, 18, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r+3, 5).value = '="NiCd・その他（　）DC "&入力シート!B41&"V  "&入力シート!B42&"AH"'
    
    # 発信機
    r = 18
    sc(ws, r, 1, '発信機', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 18, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r, 2).value = '=入力シート!B63&"型  "&入力シート!B64&"級  屋外型  "&入力シート!B65&"個"'
    
    # 中継器
    r = 19
    merge_and_set(ws, r, 1, r+3, 2, '中継器', F_SMALL, AL_C, THIN)
    for i in range(4):
        rr = r + i
        merge_and_set(ws, rr, 3, rr, 9, 'アナログ式・蓄積式・自動試験機能付き・遠隔試験機能付き・無線式・他（　）回線', F_SMALL, AL_L, THIN)
        merge_and_set(ws, rr, 10, rr, 15, '予備電源　有（V　AH）・無', F_SMALL, AL_L, THIN)
        merge_and_set(ws, rr, 16, rr, 18, '設置台数　台', F_SMALL, AL_L, THIN)
    
    # 感知器
    r = 23
    merge_and_set(ws, r, 1, r+10, 2, '感知器', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 3, r, 9, '機　種', F_SMALL, AL_C, THIN)
    sc(ws, r, 10, '自', F_SMALL, AL_C, THIN)
    sc(ws, r, 11, '遠', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 12, r, 15, '種別', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 16, r, 18, '個数', F_SMALL, AL_C, THIN)
    
    for i in range(10):
        rr = 24 + i
        merge_and_set(ws, rr, 3, rr, 9, None, F_SMALL, AL_L, THIN)
        input_cell(ws, rr, 3)
        input_cell(ws, rr, 10)
        input_cell(ws, rr, 11)
        merge_and_set(ws, rr, 12, rr, 15, None, F_SMALL, AL_C, THIN)
        input_cell(ws, rr, 12)
        merge_and_set(ws, rr, 16, rr, 18, None, F_SMALL, AL_C, THIN)
        input_cell(ws, rr, 16)
    
    # 音響装置
    r = 34
    merge_and_set(ws, r, 1, r+7, 2, '音響装置', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 3, r, 5, '種別', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 6, r, 9, '種類', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 10, r, 13, '電圧', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 14, r, 16, '電流', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 17, r, 18, '個数', F_SMALL, AL_C, THIN)
    
    labels = ['主音響装置', '副音響装置', '地区音響装置', '', '']
    for i, label in enumerate(labels):
        rr = 35 + i
        merge_and_set(ws, rr, 3, rr, 5, label, F_SMALL, AL_C, THIN)
        merge_and_set(ws, rr, 6, rr, 9, None, F_SMALL, AL_C, THIN)
        input_cell(ws, rr, 6)
        merge_and_set(ws, rr, 10, rr, 13, None, F_SMALL, AL_C, THIN)
        input_cell(ws, rr, 10)
        merge_and_set(ws, rr, 14, rr, 16, None, F_SMALL, AL_C, THIN)
        input_cell(ws, rr, 14)
        merge_and_set(ws, rr, 17, rr, 18, None, F_SMALL, AL_C, THIN)
        input_cell(ws, rr, 17)
    
    # Auto-fill first 地区音響 row from input sheet
    auto_cell(ws, 37, 10).value = '="DC "&入力シート!B58&"V"'
    auto_cell(ws, 37, 14).value = '=入力シート!B59&"mA"'
    auto_cell(ws, 37, 17).value = '=入力シート!B60'
    
    rr = 40
    merge_and_set(ws, rr, 3, rr, 5, '放送設備との連動', F_SMALL, AL_C, THIN)
    merge_and_set(ws, rr, 6, rr, 18, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, rr, 6).value = '=入力シート!B90'
    
    rr = 41
    merge_and_set(ws, rr, 3, rr, 5, '鳴動方式', F_SMALL, AL_C, THIN)
    merge_and_set(ws, rr, 6, rr, 18, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, rr, 6).value = '=入力シート!B89'
    
    return ws

def create_試験結果34(wb):
    """別記様式第34 住戸用自火報・共同住宅用非常警報設備 試験結果報告書"""
    ws = wb.create_sheet('試験結果(様式34)')
    ws.sheet_properties.tabColor = 'BF8F00'
    
    widths = [3, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 3, 3, 5, 5, 5, 5, 5]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i+1)].width = w
    
    sc(ws, 1, 1, '別記様式第34', F_SMALL, AL_L, border=None)
    sc(ws, 1, 18, '①', F_SMALL, AL_R, border=None)
    merge_and_set(ws, 3, 1, 3, 18, '住戸用自動火災報知設備・共同住宅用非常警報設備試験結果報告書', F_HEADER, AL_C, border=None)
    
    merge_and_set(ws, 5, 1, 5, 18, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 5, 1).value = '="試験実施日　"&入力シート!B86'
    
    sc(ws, 7, 8, '試験実施者', F_HEADER, AL_R, border=None)
    sc(ws, 9, 8, '住　所', F_NORMAL, AL_R, border=None)
    merge_and_set(ws, 9, 9, 9, 18, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 9, 9).value = '=VLOOKUP(入力シート!B35,入力シート!A30:G32,2,FALSE)'
    sc(ws, 10, 8, '氏　名', F_NORMAL, AL_R, border=None)
    merge_and_set(ws, 10, 9, 10, 18, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 10, 9).value = '=入力シート!B35'
    
    r = 12
    sc(ws, r, 1, '用途', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 2).value = '="（"&入力シート!B12&"）項"'
    
    r = 13
    sc(ws, r, 1, '延べ面積', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 8, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 2).value = '=入力シート!B16&"㎡"'
    sc(ws, r, 9, '階数', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 10, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 10).value = '="地上 "&入力シート!B14&"階"'
    
    r = 14
    sc(ws, r, 1, '住戸数', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 18, None, F_NORMAL, AL_L, THIN)
    auto_cell(ws, r, 2).value = '="全住戸数　"&入力シート!B20&"戸"'
    
    # 住戸用受信機
    r = 15
    merge_and_set(ws, r, 1, r+2, 3, '住戸用受信機', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 4, r, 18, '非蓄積式・蓄積式・自動試験機能付き・遠隔試験機能付き・その他（　）', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+1, 4, r+1, 5, '定格電圧', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+1, 6, r+1, 18, 'ＡＣ　100　Ｖ　・　ＤＣ　　Ｖ', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+2, 4, r+2, 5, '予備電源', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 6, r+2, 18, 'NiCd・その他（　）　Ｖ　　ＡＨ', F_SMALL, AL_L, THIN)
    
    # 共用部分設備
    r = 18
    sc(ws, r, 1, '共用部分', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 18, '住戸用自動火災報知設備・共同住宅用非常警報設備・その他（　）', F_SMALL, AL_L, THIN)
    
    # 中継器 (5 rows)
    r = 19
    merge_and_set(ws, r, 1, r+4, 2, '中継器', F_SMALL, AL_C, THIN)
    for i in range(5):
        rr = r + i
        merge_and_set(ws, rr, 3, rr, 9, '蓄積式・自動試験機能付き・遠隔試験機能付き・他（　）　回線', F_SMALL, AL_L, THIN)
        merge_and_set(ws, rr, 10, rr, 14, '予備電源 有(V AH)・無', F_SMALL, AL_L, THIN)
        merge_and_set(ws, rr, 15, rr, 18, None, F_SMALL, AL_L, THIN)
        if i == 0:
            auto_cell(ws, rr, 15).value = '="設置台数  "&入力シート!B20&"台"'
        else:
            sc(ws, rr, 15, '設置台数　　台')
    
    # 感知器
    r = 24
    merge_and_set(ws, r, 1, r+12, 2, '感知器', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 3, r, 9, '機　種', F_SMALL, AL_C, THIN)
    sc(ws, r, 10, '自', F_SMALL, AL_C, THIN)
    sc(ws, r, 11, '遠', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 12, r, 15, '種別', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 16, r, 18, '個数', F_SMALL, AL_C, THIN)
    
    for i in range(12):
        rr = 25 + i
        merge_and_set(ws, rr, 3, rr, 9, None, F_SMALL, AL_L, THIN)
        input_cell(ws, rr, 3)
        input_cell(ws, rr, 10)
        input_cell(ws, rr, 11)
        merge_and_set(ws, rr, 12, rr, 15, None, F_SMALL, AL_C, THIN)
        input_cell(ws, rr, 12)
        merge_and_set(ws, rr, 16, rr, 18, None, F_SMALL, AL_C, THIN)
        input_cell(ws, rr, 16)
    
    # 起動装置
    r = 37
    sc(ws, r, 1, '起動装置', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r, 5, '屋内型', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 6, r, 9, None, F_SMALL, AL_C, THIN)
    input_cell(ws, r, 6)
    merge_and_set(ws, r, 10, r, 13, '屋外型', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 14, r, 18, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r, 14).value = '=入力シート!B65&"個"'
    
    # 音声警報装置等
    r = 38
    merge_and_set(ws, r, 1, r+3, 2, '音声警報\n装置等', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 3, r, 5, '音声警報装置', F_SMALL, AL_C, THIN)
    sc(ws, r, 6, 'Ｌ級', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 7, r, 9, None, F_SMALL, AL_C, THIN)
    input_cell(ws, r, 7)
    sc(ws, r, 10, 'Ｍ級', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 11, r, 13, None, F_SMALL, AL_C, THIN)
    input_cell(ws, r, 11)
    sc(ws, r, 14, 'Ｓ級', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 15, r, 18, None, F_SMALL, AL_C, THIN)
    input_cell(ws, r, 15)
    
    merge_and_set(ws, r+1, 3, r+1, 5, 'ベル（サイレン）', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+1, 6, r+1, 9, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r+1, 6).value = '="電圧DC "&入力シート!B58&"V"'
    merge_and_set(ws, r+1, 10, r+1, 13, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r+1, 10).value = '="電流 "&入力シート!B59&"mA"'
    merge_and_set(ws, r+1, 14, r+1, 16, '個数', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+1, 17, r+1, 18, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r+1, 17).value = '=入力シート!B60'
    
    merge_and_set(ws, r+2, 3, r+2, 5, '戸外表示器', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 6, r+2, 13, '―――――――', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 14, r+2, 16, '個数', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 17, r+2, 18, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r+2, 17).value = '=入力シート!B20'
    
    merge_and_set(ws, r+3, 3, r+3, 5, '放送設備との連動', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+3, 6, r+3, 18, None, F_SMALL, AL_L, THIN)
    auto_cell(ws, r+3, 6).value = '=入力シート!B90&"　（音声切替装置　"&入力シート!B90&"）"'
    
    return ws

def create_配線試験(wb):
    """別記様式第28 配線の試験結果報告書"""
    ws = wb.create_sheet('配線試験結果')
    ws.sheet_properties.tabColor = 'C00000'
    
    widths = [3, 5, 5, 5, 5, 5, 5, 5, 5, 3, 3, 5, 5, 5, 5, 5, 5]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i+1)].width = w
    
    sc(ws, 1, 1, '別記様式第28', F_SMALL, AL_L, border=None)
    sc(ws, 1, 17, '①', F_SMALL, AL_R, border=None)
    merge_and_set(ws, 3, 1, 3, 17, '配 線 の 試 験 結 果 報 告 書', F_HEADER, AL_C, border=None)
    
    merge_and_set(ws, 5, 1, 5, 17, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 5, 1).value = '="試験実施日　"&入力シート!B86'
    
    sc(ws, 7, 6, '試験実施者', F_HEADER, AL_R, border=None)
    sc(ws, 9, 6, '住　所', F_NORMAL, AL_R, border=None)
    merge_and_set(ws, 9, 7, 9, 17, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 9, 7).value = '=VLOOKUP(入力シート!B36,入力シート!A30:G32,2,FALSE)'
    sc(ws, 11, 6, '氏　名', F_NORMAL, AL_R, border=None)
    merge_and_set(ws, 11, 7, 11, 17, None, F_NORMAL, AL_L, border=None)
    auto_cell(ws, 11, 7).value = '=入力シート!B36'
    
    # 消防用設備等の種類
    r = 13
    merge_and_set(ws, r, 1, r+3, 2, '消防用設備\n等の種類', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 3, r, 17, '屋内消火栓設備　スプリンクラー設備　水噴霧消火設備　泡消火設備', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+1, 3, r+1, 17, '不活性ガス消火設備　ハロゲン化物消火設備　粉末消火設備　屋外消火栓設備', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+2, 3, r+2, 17, '自動火災報知設備　　ガス漏れ火災警報設備　　漏電火災警報器', F_SMALL, AL_L, THIN)
    merge_and_set(ws, r+3, 3, r+3, 17, '消防機関へ通報する火災報知設備　非常警報設備　放送設備　誘導灯', F_SMALL, AL_L, THIN)
    
    # 試験項目
    r = 17
    merge_and_set(ws, r, 1, r, 9, '試　験　項　目', F_HEADER, AL_C, THIN)
    merge_and_set(ws, r, 10, r, 16, '種別・容量等の内容', F_HEADER, AL_C, THIN)
    sc(ws, r, 17, '結果', F_HEADER, AL_C, THIN)
    
    dv_result = DataValidation(type="list", formula1='"○,×,／"', allow_blank=True)
    ws.add_data_validation(dv_result)
    
    # 外観試験
    r = 18
    merge_and_set(ws, r, 1, r+11, 1, '外観試験', F_SMALL, AL_C, THIN)
    
    merge_and_set(ws, r, 2, r+2, 5, '電源回路の\n開閉器・遮断器等', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 6, r, 9, '設置場所等', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 10, r, 16, '―――', F_SMALL, AL_C, THIN)
    sc(ws, r, 17, '○', F_SMALL, AL_C, THIN)
    
    merge_and_set(ws, r+1, 6, r+1, 9, '開閉器', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+1, 10, r+1, 16, '―――', F_SMALL, AL_C, THIN)
    sc(ws, r+1, 17, '／', F_SMALL, AL_C, THIN)
    
    merge_and_set(ws, r+2, 6, r+2, 9, '遮断器', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 10, r+2, 16, '―――', F_SMALL, AL_C, THIN)
    sc(ws, r+2, 17, '○', F_SMALL, AL_C, THIN)
    
    merge_and_set(ws, r+3, 2, r+7, 5, '耐火耐熱\n保護配線', F_SMALL, AL_C, THIN)
    labels2 = [('保護配線の系路','電源回路・操作回路・表示灯回路・警報回路'),
               ('電線の種類・太さ','―――'),('配線方法','―――'),('接続','―――'),('工事方法','―――')]
    for i, (lbl, content) in enumerate(labels2):
        rr = r+3+i
        merge_and_set(ws, rr, 6, rr, 9, lbl, F_SMALL, AL_C, THIN)
        merge_and_set(ws, rr, 10, rr, 16, content, F_SMALL, AL_C, THIN)
        sc(ws, rr, 17, '○', F_SMALL, AL_C, THIN)
        dv_result.add(ws.cell(row=rr, column=17))
    
    merge_and_set(ws, r+8, 2, r+10, 5, '配線（耐火耐熱\n保護配線を除く）', F_SMALL, AL_C, THIN)
    labels3 = [('電線の種類・太さ','―――'),('配線方法','―――'),('接続','―――')]
    for i, (lbl, content) in enumerate(labels3):
        rr = r+8+i
        merge_and_set(ws, rr, 6, rr, 9, lbl, F_SMALL, AL_C, THIN)
        merge_and_set(ws, rr, 10, rr, 16, content, F_SMALL, AL_C, THIN)
        sc(ws, rr, 17, '○', F_SMALL, AL_C, THIN)
        dv_result.add(ws.cell(row=rr, column=17))
    
    merge_and_set(ws, r+11, 2, r+11, 9, '耐震措置', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+11, 10, r+11, 16, '―――', F_SMALL, AL_C, THIN)
    sc(ws, r+11, 17, '／', F_SMALL, AL_C, THIN)
    
    # 機能試験
    r = 30
    merge_and_set(ws, r, 1, r+7, 1, '機能試験', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r+1, 5, '接地抵抗試験', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 6, r, 9, '電圧の種別', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 10, r, 16, '低圧・高圧・特別高圧', F_SMALL, AL_C, THIN)
    sc(ws, r, 17, '／', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+1, 6, r+1, 9, '接地抵抗値', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+1, 10, r+1, 16, None, F_SMALL, AL_C, THIN)
    input_cell(ws, r+1, 10)
    sc(ws, r+1, 17, '／', F_SMALL, AL_C, THIN)
    
    # 絶縁抵抗試験
    merge_and_set(ws, r+2, 2, r+5, 5, '絶縁抵抗試験', F_SMALL, AL_C, THIN)
    resist_labels = ['電源回路','操作回路','表示灯回路','警報回路']
    resist_refs = [(73,74),(75,76),(77,78),(79,80)]
    for i, (lbl, (vref, rref)) in enumerate(zip(resist_labels, resist_refs)):
        rr = r+2+i
        merge_and_set(ws, rr, 6, rr, 9, lbl, F_SMALL, AL_C, THIN)
        merge_and_set(ws, rr, 10, rr, 13, None, F_SMALL, AL_C, THIN)
        auto_cell(ws, rr, 10).value = f'=入力シート!B{vref}'
        merge_and_set(ws, rr, 14, rr, 16, None, F_SMALL, AL_C, THIN)
        auto_cell(ws, rr, 14).value = f'=入力シート!B{rref}'
        sc(ws, rr, 17, '○', F_SMALL, AL_C, THIN)
        dv_result.add(ws.cell(row=rr, column=17))
    
    # Page 2 header
    r = 38
    sc(ws, r, 1, '配線', F_SMALL, AL_L, border=None)
    sc(ws, r, 17, '②', F_SMALL, AL_R, border=None)
    
    r = 39
    merge_and_set(ws, r, 1, r, 9, '試　験　項　目', F_HEADER, AL_C, THIN)
    merge_and_set(ws, r, 10, r, 16, '種別・容量等の内容', F_HEADER, AL_C, THIN)
    sc(ws, r, 17, '結果', F_HEADER, AL_C, THIN)
    
    r = 40
    merge_and_set(ws, r, 1, r+2, 1, '機能試験', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 2, r+1, 5, '絶縁抵抗試験', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 6, r, 9, '感知器回路', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r, 10, r, 13, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r, 10).value = '=入力シート!B81'
    merge_and_set(ws, r, 14, r, 16, None, F_SMALL, AL_C, THIN)
    auto_cell(ws, r, 14).value = '=入力シート!B82'
    sc(ws, r, 17, '○', F_SMALL, AL_C, THIN)
    dv_result.add(ws.cell(row=r, column=17))
    
    merge_and_set(ws, r+1, 6, r+1, 9, '付属装置回路等', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+1, 10, r+1, 16, None, F_SMALL, AL_C, THIN)
    input_cell(ws, r+1, 10)
    sc(ws, r+1, 17, '／', F_SMALL, AL_C, THIN)
    
    merge_and_set(ws, r+2, 2, r+2, 9, '絶縁耐力試験', F_SMALL, AL_C, THIN)
    merge_and_set(ws, r+2, 10, r+2, 16, None, F_SMALL, AL_C, THIN)
    sc(ws, r+2, 17, '／', F_SMALL, AL_C, THIN)
    
    # 備考
    r = 44
    merge_and_set(ws, r, 1, r+3, 17, '備　考', F_SMALL, AL_TL, THIN)
    auto_cell(ws, r, 1).value = '="備考\\n試験実施者が有している資格：\\n"'
    
    r = 48
    sc(ws, r, 1, '備考　１　この用紙の大きさは、日本産業規格Ａ４とすること。', F_SMALL, AL_L, border=None)
    
    return ws


# Create all sheets
create_着工届(wb)
create_設置届(wb)
create_概要表その1(wb)
create_概要表その2(wb)
create_試験結果11(wb)
create_試験結果34(wb)
create_配線試験(wb)

# Save
output_path = '/home/user/claude-project/消防設備工事_様式統合.xlsx'
wb.save(output_path)
print(f"Saved: {output_path}")
print(f"Sheets: {wb.sheetnames}")
