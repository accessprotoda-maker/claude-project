import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy
import os

wb = openpyxl.Workbook()

# Common styles
header_font = Font(name='MS ゴシック', bold=True, size=11)
input_font = Font(name='MS ゴシック', size=10)
small_font = Font(name='MS ゴシック', size=8)
title_font = Font(name='MS ゴシック', bold=True, size=14)
section_font = Font(name='MS ゴシック', bold=True, size=12)

thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
input_fill = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')  # Yellow for input cells
auto_fill = PatternFill(start_color='E0FFE0', end_color='E0FFE0', fill_type='solid')  # Green for auto-calculated
header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')  # Blue header
section_fill = PatternFill(start_color='B4C6E7', end_color='B4C6E7', fill_type='solid')

def apply_style(ws, row, col, value=None, font=None, fill=None, alignment=None, border=None):
    cell = ws.cell(row=row, column=col, value=value)
    if font: cell.font = font
    if fill: cell.fill = fill
    if alignment: cell.alignment = alignment
    if border: cell.border = border
    return cell

def apply_border_range(ws, min_row, max_row, min_col, max_col):
    for r in range(min_row, max_row+1):
        for c in range(min_col, max_col+1):
            ws.cell(row=r, column=c).border = thin_border

def make_input_cell(ws, row, col, value=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = input_fill
    cell.border = thin_border
    cell.font = input_font
    cell.alignment = Alignment(horizontal='center', vertical='center')
    return cell

def make_header_cell(ws, row, col, value=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = header_fill
    cell.border = thin_border
    cell.font = header_font
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    return cell

# ============================================================
# Sheet 1: 入力フォーム (Master Input Sheet)
# ============================================================
ws_input = wb.active
ws_input.title = '入力フォーム'
ws_input.sheet_properties.tabColor = 'FF0000'

# Column widths
for c in range(1, 20):
    ws_input.column_dimensions[get_column_letter(c)].width = 18

row = 1
ws_input.merge_cells('A1:H1')
apply_style(ws_input, 1, 1, '消防設備工事 統合入力フォーム', title_font, 
            alignment=Alignment(horizontal='center', vertical='center'))
ws_input.row_dimensions[1].height = 35

row = 2
ws_input.merge_cells('A2:H2')
apply_style(ws_input, 2, 1, '黄色セル＝入力必要  緑色セル＝自動計算', input_font,
            alignment=Alignment(horizontal='center'))

# ---- Section: 基本情報 ----
row = 4
ws_input.merge_cells(f'A{row}:H{row}')
apply_style(ws_input, row, 1, '【基本情報】', section_font, section_fill,
            Alignment(horizontal='left', vertical='center'))

labels_basic = [
    ('担当消防署', 'B5'), ('届出日（例：2025/4/1）', 'B6'), 
    ('届出者名', 'B7'), ('工事場所（住所）', 'B8'),
    ('防火対象物の名称', 'B9'), ('工事整備対象設備等の種類', 'B10'),
    ('用途コード（例：5項ロ）', 'B11'), ('用途名称', 'B12'),
    ('主要構造部', 'B13'), ('地上階数', 'B14'),
    ('地下階数', 'B15'), ('塔屋階数', 'B16'),
    ('延べ床面積（㎡）', 'B17'), ('着工予定日', 'B18'),
    ('完成予定日', 'B19'), ('理事長名', 'B20'),
]

for i, (label, _) in enumerate(labels_basic):
    r = 5 + i
    make_header_cell(ws_input, r, 1, label)
    make_input_cell(ws_input, r, 2)
    # Set some defaults
    if label == '工事整備対象設備等の種類':
        ws_input.cell(row=r, column=2, value='自動火災報知設備')
    elif label == '用途名称':
        ws_input.cell(row=r, column=2, value='共同住宅')
    elif label == '主要構造部':
        ws_input.cell(row=r, column=2, value='耐火構造')

# Dropdown for 用途コード
dv_usage = DataValidation(type="list", formula1='"5項ロ,5項イ,6項イ,6項ロ,6項ハ,16項イ"', allow_blank=True)
dv_usage.error = "リストから選択してください"
dv_usage.errorTitle = "入力エラー"
ws_input.add_data_validation(dv_usage)
dv_usage.add(ws_input['B11'])

# Dropdown for 主要構造部
dv_struct = DataValidation(type="list", formula1='"耐火構造,準耐火構造,その他"', allow_blank=True)
ws_input.add_data_validation(dv_struct)
dv_struct.add(ws_input['B13'])

# Dropdown for 設備種類
dv_equip = DataValidation(type="list", formula1='"自動火災報知設備,共同住宅用自動火災報知設備,住戸用自動火災報知設備"', allow_blank=True)
ws_input.add_data_validation(dv_equip)
dv_equip.add(ws_input['B10'])

# ---- Section: 施工者情報 ----
row = 22
ws_input.merge_cells(f'A{row}:H{row}')
apply_style(ws_input, row, 1, '【施工者情報】', section_font, section_fill,
            Alignment(horizontal='left', vertical='center'))

labels_contractor = [
    '施工会社名', '代表者名', '施工会社住所', '電話番号',
]
for i, label in enumerate(labels_contractor):
    r = 23 + i
    make_header_cell(ws_input, r, 1, label)
    make_input_cell(ws_input, r, 2)

# Default values
ws_input['B23'] = '株式会社　TSCアクセス・プロ'
ws_input['B24'] = '代表取締役　中村　勇'
ws_input['B25'] = '名古屋市中区栄４丁目１６番３号　山光堂ビル５階'
ws_input['B26'] = '052-269-9100'

# ---- Section: 消防設備士情報 ----
row = 28
ws_input.merge_cells(f'A{row}:H{row}')
apply_style(ws_input, row, 1, '【消防設備士情報】担当者選択', section_font, section_fill,
            Alignment(horizontal='left', vertical='center'))

# Staff master list
staff_headers = ['氏名', '住所', '免状種類', '交付年月日', '交付番号', '講習年月', '受講地', '交付知事', '電工資格']
for i, h in enumerate(staff_headers):
    make_header_cell(ws_input, 29, 1+i, h)

staff_data = [
    ['河合　佑樹', '愛知県尾張旭市晴丘町東236番地4', '甲4', 'R5年7月21日', '第00324号', '', '愛知', '愛知', '第二種電気工事士 R3年9月7日 第144528号'],
    ['戸田　健仁', '愛知県長久手市西原山1-1 ｾﾝﾄｱｰｽ502', '甲4', 'R5年11月7日', '第00448号', '', '愛知', '愛知', '第二種電気工事士 H22年10月13日 第106680号'],
    ['棚町　征', '岐阜県土岐市泉町大富195番地の249', '甲4', 'H14年10月30日', '第00009号', '令和元年10月', '愛知', '愛知', '第二種電気工事士 H27年12月4日 岐阜県第31879号'],
]
for i, staff in enumerate(staff_data):
    for j, val in enumerate(staff):
        cell = ws_input.cell(row=30+i, column=1+j, value=val)
        cell.border = thin_border
        cell.font = input_font

row = 34
make_header_cell(ws_input, row, 1, '届出者（選択）')
make_input_cell(ws_input, row, 2)
dv_staff1 = DataValidation(type="list", formula1='"河合　佑樹,戸田　健仁,棚町　征"', allow_blank=True)
ws_input.add_data_validation(dv_staff1)
dv_staff1.add(ws_input['B34'])
ws_input['B34'] = '戸田　健仁'

make_header_cell(ws_input, 35, 1, '試験実施者（選択）')
make_input_cell(ws_input, 35, 2)
dv_staff2 = DataValidation(type="list", formula1='"河合　佑樹,戸田　健仁,棚町　征"', allow_blank=True)
ws_input.add_data_validation(dv_staff2)
dv_staff2.add(ws_input['B35'])

make_header_cell(ws_input, 36, 1, '配線試験実施者（選択）')
make_input_cell(ws_input, 36, 2)
dv_staff3 = DataValidation(type="list", formula1='"河合　佑樹,戸田　健仁,棚町　征"', allow_blank=True)
ws_input.add_data_validation(dv_staff3)
dv_staff3.add(ws_input['B36'])
ws_input['B36'] = '戸田　健仁'

# ---- Section: 受信機情報 ----
row = 38
ws_input.merge_cells(f'A{row}:H{row}')
apply_style(ws_input, row, 1, '【受信機・機器情報】', section_font, section_fill,
            Alignment(horizontal='left', vertical='center'))

recv_labels = [
    ('共用部 受信機型級（例：1）', 'B39'), ('共用部 回線数（例：10/10）', 'B40'),
    ('共用部 予備電源DC（V）', 'B41'), ('共用部 予備電源（AH）', 'B42'),
    ('共用部 受信機設置場所', 'B43'), ('共用部 受信機メーカー', 'B44'),
    ('共用部 受信機型式番号（受第）', 'B45'),
    ('住戸部 受信機型級（例：3）', 'B46'), ('住戸部 回線数（例：1/1）', 'B47'),
    ('住戸部 受信機設置場所', 'B48'), ('住戸部 受信機メーカー', 'B49'),
    ('住戸部 受信機型式番号', 'B50'),
    ('住戸部 表示器台数', 'B51'),
    ('全住戸数', 'B52'),
    ('関連設備', 'B53'),
]
for i, (label, _) in enumerate(recv_labels):
    r = 39 + i
    make_header_cell(ws_input, r, 1, label)
    make_input_cell(ws_input, r, 2)

# Defaults
ws_input['B39'] = 1
ws_input['B40'] = '10/10'
ws_input['B41'] = 24
ws_input['B42'] = 0.45
ws_input['B43'] = '管理人室'
ws_input['B44'] = 'パナソニック㈱'
ws_input['B46'] = 3
ws_input['B47'] = '1/1'
ws_input['B48'] = '各住戸リビング'
ws_input['B49'] = 'パナソニック㈱'
ws_input['B53'] = 'インターホン設備'

# Dropdown for メーカー
dv_maker = DataValidation(type="list", formula1='"パナソニック㈱,ニッタン(株),能美防災㈱,ホーチキ㈱"', allow_blank=True)
ws_input.add_data_validation(dv_maker)
dv_maker.add(ws_input['B44'])
dv_maker.add(ws_input['B49'])

# ---- Section: 地区音響装置 ----
row = 55
ws_input.merge_cells(f'A{row}:H{row}')
apply_style(ws_input, row, 1, '【地区音響装置・発信機】', section_font, section_fill,
            Alignment(horizontal='left', vertical='center'))

sound_labels = [
    ('地区音響 認評音第（例：13-5）', 'B56'), ('地区音響 メーカー', 'B57'),
    ('地区音響 電圧DC（V）', 'B58'), ('地区音響 電流（mA）', 'B59'),
    ('地区音響 個数', 'B60'), ('地区音響 音量（dB）', 'B61'),
    ('発信機 型（P等）', 'B62'), ('発信機 級', 'B63'), ('発信機 個数', 'B64'),
    ('発信機 型式番号（発第）', 'B65'), ('発信機 メーカー', 'B66'),
    ('表示灯 電圧DC（V）', 'B67'), ('表示灯 電流（mA）', 'B68'),
    ('表示灯 個数', 'B69'),
]
for i, (label, _) in enumerate(sound_labels):
    r = 56 + i
    make_header_cell(ws_input, r, 1, label)
    make_input_cell(ws_input, r, 2)

# Defaults
ws_input['B56'] = '13-5'
ws_input['B57'] = 'パナソニック㈱'
ws_input['B58'] = 24
ws_input['B59'] = 10
ws_input['B61'] = 95
ws_input['B62'] = 'P'
ws_input['B63'] = 1
ws_input['B65'] = '29-13'
ws_input['B66'] = 'パナソニック㈱'
ws_input['B67'] = 24
ws_input['B68'] = 9

# ---- Section: 工事種別 ----
row = 71
ws_input.merge_cells(f'A{row}:H{row}')
apply_style(ws_input, row, 1, '【工事種別】', section_font, section_fill,
            Alignment(horizontal='left', vertical='center'))

make_header_cell(ws_input, 72, 1, '工事種別')
make_input_cell(ws_input, 72, 2)
dv_work = DataValidation(type="list", formula1='"施設,増設,移設,取替え,改造,その他"', allow_blank=True)
ws_input.add_data_validation(dv_work)
dv_work.add(ws_input['B72'])

# ---- Section: 絶縁抵抗 ----
row = 74
ws_input.merge_cells(f'A{row}:H{row}')
apply_style(ws_input, row, 1, '【配線試験 絶縁抵抗値】', section_font, section_fill,
            Alignment(horizontal='left', vertical='center'))

resist_labels = [
    ('電源回路 電圧（V）', 'B75'), ('電源回路 抵抗（MΩ）', 'B76'),
    ('操作回路 電圧（V）', 'B77'), ('操作回路 抵抗（MΩ）', 'B78'),
    ('表示灯回路 電圧（V）', 'B79'), ('表示灯回路 抵抗（MΩ）', 'B80'),
    ('警報回路 電圧（V）', 'B81'), ('警報回路 抵抗（MΩ）', 'B82'),
    ('感知器回路 電圧（V）', 'B83'), ('感知器回路 抵抗（MΩ）', 'B84'),
]
for i, (label, _) in enumerate(resist_labels):
    r = 75 + i
    make_header_cell(ws_input, r, 1, label)
    make_input_cell(ws_input, r, 2)

# Defaults
ws_input['B75'] = 250; ws_input['B76'] = 150
ws_input['B77'] = 50; ws_input['B78'] = 20
ws_input['B79'] = 50; ws_input['B80'] = 20
ws_input['B81'] = 50; ws_input['B82'] = 20
ws_input['B83'] = 50; ws_input['B84'] = 20

# ============================================================
# Sheet 2: 共用部感知器入力
# ============================================================
ws_common = wb.create_sheet('共用部感知器')
ws_common.sheet_properties.tabColor = '0070C0'

ws_common.merge_cells('A1:Q1')
apply_style(ws_common, 1, 1, '共用部 感知器・設備 入力表', title_font,
            alignment=Alignment(horizontal='center'))

headers_common = ['警戒区域番号', '名称等', '住戸用受信機', '差動式スポット型', '定温式スポット型',
                  '補償式スポット型', '熱アナログ式スポット型', '光電式スポット型', '光電アナログ式スポット型',
                  'イオン化式スポット型', 'イオン化アナログ式スポット型', '炎感知器',
                  '音声警報装置', '戸外表示器', '外観試験', '機能試験', '備考']

for i, h in enumerate(headers_common):
    make_header_cell(ws_common, 3, i+1, h)
    ws_common.column_dimensions[get_column_letter(i+1)].width = 12 if i >= 3 else 16

# Data rows (25 rows for input)
for r in range(4, 28):
    for c in range(1, 18):
        if c <= 14:
            make_input_cell(ws_common, r, c)
        elif c <= 16:
            # Auto: 〇 if any sensor data
            cell = ws_common.cell(row=r, column=c)
            cell.fill = auto_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
            col_letter = get_column_letter(c)
            cell.value = f'=IF(COUNTA(C{r}:N{r})=0,"","〇")'
        else:
            make_input_cell(ws_common, r, c)

# Subtotal row
r = 28
make_header_cell(ws_common, r, 1, '小計')
ws_common.merge_cells(f'A{r}:B{r}')
for c in range(3, 15):
    cell = ws_common.cell(row=r, column=c)
    cell.fill = auto_fill
    cell.border = thin_border
    cell.font = Font(name='MS ゴシック', bold=True)
    cl = get_column_letter(c)
    cell.value = f'=SUM({cl}4:{cl}27)'

# Dropdown for result
dv_result = DataValidation(type="list", formula1='"〇,×"', allow_blank=True)
ws_common.add_data_validation(dv_result)

# Sample data
sample_common = [
    [1, '1階ｴﾝﾄﾗﾝｽ', '', '', 5, '', '', 1, '', '', '', '', '', ''],
    [2, 'エレベーター', '', '', '', '', '', 1, '', '', '', '', '', ''],
    [3, '1階共用部 防火戸', '', '', '', '', '', 2, '', '', '', '', '', ''],
]
for i, data in enumerate(sample_common):
    for j, val in enumerate(data):
        if val != '':
            ws_common.cell(row=4+i, column=1+j, value=val)

# ============================================================
# Sheet 3: 専有部感知器入力
# ============================================================
ws_private = wb.create_sheet('専有部感知器')
ws_private.sheet_properties.tabColor = '00B050'

ws_private.merge_cells('A1:Q1')
apply_style(ws_private, 1, 1, '専有部（住戸部）感知器・設備 入力表', title_font,
            alignment=Alignment(horizontal='center'))

apply_style(ws_private, 2, 1, '※ 住戸数が多い場合は行を追加してください', small_font)

for i, h in enumerate(headers_common):
    make_header_cell(ws_private, 3, i+1, h)
    ws_private.column_dimensions[get_column_letter(i+1)].width = 12 if i >= 3 else 16

# 80 rows for private units
for r in range(4, 84):
    for c in range(1, 18):
        if c <= 14:
            make_input_cell(ws_private, r, c)
        elif c <= 16:
            cell = ws_private.cell(row=r, column=c)
            cell.fill = auto_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.value = f'=IF(COUNTA(C{r}:N{r})=0,"","〇")'
        else:
            make_input_cell(ws_private, r, c)

# Subtotal
r = 84
make_header_cell(ws_private, r, 1, '合計')
ws_private.merge_cells(f'A{r}:B{r}')
for c in range(3, 15):
    cell = ws_private.cell(row=r, column=c)
    cell.fill = auto_fill
    cell.border = thin_border
    cell.font = Font(name='MS ゴシック', bold=True)
    cl = get_column_letter(c)
    cell.value = f'=SUM({cl}4:{cl}83)'

# ============================================================
# Sheet 4: 共用部感知器入力（概要表用）
# ============================================================
ws_sensor_common = wb.create_sheet('共用部概要表入力')
ws_sensor_common.sheet_properties.tabColor = 'FFC000'

ws_sensor_common.merge_cells('A1:M1')
apply_style(ws_sensor_common, 1, 1, '共用部 感知器概要表（プルダウン入力）', title_font,
            alignment=Alignment(horizontal='center'))

sensor_headers = ['機種名', '型（スポット型等）', '蓄積', '自動', '遠隔', '種別', '個数',
                  '型式番号From', '型式番号To', '製造会社名']
for i, h in enumerate(sensor_headers):
    make_header_cell(ws_sensor_common, 3, i+1, h)
    ws_sensor_common.column_dimensions[get_column_letter(i+1)].width = 14

# Dropdowns
dv_type = DataValidation(type="list", formula1='"差動式,定温式,光電式,補償式,熱アナログ式,光電アナログ式,イオン化式"', allow_blank=True)
ws_sensor_common.add_data_validation(dv_type)

dv_yesno = DataValidation(type="list", formula1='"〇,"', allow_blank=True)
ws_sensor_common.add_data_validation(dv_yesno)

dv_grade = DataValidation(type="list", formula1='"特種,1種,2種,3種"', allow_blank=True)
ws_sensor_common.add_data_validation(dv_grade)

for r in range(4, 14):
    for c in range(1, 11):
        make_input_cell(ws_sensor_common, r, c)
    dv_type.add(ws_sensor_common.cell(row=r, column=1))
    dv_yesno.add(ws_sensor_common.cell(row=r, column=3))
    dv_yesno.add(ws_sensor_common.cell(row=r, column=4))
    dv_yesno.add(ws_sensor_common.cell(row=r, column=5))
    dv_grade.add(ws_sensor_common.cell(row=r, column=6))
    dv_maker2 = DataValidation(type="list", formula1='"パナソニック㈱,ニッタン(株),能美防災㈱,ホーチキ㈱"', allow_blank=True)
    ws_sensor_common.add_data_validation(dv_maker2)
    dv_maker2.add(ws_sensor_common.cell(row=r, column=10))

# Same for 住戸部
row = 16
ws_sensor_common.merge_cells(f'A{row}:M{row}')
apply_style(ws_sensor_common, row, 1, '住戸部 感知器概要表（プルダウン入力）', title_font,
            alignment=Alignment(horizontal='center'))

for i, h in enumerate(sensor_headers):
    make_header_cell(ws_sensor_common, 18, i+1, h)

for r in range(19, 29):
    for c in range(1, 11):
        make_input_cell(ws_sensor_common, r, c)
    dv_type.add(ws_sensor_common.cell(row=r, column=1))
    dv_yesno.add(ws_sensor_common.cell(row=r, column=3))
    dv_yesno.add(ws_sensor_common.cell(row=r, column=4))
    dv_yesno.add(ws_sensor_common.cell(row=r, column=5))
    dv_grade.add(ws_sensor_common.cell(row=r, column=6))

# ============================================================
# Sheet 5: 自火報その2（警戒区域入力）
# ============================================================
ws_zone = wb.create_sheet('警戒区域入力')
ws_zone.sheet_properties.tabColor = '7030A0'

ws_zone.merge_cells('A1:N1')
apply_style(ws_zone, 1, 1, '自動火災報知設備 警戒区域別 感知器・音響装置（その2）', title_font,
            alignment=Alignment(horizontal='center'))

zone_headers = ['表示番号', '名称', '差動式分布型', '差動式スポット型', '補償式スポット型',
                '定温式スポット型', '定温式感知線型', 'イオン化式スポット型', 
                '光電式スポット型', '光電式分離型', '炎感知器', '地区音響装置', '結果']
for i, h in enumerate(zone_headers):
    make_header_cell(ws_zone, 3, i+1, h)
    ws_zone.column_dimensions[get_column_letter(i+1)].width = 14

dv_result2 = DataValidation(type="list", formula1='"〇,×,／"', allow_blank=True)
ws_zone.add_data_validation(dv_result2)

for r in range(4, 24):
    for c in range(1, 14):
        if c == 13:
            make_input_cell(ws_zone, r, c)
            dv_result2.add(ws_zone.cell(row=r, column=c))
        else:
            make_input_cell(ws_zone, r, c)

# Total row
r = 24
make_header_cell(ws_zone, r, 1, '合計')
ws_zone.merge_cells(f'A{r}:B{r}')
for c in range(3, 13):
    cell = ws_zone.cell(row=r, column=c)
    cell.fill = auto_fill
    cell.border = thin_border
    cell.font = Font(name='MS ゴシック', bold=True)
    cl = get_column_letter(c)
    cell.value = f'=SUM({cl}4:{cl}23)'

# ============================================================
# Sheet 6: 面積概要表入力
# ============================================================
ws_area = wb.create_sheet('面積概要表入力')
ws_area.sheet_properties.tabColor = 'ED7D31'

ws_area.merge_cells('A1:F1')
apply_style(ws_area, 1, 1, '面積概要表 入力', title_font,
            alignment=Alignment(horizontal='center'))

area_headers = ['階', '床面積（㎡）', '用途・室名', '構造', '内装天井', '内装壁']
for i, h in enumerate(area_headers):
    make_header_cell(ws_area, 3, i+1, h)
    ws_area.column_dimensions[get_column_letter(i+1)].width = 16

for r in range(4, 20):
    for c in range(1, 7):
        make_input_cell(ws_area, r, c)
    # Auto-fill floor name
    cell = ws_area.cell(row=r, column=1)
    cell.fill = auto_fill
    floor_num = r - 3
    if floor_num <= 15:
        cell.value = f'=IF({floor_num}<=入力フォーム!B14,"{floor_num}階","")'

# ============================================================
# Sheet 7: チェックリスト
# ============================================================
ws_check = wb.create_sheet('提出前チェックリスト')
ws_check.sheet_properties.tabColor = 'FF0000'

ws_check.merge_cells('A1:D1')
apply_style(ws_check, 1, 1, '提出前チェックリスト', title_font,
            alignment=Alignment(horizontal='center'))

ws_check.column_dimensions['A'].width = 6
ws_check.column_dimensions['B'].width = 40
ws_check.column_dimensions['C'].width = 12
ws_check.column_dimensions['D'].width = 30

make_header_cell(ws_check, 3, 1, 'No.')
make_header_cell(ws_check, 3, 2, 'チェック項目')
make_header_cell(ws_check, 3, 3, '確認')
make_header_cell(ws_check, 3, 4, '備考')

dv_check = DataValidation(type="list", formula1='"✓,―"', allow_blank=True)
ws_check.add_data_validation(dv_check)

checklist_items_着工 = [
    '表紙：届出日が記入されているか',
    '表紙：届出者名・住所が正しいか',
    '表紙：消防設備士の免状情報が正しいか',
    '表紙：工事種別に○がついているか',
    '表紙：着工予定日・完成予定日が記入されているか',
    '面積概要表：延べ床面積が正しいか',
    '面積概要表：階数が正しいか',
    '共用部概要表：感知器の種類・個数が図面と一致するか',
    '共用部概要表：発信機・表示灯の個数が正しいか',
    '住戸部概要表：感知器の種類・個数が図面と一致するか',
    '住戸部概要表：住戸用受信機台数が正しいか',
    '共用部・住戸部の地区音響装置情報が正しいか',
    '地図に該当建物が色付けされているか',
    '図面のコピーに正は色付けされているか',
]

checklist_items_設置 = [
    '設置届出書：完成年月日が記入されているか',
    '試験結果報告書：試験実施日が記入されているか',
    '試験結果報告書：試験実施者名・住所が正しいか',
    '機器台数一覧：共用部の感知器台数が正しいか',
    '機器台数一覧：専有部の感知器台数が正しいか',
    '機器台数一覧：結果欄がすべて記入されているか',
    '配線試験：絶縁抵抗値が記入されているか',
    '配線試験：試験実施者の資格情報が正しいか',
]

row = 4
apply_style(ws_check, row, 1, '', section_font)
ws_check.merge_cells(f'A{row}:D{row}')
apply_style(ws_check, row, 1, '【着工届】', section_font, section_fill)
row += 1

for i, item in enumerate(checklist_items_着工):
    r = row + i
    ws_check.cell(row=r, column=1, value=i+1).border = thin_border
    ws_check.cell(row=r, column=2, value=item).border = thin_border
    cell = ws_check.cell(row=r, column=3)
    cell.border = thin_border
    cell.fill = input_fill
    cell.alignment = Alignment(horizontal='center')
    dv_check.add(cell)
    ws_check.cell(row=r, column=4).border = thin_border
    ws_check.cell(row=r, column=4).fill = input_fill

row = row + len(checklist_items_着工) + 1
ws_check.merge_cells(f'A{row}:D{row}')
apply_style(ws_check, row, 1, '【設置届】', section_font, section_fill)
row += 1

for i, item in enumerate(checklist_items_設置):
    r = row + i
    ws_check.cell(row=r, column=1, value=i+1).border = thin_border
    ws_check.cell(row=r, column=2, value=item).border = thin_border
    cell = ws_check.cell(row=r, column=3)
    cell.border = thin_border
    cell.fill = input_fill
    cell.alignment = Alignment(horizontal='center')
    dv_check.add(cell)
    ws_check.cell(row=r, column=4).border = thin_border
    ws_check.cell(row=r, column=4).fill = input_fill

# ============================================================
# Sheet 8: 必要書類一覧
# ============================================================
ws_docs = wb.create_sheet('必要書類一覧')
ws_docs.sheet_properties.tabColor = '00B0F0'

ws_docs.merge_cells('A1:E1')
apply_style(ws_docs, 1, 1, '必要書類一覧', title_font,
            alignment=Alignment(horizontal='center'))

ws_docs.column_dimensions['A'].width = 6
ws_docs.column_dimensions['B'].width = 40
ws_docs.column_dimensions['C'].width = 12
ws_docs.column_dimensions['D'].width = 6
ws_docs.column_dimensions['E'].width = 40

make_header_cell(ws_docs, 3, 1, 'No.')
make_header_cell(ws_docs, 3, 2, '着工届')
make_header_cell(ws_docs, 3, 3, '正・副・控')
make_header_cell(ws_docs, 3, 4, 'No.')
make_header_cell(ws_docs, 3, 5, '設置届')

着工_docs = [
    '表紙（着工届出書）',
    '面積概要表',
    '地図（該当建物に色付け）',
    '共用部その1・その2',
    '住戸部その1・その2',
    '消防に関する機器の図面',
    '図面のコピー（正は色付け）',
    '自火報・インターホンの系統図',
    '平面図（各フロア図）',
]

設置_docs = [
    '設置届出書',
    '試験結果報告書',
    '　・自動火災報知設備試験結果報告書',
    '　・共同住宅用自火報試験結果報告書',
    '　・住戸用自火報試験結果報告書',
    '機器台数一覧',
    '　・共同住宅用非常警報設備',
    '　・住戸用自火報・非常警報設備',
    '配線試験結果',
]

for i, doc in enumerate(着工_docs):
    r = 4 + i
    ws_docs.cell(row=r, column=1, value=f'①' if i == 0 else f'{"②③④⑤⑥⑦⑧⑨"[i-1]}' if i < 9 else '').border = thin_border
    ws_docs.cell(row=r, column=2, value=doc).border = thin_border
    ws_docs.cell(row=r, column=3, value='正・副・控').border = thin_border

for i, doc in enumerate(設置_docs):
    r = 4 + i
    nums = ['①','②','②','②','②','③','③','③','④']; ws_docs.cell(row=r, column=4, value=nums[i] if i < len(nums) else '').border = thin_border
    ws_docs.cell(row=r, column=5, value=doc).border = thin_border

# ============================================================
# Save
# ============================================================
output_path = '/home/user/claude-project/消防設備工事_統合書類.xlsx'
wb.save(output_path)
print(f"Saved to {output_path}")
