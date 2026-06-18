#!/usr/bin/env python3
"""
消防設備工事用書類（Word様式）をExcelに変換し、入力シートから自動反映させる統合ファイル
Word文書のテーブルレイアウトを忠実に再現
"""
import sys
from docx import Document
from docx.oxml.ns import qn
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# ============================================================
# Styles
# ============================================================
F_TITLE = Font(name='ＭＳ ゴシック', bold=True, size=14)
F_HEADER = Font(name='ＭＳ ゴシック', bold=True, size=9)
F_NORMAL = Font(name='ＭＳ ゴシック', size=9)
F_SMALL = Font(name='ＭＳ ゴシック', size=7.5)
F_SECTION = Font(name='ＭＳ ゴシック', bold=True, size=11)

THIN = Border(
    left=Side('thin'), right=Side('thin'),
    top=Side('thin'), bottom=Side('thin'))

FILL_INPUT = PatternFill('solid', fgColor='FFFFCC')
FILL_AUTO = PatternFill('solid', fgColor='E0FFE0')
FILL_HEADER = PatternFill('solid', fgColor='D9E1F2')
FILL_SECTION = PatternFill('solid', fgColor='B4C6E7')

AL_C = Alignment(horizontal='center', vertical='center', wrap_text=True)
AL_L = Alignment(horizontal='left', vertical='center', wrap_text=True)
AL_R = Alignment(horizontal='right', vertical='center', wrap_text=True)
AL_TL = Alignment(horizontal='left', vertical='top', wrap_text=True)


def sc(ws, r, c, val=None, font=F_NORMAL, al=AL_C, border=THIN, fill=None):
    cell = ws.cell(row=r, column=c, value=val)
    if font: cell.font = font
    if al: cell.alignment = al
    if border: cell.border = border
    if fill: cell.fill = fill
    return cell


def input_cell(ws, r, c, val=None):
    return sc(ws, r, c, val, fill=FILL_INPUT)


def header_cell(ws, r, c, val=None):
    return sc(ws, r, c, val, font=F_HEADER, fill=FILL_HEADER)


def merge_and_set(ws, r1, c1, r2, c2, val=None, font=F_NORMAL, al=AL_C, border=THIN, fill=None):
    if r1 != r2 or c1 != c2:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    sc(ws, r1, c1, val, font, al, border, fill)
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            ws.cell(row=r, column=c).border = border


# ============================================================
# Word table -> Excel conversion engine
# ============================================================

def get_grid_col_widths(table):
    """Get column widths from tblGrid XML element (in EMU/twips)."""
    grid = table._tbl.find(qn('w:tblGrid'))
    if grid is None:
        return []
    cols = grid.findall(qn('w:gridCol'))
    return [int(c.get(qn('w:w'), '0')) for c in cols]


def emu_to_excel_width(emu_widths, target_total=85):
    """Convert EMU/twip widths to Excel character widths proportionally."""
    total = sum(emu_widths)
    if total == 0:
        return [10] * len(emu_widths)
    return [(w / total) * target_total for w in emu_widths]


def parse_table_cells(table):
    """Parse table XML to get cells with accurate merge info.
    Returns list of rows, each row is list of dicts:
    {col_idx, span, text, is_vmerge_start, is_vmerge_continue}
    """
    rows = []
    for tr in table._tbl.findall(qn('w:tr')):
        row_cells = []
        col_idx = 0
        for tc in tr.findall(qn('w:tc')):
            tcPr = tc.find(qn('w:tcPr'))

            # Horizontal span
            gridSpan_elem = tcPr.find(qn('w:gridSpan')) if tcPr is not None else None
            span = int(gridSpan_elem.get(qn('w:val'))) if gridSpan_elem is not None else 1

            # Vertical merge
            vMerge_elem = tcPr.find(qn('w:vMerge')) if tcPr is not None else None
            is_vmerge_start = False
            is_vmerge_continue = False
            if vMerge_elem is not None:
                val = vMerge_elem.get(qn('w:val'))
                if val == 'restart':
                    is_vmerge_start = True
                else:
                    is_vmerge_continue = True

            # Cell text - join paragraphs with newline
            paragraphs = tc.findall(qn('w:p'))
            texts = []
            for p in paragraphs:
                # Collect all text runs in this paragraph
                p_text = ''
                for node in p.iter():
                    if node.tag == qn('w:t'):
                        p_text += (node.text or '')
                texts.append(p_text)
            text = '\n'.join(texts).strip()

            row_cells.append({
                'col_idx': col_idx,
                'span': span,
                'text': text,
                'is_vmerge_start': is_vmerge_start,
                'is_vmerge_continue': is_vmerge_continue,
            })
            col_idx += span
        rows.append(row_cells)
    return rows


def compute_vmerge_ranges(parsed_rows):
    """Compute vertical merge ranges. Returns dict: (start_row, col_idx) -> end_row."""
    # Track which col_idx has an active vmerge and from which start row
    active = {}  # col_idx -> start_row
    vmerge_map = {}  # (start_row, col_idx) -> end_row

    for row_idx, row in enumerate(parsed_rows):
        # Track which col_idxs appear in this row
        seen_cols = set()
        for cell in row:
            ci = cell['col_idx']
            seen_cols.add(ci)
            if cell['is_vmerge_start']:
                active[ci] = row_idx
                vmerge_map[(row_idx, ci)] = row_idx
            elif cell['is_vmerge_continue']:
                if ci in active:
                    vmerge_map[(active[ci], ci)] = row_idx
            else:
                # No vmerge - end any active merge for this col
                if ci in active:
                    del active[ci]

    return vmerge_map


def write_table_to_sheet(ws, table, start_row=1, target_width=85):
    """Write a Word table to an Excel worksheet starting at start_row.
    Returns the next available row after the table.
    """
    grid_widths = get_grid_col_widths(table)
    num_grid_cols = len(grid_widths)
    excel_widths = emu_to_excel_width(grid_widths, target_width)

    # Set column widths
    for i, w in enumerate(excel_widths):
        col_letter = get_column_letter(i + 1)
        if ws.column_dimensions[col_letter].width is None or ws.column_dimensions[col_letter].width < w:
            ws.column_dimensions[col_letter].width = w

    parsed_rows = parse_table_cells(table)
    vmerge_map = compute_vmerge_ranges(parsed_rows)

    for row_idx, row in enumerate(parsed_rows):
        excel_row = start_row + row_idx
        for cell in row:
            ci = cell['col_idx']
            excel_col = ci + 1  # 1-based
            span = cell['span']

            if cell['is_vmerge_continue']:
                # Part of a vertical merge from above - just apply border
                for c in range(excel_col, excel_col + span):
                    ws.cell(row=excel_row, column=c).border = THIN
                    ws.cell(row=excel_row, column=c).font = F_NORMAL
                continue

            # Determine vertical merge extent
            vmerge_end_row = row_idx
            if cell['is_vmerge_start'] and (row_idx, ci) in vmerge_map:
                vmerge_end_row = vmerge_map[(row_idx, ci)]

            end_excel_row = start_row + vmerge_end_row
            end_excel_col = excel_col + span - 1

            # Merge cells if needed
            if end_excel_row > excel_row or end_excel_col > excel_col:
                try:
                    ws.merge_cells(
                        start_row=excel_row, start_column=excel_col,
                        end_row=end_excel_row, end_column=end_excel_col
                    )
                except Exception:
                    pass  # Already merged or overlap

            # Set value and style
            text = cell['text']
            c = ws.cell(row=excel_row, column=excel_col)
            c.value = text if text else None
            c.font = F_NORMAL
            c.alignment = AL_C
            c.border = THIN

            # Apply borders to all cells in the merge range
            for rr in range(excel_row, end_excel_row + 1):
                for cc in range(excel_col, end_excel_col + 1):
                    ws.cell(row=rr, column=cc).border = THIN
                    ws.cell(row=rr, column=cc).font = F_NORMAL

        # Set row height
        # Check if any cell in this row has multi-line text
        max_lines = 1
        for cell in row:
            if not cell['is_vmerge_continue']:
                lines = cell['text'].count('\n') + 1 if cell['text'] else 1
                max_lines = max(max_lines, lines)
        ws.row_dimensions[excel_row].height = max(15, min(max_lines * 13, 60))

    return start_row + len(parsed_rows)


def write_paragraphs_to_sheet(ws, paragraphs, start_row, max_col=None):
    """Write paragraph texts to sheet. Returns next available row."""
    row = start_row
    for p in paragraphs:
        text = p.text.strip() if hasattr(p, 'text') else str(p).strip()
        if text:
            c = ws.cell(row=row, column=1, value=text)
            c.font = F_NORMAL
            c.alignment = AL_L
            if max_col and max_col > 1:
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
            row += 1
    return row


def setup_print_area(ws, max_row, max_col):
    """Set print area and page setup for A4."""
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.print_area = f'A1:{get_column_letter(max_col)}{max_row}'


# ============================================================
# Source files
# ============================================================
FILES = {
    '着工届': '/root/.claude/uploads/1aedaef9-61f4-5511-9d57-68b254cd1569/9b0848a4-_____.docx',
    '設置届': '/root/.claude/uploads/1aedaef9-61f4-5511-9d57-68b254cd1569/a3c02a33-_____.docx',
    '概要表': '/tmp/docconv/dc729bb0-_____.docx',
    '様式11': '/tmp/docconv/18d9644f-sikenkekkahoukoku11.docx',
    '様式34': '/tmp/docconv/5b588e9e-__sikenkekkahoukoku34.docx',
}

OUTPUT = '/home/user/claude-project/消防設備工事_様式統合.xlsx'


def main():
    wb = openpyxl.Workbook()

    # ============================================================
    # SHEET 1: 入力シート
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
        ('担当消防署（例：○○消防署長　殿）', None),
        ('届出日', None),
        ('届出者 住所', None),
        ('届出者 氏名（管理組合名等）', None),
        ('工事場所（住所）', None),
        ('防火対象物の名称', None),
        ('設備の種類', '自動火災報知設備'),
        ('用途（例：5項ロ）', None),
        ('構造', '耐火構造'),
        ('地上階数', None),
        ('地下階数', None),
        ('延べ面積（㎡）', None),
        ('着工予定日', None),
        ('完成予定日', None),
        ('理事長名（設置届用）', None),
        ('全住戸数', None),
    ]

    for i, (label, default) in enumerate(fields):
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

    setup_print_area(ws, 92, 7)

    # ============================================================
    # SHEET 2: 着工届出書
    # ============================================================
    print("Processing 着工届出書...")
    doc = Document(FILES['着工届'])
    ws2 = wb.create_sheet('着工届出書')

    # Title paragraphs before table
    pre_paras = []
    for elem in doc.element.body:
        if elem.tag == qn('w:tbl'):
            break
        if elem.tag == qn('w:p'):
            text = ''.join(t.text or '' for t in elem.iter() if t.tag == qn('w:t'))
            if text.strip():
                pre_paras.append(text.strip())

    row = 1
    for text in pre_paras:
        c = ws2.cell(row=row, column=1, value=text)
        c.font = F_NORMAL
        c.alignment = AL_L
        row += 1
    if pre_paras:
        row += 1  # blank row

    next_row = write_table_to_sheet(ws2, doc.tables[0], start_row=row)

    # Post-table paragraphs (備考)
    post_paras = []
    found_table = False
    for elem in doc.element.body:
        if elem.tag == qn('w:tbl'):
            found_table = True
            continue
        if found_table and elem.tag == qn('w:p'):
            text = ''.join(t.text or '' for t in elem.iter() if t.tag == qn('w:t'))
            if text.strip():
                post_paras.append(text.strip())

    next_row += 1
    for text in post_paras:
        c = ws2.cell(row=next_row, column=1, value=text)
        c.font = F_SMALL
        c.alignment = AL_L
        next_row += 1

    grid_widths = get_grid_col_widths(doc.tables[0])
    setup_print_area(ws2, next_row, len(grid_widths))

    # ============================================================
    # SHEET 3: 設置届出書
    # ============================================================
    print("Processing 設置届出書...")
    doc = Document(FILES['設置届'])
    ws3 = wb.create_sheet('設置届出書')

    pre_paras = []
    for elem in doc.element.body:
        if elem.tag == qn('w:tbl'):
            break
        if elem.tag == qn('w:p'):
            text = ''.join(t.text or '' for t in elem.iter() if t.tag == qn('w:t'))
            if text.strip():
                pre_paras.append(text.strip())

    row = 1
    for text in pre_paras:
        c = ws3.cell(row=row, column=1, value=text)
        c.font = F_NORMAL
        c.alignment = AL_L
        row += 1
    if pre_paras:
        row += 1

    next_row = write_table_to_sheet(ws3, doc.tables[0], start_row=row)

    post_paras = []
    found_table = False
    for elem in doc.element.body:
        if elem.tag == qn('w:tbl'):
            found_table = True
            continue
        if found_table and elem.tag == qn('w:p'):
            text = ''.join(t.text or '' for t in elem.iter() if t.tag == qn('w:t'))
            if text.strip():
                post_paras.append(text.strip())

    next_row += 1
    for text in post_paras:
        c = ws3.cell(row=next_row, column=1, value=text)
        c.font = F_SMALL
        c.alignment = AL_L
        next_row += 1

    grid_widths = get_grid_col_widths(doc.tables[0])
    setup_print_area(ws3, next_row, len(grid_widths))

    # ============================================================
    # SHEET 4 & 5: 概要表（その1）（その2）
    # ============================================================
    print("Processing 概要表...")
    doc = Document(FILES['概要表'])

    # Collect paragraphs between elements
    body_elements = list(doc.element.body)
    table_indices = [i for i, elem in enumerate(body_elements) if elem.tag == qn('w:tbl')]

    for tbl_num, sheet_name in enumerate(['概要表（その1）', '概要表（その2）']):
        ws_g = wb.create_sheet(sheet_name)

        # Find paragraphs before this table
        if tbl_num < len(table_indices):
            tbl_idx = table_indices[tbl_num]
            # Look back for paragraphs
            start_search = table_indices[tbl_num - 1] + 1 if tbl_num > 0 else 0
            pre_texts = []
            for i in range(start_search, tbl_idx):
                elem = body_elements[i]
                if elem.tag == qn('w:p'):
                    text = ''.join(t.text or '' for t in elem.iter() if t.tag == qn('w:t'))
                    if text.strip():
                        pre_texts.append(text.strip())

            row = 1
            for text in pre_texts:
                c = ws_g.cell(row=row, column=1, value=text)
                c.font = F_NORMAL
                c.alignment = AL_L
                row += 1
            if pre_texts:
                row += 1

            next_row = write_table_to_sheet(ws_g, doc.tables[tbl_num], start_row=row)
            grid_widths = get_grid_col_widths(doc.tables[tbl_num])
            setup_print_area(ws_g, next_row, len(grid_widths))

    # Post-table paragraphs (備考) on その2
    if len(table_indices) >= 2:
        last_tbl_idx = table_indices[-1]
        post_texts = []
        for i in range(last_tbl_idx + 1, len(body_elements)):
            elem = body_elements[i]
            if elem.tag == qn('w:p'):
                text = ''.join(t.text or '' for t in elem.iter() if t.tag == qn('w:t'))
                if text.strip():
                    post_texts.append(text.strip())
        if post_texts:
            next_row += 1
            for text in post_texts:
                c = ws_g.cell(row=next_row, column=1, value=text)
                c.font = F_SMALL
                c.alignment = AL_L
                next_row += 1

    # ============================================================
    # SHEET 6: 試験結果(様式11)
    # ============================================================
    print("Processing 様式11...")
    doc = Document(FILES['様式11'])
    ws11 = wb.create_sheet('試験結果(様式11)')

    # Get all paragraphs and tables in order
    body_elements = list(doc.element.body)
    row = 1

    # Process all elements in order
    current_table_idx = 0
    for elem in body_elements:
        if elem.tag == qn('w:p'):
            text = ''.join(t.text or '' for t in elem.iter() if t.tag == qn('w:t'))
            if text.strip():
                c = ws11.cell(row=row, column=1, value=text.strip())
                c.font = F_NORMAL
                c.alignment = AL_L
                row += 1
        elif elem.tag == qn('w:tbl') and current_table_idx < len(doc.tables):
            table = doc.tables[current_table_idx]
            # For multi-table sheets, we need to handle different column counts
            # Write each table starting at current row
            grid_widths = get_grid_col_widths(table)
            num_cols = len(grid_widths)

            # Set column widths for this table (max across all tables)
            excel_widths = emu_to_excel_width(grid_widths, 85)
            for i, w in enumerate(excel_widths):
                col_letter = get_column_letter(i + 1)
                existing = ws11.column_dimensions[col_letter].width
                if existing is None or existing < w:
                    ws11.column_dimensions[col_letter].width = w

            row = write_table_to_sheet(ws11, table, start_row=row, target_width=85)
            row += 1  # blank row between tables
            current_table_idx += 1

    # Find max columns across all tables
    max_cols = max(len(get_grid_col_widths(t)) for t in doc.tables)
    setup_print_area(ws11, row, max_cols)

    # ============================================================
    # SHEET 7: 試験結果(様式34)
    # ============================================================
    print("Processing 様式34...")
    doc = Document(FILES['様式34'])
    ws34 = wb.create_sheet('試験結果(様式34)')

    body_elements = list(doc.element.body)
    row = 1
    current_table_idx = 0

    for elem in body_elements:
        if elem.tag == qn('w:p'):
            text = ''.join(t.text or '' for t in elem.iter() if t.tag == qn('w:t'))
            if text.strip():
                c = ws34.cell(row=row, column=1, value=text.strip())
                c.font = F_NORMAL
                c.alignment = AL_L
                row += 1
        elif elem.tag == qn('w:tbl') and current_table_idx < len(doc.tables):
            table = doc.tables[current_table_idx]
            grid_widths = get_grid_col_widths(table)
            excel_widths = emu_to_excel_width(grid_widths, 85)
            for i, w in enumerate(excel_widths):
                col_letter = get_column_letter(i + 1)
                existing = ws34.column_dimensions[col_letter].width
                if existing is None or existing < w:
                    ws34.column_dimensions[col_letter].width = w

            row = write_table_to_sheet(ws34, table, start_row=row, target_width=85)
            row += 1
            current_table_idx += 1

    max_cols = max(len(get_grid_col_widths(t)) for t in doc.tables)
    setup_print_area(ws34, row, max_cols)

    # ============================================================
    # Save
    # ============================================================
    wb.save(OUTPUT)
    print(f"Saved to {OUTPUT}")
    print(f"Sheets: {wb.sheetnames}")


if __name__ == '__main__':
    main()
