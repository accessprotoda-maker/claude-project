#!/usr/bin/env python3
"""入居者一覧(.xls)から住戸内工事日程表(.xlsx)を生成する。

使い方:
    python3 generate_kouji_schedule.py 入居者一覧.xls 工程表.xlsx
"""
import re
import sys

import xlrd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

TIME_SLOTS = [9, 10, 11, 13, 14, 15, 16, 17]
AM_SLOTS = [9, 10, 11]
PM_SLOTS = [13, 14, 15, 16, 17]
TOTAL_COLS = 1 + len(AM_SLOTS) + len(PM_SLOTS)  # 工事予定日 + 8時間帯

DATE_RE = re.compile(r"(\d{1,2})[/月](\d{1,2})日?[（(](.)[）)]\s*(\d{1,2})[:：](\d{2})")

THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
YELLOW = PatternFill("solid", fgColor="FFFF00")
ORANGE = PatternFill("solid", fgColor="FFC000")
GRAY = PatternFill("solid", fgColor="D9D9D9")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

WEEKDAYS = ["月", "火", "水", "木", "金", "土", "日"]


def common_area_date(month, day, weekday):
    """住戸内工事の初日の前日（共用部工事日）を求める。"""
    import datetime
    d = datetime.date(2001, month, day) - datetime.timedelta(days=1)
    prev_weekday = WEEKDAYS[(WEEKDAYS.index(weekday) - 1) % 7]
    return d.month, d.day, prev_weekday


def option_marker(remark):
    """備考からインターホン受話器・カメラ付きオプションの表示を作る。"""
    if not remark:
        return None
    suffix = ""
    if "カメラ" in remark:
        suffix += "A"
    if "受話器" in remark:
        suffix += "B"
    return suffix or None


def format_room(room):
    if isinstance(room, float) and room == int(room):
        return str(int(room))
    return str(room).strip()


def parse_schedule(text):
    m = DATE_RE.search(text)
    if not m:
        return None
    month, day, weekday, hour, minute = m.groups()
    return {
        "month": int(month),
        "day": int(day),
        "weekday": weekday,
        "hour": int(hour),
        "minute": int(minute),
    }


def load_rooms(path_or_bytes):
    if isinstance(path_or_bytes, bytes):
        wb_in = xlrd.open_workbook(file_contents=path_or_bytes)
    else:
        wb_in = xlrd.open_workbook(path_or_bytes)
    building_name = wb_in.sheet_by_index(0).cell_value(0, 0)

    rooms = []
    for sheet in wb_in.sheets():
        if sheet.nrows < 3:
            continue
        for r in range(2, sheet.nrows):
            room = sheet.cell_value(r, 0)
            if not room:
                continue
            name = sheet.cell_value(r, 1)
            schedule_text = sheet.cell_value(r, 4)
            remark = sheet.cell_value(r, 5)
            rooms.append({
                "room": format_room(room),
                "name": name,
                "schedule": parse_schedule(schedule_text) if schedule_text else None,
                "schedule_text": schedule_text,
                "remark": remark,
                "option": option_marker(remark),
            })
    return building_name, rooms


def build_table_data(rooms):
    dates = {}  # (month, day) -> {"weekday": str, "slots": {hour: [(room, minute)]}}
    out_of_period = []
    vacant = []
    not_submitted = []

    for r in rooms:
        if r["name"] == "空室":
            vacant.append(r["room"])
            continue
        if r["remark"] and "工期外" in r["remark"]:
            out_of_period.append(r)
            continue
        sched = r["schedule"]
        if not sched:
            not_submitted.append(r["room"])
            continue
        key = (sched["month"], sched["day"])
        d = dates.setdefault(key, {"weekday": sched["weekday"],
                                    "slots": {h: [] for h in TIME_SLOTS}})
        d["slots"].setdefault(sched["hour"], []).append((r["room"], sched["minute"], r["option"]))

    return dates, out_of_period, vacant, not_submitted


def generate_workbook(path_or_bytes):
    """入居者一覧(.xls)から工程表のWorkbookを生成する。"""
    building_name, rooms = load_rooms(path_or_bytes)
    dates, out_of_period, vacant, not_submitted = build_table_data(rooms)

    wb = Workbook()
    ws = wb.active
    ws.title = "工程表"

    # タイトル
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TOTAL_COLS)
    title_cell = ws.cell(row=1, column=1, value=f"{building_name}　様　住戸内工事日程表")
    title_cell.font = Font(size=16, bold=True)
    title_cell.alignment = CENTER
    for c in range(1, TOTAL_COLS + 1):
        ws.cell(row=1, column=c).border = BORDER
    ws.row_dimensions[1].height = 30

    # 注記
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=TOTAL_COLS)
    note_cell = ws.cell(row=2, column=1, value="※表中の部屋数は工事可能な部屋数を示します。")
    note_cell.alignment = LEFT

    header_row1 = 3
    header_row2 = 4

    # 工事予定日 (2行分マージ)
    ws.merge_cells(start_row=header_row1, start_column=1,
                    end_row=header_row2, end_column=1)
    c = ws.cell(row=header_row1, column=1, value="工事予定日")
    c.font = Font(bold=True)
    c.alignment = CENTER

    # 午前
    ws.merge_cells(start_row=header_row1, start_column=2,
                    end_row=header_row1, end_column=1 + len(AM_SLOTS))
    c = ws.cell(row=header_row1, column=2, value="午前（9時～12時）")
    c.font = Font(bold=True)
    c.alignment = CENTER

    # 午後
    pm_start = 2 + len(AM_SLOTS)
    ws.merge_cells(start_row=header_row1, start_column=pm_start,
                    end_row=header_row1, end_column=TOTAL_COLS)
    c = ws.cell(row=header_row1, column=pm_start, value="午後（13時～18時）")
    c.font = Font(bold=True)
    c.alignment = CENTER

    # 時間帯ラベル
    for i, h in enumerate(TIME_SLOTS):
        col = 2 + i
        c = ws.cell(row=header_row2, column=col, value=f"{h}時頃")
        c.font = Font(bold=True)
        c.alignment = CENTER

    for c in range(1, TOTAL_COLS + 1):
        ws.cell(row=header_row1, column=c).border = BORDER
        ws.cell(row=header_row2, column=c).border = BORDER

    # データ行
    row = header_row2 + 1

    # 初日の前日を共用部工事日として追加
    if dates:
        (first_month, first_day), first_info = sorted(dates.items())[0]
        ca_month, ca_day, ca_weekday = common_area_date(first_month, first_day, first_info["weekday"])

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=1)
        date_cell = ws.cell(row=row, column=1, value=f"{ca_month}月{ca_day}日（{ca_weekday}）")
        date_cell.font = Font(bold=True)
        date_cell.alignment = CENTER

        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=TOTAL_COLS)
        ca_cell = ws.cell(row=row, column=2, value="共　用　部　工　事\n（お部屋の工事は出来ません）")
        ca_cell.font = Font(bold=True, size=12)
        ca_cell.alignment = CENTER
        ca_cell.fill = GRAY

        for c in range(1, TOTAL_COLS + 1):
            ws.cell(row=row, column=c).border = BORDER
            ws.cell(row=row, column=c).fill = GRAY
        ws.row_dimensions[row].height = 40
        row += 1

    for (month, day), info in sorted(dates.items()):
        slots = info["slots"]
        weekday = info["weekday"]
        n_rows = max(1, max(len(slots.get(h, [])) for h in TIME_SLOTS))

        # 日付セル
        ws.merge_cells(start_row=row, start_column=1,
                        end_row=row + n_rows - 1, end_column=1)
        date_cell = ws.cell(row=row, column=1, value=f"{month}月{day}日（{weekday}）")
        date_cell.font = Font(bold=True)
        date_cell.alignment = CENTER

        for i, h in enumerate(TIME_SLOTS):
            col = 2 + i
            entries = sorted(slots.get(h, []), key=lambda e: e[1])
            for sub in range(n_rows):
                cell = ws.cell(row=row + sub, column=col)
                if sub < len(entries):
                    room, minute, option = entries[sub]
                    if option:
                        cell.value = f"{h}:{minute:02d}\n{room}{option}"
                        cell.fill = ORANGE
                    else:
                        cell.value = f"{h}:{minute:02d}\n{room}"
                        cell.fill = YELLOW
                    cell.font = Font(bold=True)
                cell.alignment = CENTER

        for r2 in range(row, row + n_rows):
            for c in range(1, TOTAL_COLS + 1):
                ws.cell(row=r2, column=c).border = BORDER

        row += n_rows

    # 列幅・行高
    ws.column_dimensions["A"].width = 16
    for i in range(len(TIME_SLOTS)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 10
    for r2 in range(header_row2 + 1, row):
        ws.row_dimensions[r2].height = 30

    # オプション（インターホン受話器・カメラ付き）凡例
    has_option = any(opt for d in dates.values() for entries in d["slots"].values()
                      for _, _, opt in entries)
    if has_option:
        cell = ws.cell(row=row, column=1, value="A：カメラ付き　B：受話器付きのお部屋です")
        cell.fill = ORANGE
        cell.font = Font(bold=True)
        row += 1

    # 工期外希望
    row += 1
    for r in out_of_period:
        sched = r["schedule"]
        if sched:
            cell = ws.cell(
                row=row, column=1,
                value=(f"{r['room']} 工期外希望"
                       f"（{sched['month']}/{sched['day']}{sched['weekday']}"
                       f"{sched['hour']}:{sched['minute']:02d}）"))
        else:
            cell = ws.cell(row=row, column=1, value=f"{r['room']} 工期外希望")
        cell.font = Font(color="0000FF", bold=True)
        row += 1

    # 未提出・空室
    if not_submitted:
        cell = ws.cell(row=row, column=1, value=f"未提出　{', '.join(not_submitted)}")
        cell.font = Font(color="FF0000", bold=True)
        row += 1
    if vacant:
        cell = ws.cell(row=row, column=1, value=f"空室　{', '.join(vacant)}")
        cell.font = Font(color="FF0000", bold=True)
        row += 1

    # A4用紙1枚に収まるよう印刷設定
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4
    ws.print_area = f"A1:{get_column_letter(TOTAL_COLS)}{row - 1}"

    return wb


def main():
    in_path = sys.argv[1] if len(sys.argv) > 1 else "入居者一覧.xls"
    out_path = sys.argv[2] if len(sys.argv) > 2 else "工程表.xlsx"

    wb = generate_workbook(in_path)
    wb.save(out_path)
    print(f"saved: {out_path}")


if __name__ == "__main__":
    main()
