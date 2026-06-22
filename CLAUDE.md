# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Fire protection equipment construction document tools for Japanese fire departments (消防設備工事書類ツール). All UI text, form labels, and document output are in Japanese.

Two main capabilities:
1. **工程表ジェネレーター** — Streamlit web app that converts a tenant list (.xls) into a construction schedule (.xlsx)
2. **様式統合Excel生成** — Python scripts that produce a single Excel workbook consolidating official fire department form templates (別記様式) with a shared input sheet

## Running the App

```bash
pip install -r requirements.txt          # streamlit, openpyxl, xlrd
streamlit run app.py                     # launches 工程表ジェネレーター
```

For form generation scripts (standalone, no server):
```bash
pip install python-docx                  # additional dependency for Word conversion
python3 create_forms_v2.py               # generates 消防設備工事_様式統合.xlsx
```

## Architecture

### 工程表ジェネレーター (Schedule Generator)
- `app.py` — Streamlit entry point; uploads .xls, calls generator, offers .xlsx download
- `tools/generate_kouji_schedule.py` — Core logic: parses tenant list with xlrd, generates formatted schedule with openpyxl. Also works as CLI: `python3 tools/generate_kouji_schedule.py input.xls output.xlsx`
- `gas/GenerateKoujiSchedule.gs` — Google Apps Script port (same logic for Google Sheets)
- `vba/GenerateKoujiSchedule.bas` — VBA port (same logic for Excel macros)

### 様式統合 (Form Integration)
- `create_forms_v2.py` — **Current version**. Generates a 9-sheet Excel workbook: 入力シート (input) + 着工届出書 (hand-crafted to match official PDF layout) + 7 auto-converted form sheets. Uses python-docx XML parsing for Word→Excel conversion.
- `convert_word_to_excel.py` — Previous auto-conversion approach (superseded by v2)
- `create_form_exact.py` / `create_integrated_form.py` — Earlier iterations (superseded)

Source Word documents are in `/root/.claude/uploads/` and `/tmp/docconv/` (LibreOffice-converted .doc→.docx).

## Key Patterns

- All Excel generation uses openpyxl with ＭＳ ゴシック font
- Form sheets reference 入力シート via Excel formulas (`=入力シート!B9`, `=VLOOKUP(入力シート!B34,...)`)
- Word table parsing uses direct XML access (`table._tbl`, `qn('w:gridCol')`, `qn('w:gridSpan')`, `qn('w:vMerge')`) rather than python-docx's high-level API, because the high-level API doesn't accurately report merged cells
- Column widths from Word EMU values are converted proportionally with a minimum of 3 chars to prevent invisible columns

## Language Note

All commit messages, file names, sheet names, and cell content should be in Japanese to match the domain. Code comments and variable names may be in English or Japanese.
