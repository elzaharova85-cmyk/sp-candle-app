import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
import os, io, zipfile, tempfile, json
from datetime import datetime

# ─────────────────────────────────────────────
#  СЕБЕСТОИМОСТЬ  (по типу товара)
# ─────────────────────────────────────────────
DEFAULT_SEBES = {
    'автодиффузор 7мл':            66,
    'диффузор-спрей 100мл':       327,
    'диффузор 100мл':              327,
    'диффузор 50мл':               198,
    'диффузор 500мл':             1264,
    'диффузор 250мл':              660,
    'диффузор 30мл':               288,
    'тестер 30мл':                 288,
    'срезанный флакон 30мл':       288,
    'свеча 320г':                 1334,
    'свеча 220г':                  372,
    'свеча 100г':                  174,
    'крем для рук дой-пак':        793,
    'крем для рук 1л':             793,
    'крем для рук 45мл':           191,
    'крем для рук 250мл':          290,
    'подарочная коробка бордовая': 245,
    'фибровые палочки 3/20':        20,
    'тестер шайба':                 31,
    'спички каминные':              32.5,
}

SEBES_FILE = os.path.join(os.path.dirname(__file__), 'sebes.json')

def load_sebes() -> dict:
    """Загружает себестоимость из файла (если есть) или возвращает дефолтную."""
    if os.path.exists(SEBES_FILE):
        try:
            with open(SEBES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return dict(DEFAULT_SEBES)

def save_sebes(d: dict):
    """Сохраняет себестоимость в файл."""
    with open(SEBES_FILE, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

def get_sebes(name: str, sebes_dict: dict = None) -> float | None:
    """Возвращает себестоимость по наименованию товара."""
    if not isinstance(name, str):
        return None
    d = sebes_dict if sebes_dict is not None else load_sebes()
    n = name.lower()
    for key, val in d.items():
        if key in n:
            return val
    return None


# ─────────────────────────────────────────────
#  ПАРСИНГ ФАЙЛОВ
# ─────────────────────────────────────────────

def parse_1c_file(file_obj, channel_name: str, sebes_dict: dict) -> pd.DataFrame:
    import xlrd
    with tempfile.NamedTemporaryFile(suffix='.xls', delete=False) as tmp:
        tmp.write(file_obj.read())
        tmp_path = tmp.name

    df_raw = pd.read_excel(tmp_path, engine='xlrd', header=None, skiprows=11)
    os.unlink(tmp_path)

    header_row = None
    for i, row in df_raw.iterrows():
        vals = [str(v).strip() for v in row.values if pd.notna(v)]
        if any('наименование' in v.lower() for v in vals):
            header_row = i
            break

    if header_row is None:
        for i, row in df_raw.iterrows():
            vals = [str(v).strip() for v in row.values if pd.notna(v)]
            if any('кол-во' in v.lower() for v in vals):
                header_row = i
                break

    data_start = (header_row + 2) if header_row is not None else 3

    rows = []
    for i in range(data_start, len(df_raw)):
        row = df_raw.iloc[i]
        name = row.iloc[3] if len(row) > 3 else None
        qty  = row.iloc[5] if len(row) > 5 else None
        price= row.iloc[6] if len(row) > 6 else None
        sebes_1c = row.iloc[7] if len(row) > 7 else None
        sum_sales = row.iloc[9] if len(row) > 9 else None
        sum_sebes = row.iloc[10] if len(row) > 10 else None
        profit = row.iloc[13] if len(row) > 13 else None

        if not isinstance(name, str) or len(name.strip()) < 2:
            continue
        if name.strip().lower() in ('наименование', 'итого', ''):
            continue
        try:
            qty_f = float(qty)
        except (TypeError, ValueError):
            continue

        rows.append({
            'Канал':           channel_name,
            'Наименование':    name.strip(),
            'Кол-во':          qty_f,
            'Цена':            _to_float(price),
            'Себес_1С':        _to_float(sebes_1c),
            'Себес_авто':      get_sebes(name, sebes_dict),
            'Сумма продаж':    _to_float(sum_sales),
            'Сумма себест':    _to_float(sum_sebes),
            'Прибыль':         _to_float(profit),
        })

    df = pd.DataFrame(rows)
    return df


def parse_ozon_file(file_obj) -> dict:
    xl = pd.ExcelFile(file_obj)
    result = {}

    for sh in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sh, header=None)
        if 'Сводка' in sh or sh == 'Сводка':
            result['сводка'] = df
        elif 'Детализация' in sh or sh == 'Детализация':
            df2 = pd.read_excel(xl, sheet_name=sh, skiprows=0)
            df2.columns = df2.iloc[0]
            df2 = df2.iloc[1:].reset_index(drop=True)
            result['детализация'] = df2
        elif 'типам' in sh.lower():
            result['по_типам'] = df

    return result


def parse_wb_files(zip_files, date_from=None, date_to=None) -> dict:
    """
    Принимает список zip-файлов WB. Извлекает XLSX и объединяет.
    date_from / date_to — объекты datetime.date для фильтрации периода.
    """
    all_dfs = []
    for zf in zip_files:
        zf.seek(0)
        with zipfile.ZipFile(zf) as z:
            for name in z.namelist():
                if name.endswith('.xlsx'):
                    with z.open(name) as f:
                        df = pd.read_excel(f, sheet_name='Sheet1')
                        all_dfs.append(df)

    if not all_dfs:
        return {}

    combined = pd.concat(all_dfs, ignore_index=True)

    # Фильтрация по дате продажи
    if date_from or date_to:
        date_col = None
        for cname in combined.columns:
            if 'дата продажи' in str(cname).lower() or 'дата заказа' in str(cname).lower():
                date_col = cname
                break
        if date_col:
            combined[date_col] = pd.to_datetime(combined[date_col], errors='coerce')
            if date_from:
                combined = combined[combined[date_col].dt.date >= date_from]
            if date_to:
                combined = combined[combined[date_col].dt.date <= date_to]

    return {'детализация': combined}


def parse_ms_summary(file_obj) -> dict:
    xl = pd.ExcelFile(file_obj)
    result = {}
    for sh in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sh, header=None)
        result[sh] = df
    return result


def _to_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ─────────────────────────────────────────────
#  СТИЛИ EXCEL
# ─────────────────────────────────────────────
COLORS = {
    'header_dark':  '1F3864',
    'header_blue':  '2E75B6',
    'header_green': '375623',
    'header_red':   '843C0C',
    'ozon_blue':    '005BFF',
    'wb_violet':    '7B2FBE',
    'row_white':    'FFFFFF',   # все строки белые
    'total':        'BDD7EE',
    'grand_total':  '1F3864',
    'yellow':       'FFEB9C',
}

def hfill(color): return PatternFill('solid', fgColor=color)
def hfont(bold=False, white=False, size=10):
    return Font(bold=bold, color='FFFFFF' if white else '000000', size=size, name='Calibri')
def center(): return Alignment(horizontal='center', vertical='center', wrap_text=True)
def left():   return Alignment(horizontal='left',   vertical='center', wrap_text=True)
def right():  return Alignment(horizontal='right',  vertical='center')

BORDER_THIN = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'),  bottom=Side(style='thin')
)

def _set_col_widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

def _row_fill():
    """Все строки данных — белые."""
    return hfill(COLORS['row_white'])


# ─────────────────────────────────────────────
#  ПОСТРОЕНИЕ EXCEL-ОТЧЁТА
# ─────────────────────────────────────────────

def build_excel(ms_data: dict, ozon_data: dict, wb_data: dict,
                period: str, sebes_dict: dict) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    ws_svod = wb.create_sheet('СВОДНЫЙ')
    _build_svod_sheet(ws_svod, ms_data, ozon_data, wb_data, period, sebes_dict)

    channel_map = {
        'Интернет':   [k for k in ms_data if 'интернет' in k.lower() or 'Интернет' in k],
        'ОПТ':        [k for k in ms_data if 'опт' in k.lower() and 'свод' not in k.lower()],
        'СТМ':        [k for k in ms_data if 'стм' in k.lower() or 'сзм' in k.lower()],
        'Столярный':  [k for k in ms_data if 'столяр' in k.lower()],
    }

    for ch_name, keys in channel_map.items():
        if not keys:
            continue
        key = keys[0]
        ws_ch = wb.create_sheet(f'МС_{ch_name}')
        _build_channel_sheet(ws_ch, ms_data[key], ch_name, period, sebes_dict)

    if ozon_data:
        ws_ozon_svod = wb.create_sheet('Ozon_Сводка')
        _build_ozon_svod(ws_ozon_svod, ozon_data, period)

        if 'детализация' in ozon_data:
            ws_ozon_det = wb.create_sheet('Ozon_Детализация')
            _build_ozon_detail(ws_ozon_det, ozon_data['детализация'], sebes_dict)

        if 'по_типам' in ozon_data:
            ws_ozon_types = wb.create_sheet('Ozon_По типам')
            _build_raw_sheet(ws_ozon_types, ozon_data['по_типам'],
                             'Продажи Ozon по типам товара', COLORS['ozon_blue'])

    if wb_data and 'детализация' in wb_data:
        ws_wb_svod = wb.create_sheet('WB_Сводка')
        _build_wb_svod(ws_wb_svod, wb_data['детализация'], period, sebes_dict)

        ws_wb_det = wb.create_sheet('WB_Детализация')
        _build_wb_detail(ws_wb_det, wb_data['детализация'], sebes_dict)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
#  ЛИСТ: СВОДНЫЙ
# ─────────────────────────────────────────────
def _build_svod_sheet(ws, ms_data, ozon_data, wb_data, period, sebes_dict):
    ws.sheet_view.showGridLines = False

    ws.merge_cells('A1:H1')
    c = ws['A1']
    c.value = f'СВОДНЫЙ ОТЧЁТ SP CANDLE — {period.upper()}'
    c.fill = hfill(COLORS['header_dark'])
    c.font = hfont(bold=True, white=True, size=13)
    c.alignment = center()
    ws.row_dimensions[1].height = 30

    row = 3

    ws.merge_cells(f'A{row}:H{row}')
    c = ws[f'A{row}']
    c.value = '1С — ПРОДАЖИ ПО КАНАЛАМ'
    c.fill = hfill(COLORS['header_blue'])
    c.font = hfont(bold=True, white=True)
    c.alignment = center()
    row += 1

    headers = ['Канал', 'Кол-во', 'Выручка', 'Себес итого', 'Прибыль', 'Маржа %']
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col_idx, value=h)
        c.fill = hfill('BDD7EE')
        c.font = hfont(bold=True)
        c.alignment = center()
        c.border = BORDER_THIN
    row += 1

    ms_totals = {'кол': 0, 'выр': 0, 'себ': 0, 'приб': 0}
    svod_key = next((k for k in ms_data if 'свод' in k.lower()), None)

    if svod_key is not None:
        df_svod = ms_data[svod_key]
        data_rows = []
        for i, r in df_svod.iterrows():
            name = str(r.iloc[0]).strip() if pd.notna(r.iloc[0]) else ''
            qty  = r.iloc[1] if len(r) > 1 else None
            rev  = r.iloc[2] if len(r) > 2 else None
            sebes= r.iloc[4] if len(r) > 4 else None
            prof = r.iloc[5] if len(r) > 5 else None
            marg = r.iloc[6] if len(r) > 6 else None
            if name in ('', 'Канал', 'ИТОГО', 'nan') or 'основные' in name.lower():
                continue
            try:
                float(qty)
            except (TypeError, ValueError):
                continue
            data_rows.append((name, qty, rev, sebes, prof, marg))

        for i, (name, qty, rev, sebes, prof, marg) in enumerate(data_rows):
            vals = [name, qty, rev, sebes, prof, marg]
            for col_idx, v in enumerate(vals, 1):
                c = ws.cell(row=row, column=col_idx, value=_to_float(v) if col_idx > 1 else v)
                c.fill = _row_fill()
                c.border = BORDER_THIN
                c.alignment = left() if col_idx == 1 else right()
                if col_idx in (3, 4, 5):
                    c.number_format = '#,##0.00'
                if col_idx == 6:
                    c.number_format = '0.0%'
            ms_totals['кол']  += _to_float(qty)
            ms_totals['выр']  += _to_float(rev)
            ms_totals['себ']  += _to_float(sebes)
            ms_totals['приб'] += _to_float(prof)
            row += 1

    itog_vals = ['ИТОГО МС (1С)', ms_totals['кол'], ms_totals['выр'],
                 ms_totals['себ'], ms_totals['приб'],
                 ms_totals['приб'] / ms_totals['выр'] if ms_totals['выр'] else 0]
    for col_idx, v in enumerate(itog_vals, 1):
        c = ws.cell(row=row, column=col_idx, value=v)
        c.fill = hfill(COLORS['total'])
        c.font = hfont(bold=True)
        c.border = BORDER_THIN
        c.alignment = left() if col_idx == 1 else right()
        if col_idx in (3, 4, 5):
            c.number_format = '#,##0.00'
        if col_idx == 6:
            c.number_format = '0.0%'
    row += 2

    # ── Блок Ozon ──
    ws.merge_cells(f'A{row}:H{row}')
    c = ws[f'A{row}']
    c.value = 'OZON — НАЧИСЛЕНИЯ'
    c.fill = hfill(COLORS['ozon_blue'])
    c.font = hfont(bold=True, white=True)
    c.alignment = center()
    row += 1

    ozon_summary = _extract_ozon_summary(ozon_data)
    ozon_headers = ['Статья', 'Сумма, руб.']
    for col_idx, h in enumerate(ozon_headers, 1):
        c = ws.cell(row=row, column=col_idx, value=h)
        c.fill = hfill('BDD7EE')
        c.font = hfont(bold=True)
        c.alignment = center()
        c.border = BORDER_THIN
    row += 1

    for i, (stat, val) in enumerate(ozon_summary.items()):
        is_total = 'итого' in stat.lower() or 'перечислению' in stat.lower()
        c1 = ws.cell(row=row, column=1, value=stat)
        c1.fill = hfill(COLORS['total']) if is_total else _row_fill()
        c1.font = hfont(bold=is_total)
        c1.border = BORDER_THIN
        c1.alignment = left()
        c2 = ws.cell(row=row, column=2, value=val)
        c2.fill = hfill(COLORS['total']) if is_total else _row_fill()
        c2.font = hfont(bold=is_total)
        c2.number_format = '#,##0.00'
        c2.border = BORDER_THIN
        c2.alignment = right()
        row += 1
    row += 1

    # ── Блок WB ──
    ws.merge_cells(f'A{row}:H{row}')
    c = ws[f'A{row}']
    c.value = 'WILDBERRIES — НАЧИСЛЕНИЯ'
    c.fill = hfill(COLORS['wb_violet'])
    c.font = hfont(bold=True, white=True)
    c.alignment = center()
    row += 1

    wb_summary = _extract_wb_summary(wb_data)
    for col_idx, h in enumerate(['Статья', 'Сумма, руб.'], 1):
        c = ws.cell(row=row, column=col_idx, value=h)
        c.fill = hfill('BDD7EE')
        c.font = hfont(bold=True)
        c.alignment = center()
        c.border = BORDER_THIN
    row += 1

    for i, (stat, val) in enumerate(wb_summary.items()):
        is_total = 'итого' in stat.lower() or 'перечислению' in stat.lower()
        c1 = ws.cell(row=row, column=1, value=stat)
        c1.fill = hfill(COLORS['total']) if is_total else _row_fill()
        c1.font = hfont(bold=is_total)
        c1.border = BORDER_THIN
        c1.alignment = left()
        c2 = ws.cell(row=row, column=2, value=val)
        c2.fill = hfill(COLORS['total']) if is_total else _row_fill()
        c2.font = hfont(bold=is_total)
        c2.number_format = '#,##0.00'
        c2.border = BORDER_THIN
        c2.alignment = right()
        row += 1

    _set_col_widths(ws, {'A': 38, 'B': 14, 'C': 16, 'D': 16, 'E': 14, 'F': 10, 'G': 12, 'H': 12})


# ─────────────────────────────────────────────
#  ЛИСТ: МС — канал
# ─────────────────────────────────────────────
def _build_channel_sheet(ws, df_raw, channel_name: str, period: str, sebes_dict: dict):
    ws.sheet_view.showGridLines = False

    ws.merge_cells('A1:G1')
    c = ws['A1']
    c.value = f'{channel_name.upper()} — {period.upper()}'
    c.fill = hfill(COLORS['header_blue'])
    c.font = hfont(bold=True, white=True, size=12)
    c.alignment = center()
    ws.row_dimensions[1].height = 25

    headers = ['Наименование', 'Кол-во', 'Выручка', 'Себес 1 ед (план)', 'Себес итого', 'Прибыль', 'Маржа %']
    row = 3
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col_idx, value=h)
        c.fill = hfill('BDD7EE')
        c.font = hfont(bold=True)
        c.alignment = center()
        c.border = BORDER_THIN
    row += 1

    data_rows = _parse_ms_channel_df(df_raw, sebes_dict)
    for i, r in enumerate(data_rows):
        is_group = r.get('is_group', False)
        is_total = r.get('is_total', False)
        if is_group:
            fill_color = COLORS['header_blue']
        elif is_total:
            fill_color = COLORS['total']
        else:
            fill_color = COLORS['row_white']

        vals = [r['name'], r['qty'], r['rev'], r['sebes1'], r['sebes_all'], r['profit'], r['margin']]
        for col_idx, v in enumerate(vals, 1):
            c = ws.cell(row=row, column=col_idx, value=v)
            c.fill = hfill(fill_color)
            c.font = hfont(bold=(is_group or is_total), white=is_group)
            c.border = BORDER_THIN
            c.alignment = left() if col_idx == 1 else right()
            if col_idx in (3, 4, 5, 6) and not is_group:
                c.number_format = '#,##0.00'
            if col_idx == 7 and not is_group:
                c.number_format = '0.0%'
        row += 1

    _set_col_widths(ws, {'A': 35, 'B': 10, 'C': 16, 'D': 18, 'E': 16, 'F': 14, 'G': 10})


def _parse_ms_channel_df(df_raw: pd.DataFrame, sebes_dict: dict) -> list:
    rows = []
    header_row = None
    for i, r in df_raw.iterrows():
        vals = [str(v).strip().lower() for v in r.values if pd.notna(v)]
        if 'наименование' in vals:
            header_row = i
            break
    if header_row is None:
        return rows

    for i in range(header_row + 1, len(df_raw)):
        r = df_raw.iloc[i]
        name = str(r.iloc[0]).strip() if pd.notna(r.iloc[0]) else ''
        if not name or name == 'nan':
            continue

        qty      = r.iloc[1] if len(r) > 1 else None
        rev      = r.iloc[2] if len(r) > 2 else None
        sebes1   = r.iloc[3] if len(r) > 3 else None
        sebes_all= r.iloc[4] if len(r) > 4 else None
        profit   = r.iloc[5] if len(r) > 5 else None
        margin   = r.iloc[6] if len(r) > 6 else None

        is_group = pd.isna(qty) and pd.isna(rev)
        try:
            qty_f = float(qty)
        except (TypeError, ValueError):
            qty_f = None

        is_total = 'итого' in name.lower()

        sebes1_auto = get_sebes(name, sebes_dict)
        sebes1_val = _to_float(sebes1) if pd.notna(sebes1) and _to_float(sebes1) > 0 else (sebes1_auto or 0)

        rows.append({
            'name':      name,
            'qty':       qty_f if not is_group else None,
            'rev':       _to_float(rev) if not is_group else None,
            'sebes1':    sebes1_val if not is_group else None,
            'sebes_all': _to_float(sebes_all) if not is_group else None,
            'profit':    _to_float(profit) if not is_group else None,
            'margin':    _to_float(margin) if not is_group else None,
            'is_group':  is_group,
            'is_total':  is_total,
        })

    return rows


# ─────────────────────────────────────────────
#  ЛИСТ: OZON
# ─────────────────────────────────────────────
def _extract_ozon_summary(ozon_data: dict) -> dict:
    if not ozon_data or 'сводка' not in ozon_data:
        return {}
    df = ozon_data['сводка']
    result = {}
    for i, r in df.iterrows():
        label = str(r.iloc[0]).strip() if pd.notna(r.iloc[0]) else ''
        val   = r.iloc[1] if len(r) > 1 and pd.notna(r.iloc[1]) else None
        if label and label != 'nan' and val is not None:
            try:
                result[label] = float(val)
            except (TypeError, ValueError):
                pass
    return result


def _build_ozon_svod(ws, ozon_data: dict, period: str):
    ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:C1')
    c = ws['A1']
    c.value = f'OZON — СВОДКА НАЧИСЛЕНИЙ — {period.upper()}'
    c.fill = hfill(COLORS['ozon_blue'])
    c.font = hfont(bold=True, white=True, size=12)
    c.alignment = center()
    ws.row_dimensions[1].height = 25

    summary = _extract_ozon_summary(ozon_data)
    row = 3
    for col_idx, h in enumerate(['Статья', 'Сумма, руб.'], 1):
        c = ws.cell(row=row, column=col_idx, value=h)
        c.fill = hfill('BDD7EE')
        c.font = hfont(bold=True)
        c.alignment = center()
        c.border = BORDER_THIN
    row += 1

    for i, (stat, val) in enumerate(summary.items()):
        is_total = any(x in stat.lower() for x in ('итого', 'перечислению', 'начислено'))
        c1 = ws.cell(row=row, column=1, value=stat)
        c1.fill = hfill(COLORS['total']) if is_total else _row_fill()
        c1.font = hfont(bold=is_total)
        c1.border = BORDER_THIN
        c1.alignment = left()
        c2 = ws.cell(row=row, column=2, value=val)
        c2.fill = hfill(COLORS['total']) if is_total else _row_fill()
        c2.font = hfont(bold=is_total)
        c2.number_format = '#,##0.00'
        c2.border = BORDER_THIN
        c2.alignment = right()
        row += 1

    if 'по_типам' in ozon_data:
        row += 1
        df = ozon_data['по_типам']
        ws.merge_cells(f'A{row}:F{row}')
        c = ws[f'A{row}']
        c.value = 'ПРОДАЖИ ПО ТИПАМ ТОВАРА (Ozon)'
        c.fill = hfill(COLORS['ozon_blue'])
        c.font = hfont(bold=True, white=True)
        c.alignment = center()
        row += 1
        _dump_df_to_ws(ws, df, row)

    _set_col_widths(ws, {'A': 45, 'B': 16, 'C': 14, 'D': 14, 'E': 14, 'F': 16})


def _build_ozon_detail(ws, df: pd.DataFrame, sebes_dict: dict):
    ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:J1')
    c = ws['A1']
    c.value = 'OZON — ДЕТАЛИЗАЦИЯ НАЧИСЛЕНИЙ'
    c.fill = hfill(COLORS['ozon_blue'])
    c.font = hfont(bold=True, white=True, size=12)
    c.alignment = center()
    ws.row_dimensions[1].height = 25

    if 'Название товара' in df.columns:
        df = df.copy()
        df['Себес (авто)'] = df['Название товара'].apply(lambda x: get_sebes(x, sebes_dict))

    row = 3
    cols = list(df.columns)
    for col_idx, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=col_idx, value=str(h))
        c.fill = hfill(COLORS['ozon_blue'])
        c.font = hfont(bold=True, white=True)
        c.alignment = center()
        c.border = BORDER_THIN
    row += 1

    for i, (_, dr) in enumerate(df.iterrows()):
        for col_idx, v in enumerate(dr.values, 1):
            v_out = None if pd.isna(v) else v
            c = ws.cell(row=row, column=col_idx, value=v_out)
            c.fill = _row_fill()
            c.border = BORDER_THIN
            c.alignment = left() if col_idx <= 3 else right()
        row += 1

    _set_col_widths(ws, {'A': 18, 'B': 20, 'C': 22, 'D': 25, 'E': 14, 'F': 14,
                         'G': 40, 'H': 10, 'I': 12, 'J': 16})


# ─────────────────────────────────────────────
#  ЛИСТ: WB
# ─────────────────────────────────────────────
def _extract_wb_summary(wb_data: dict) -> dict:
    if not wb_data or 'детализация' not in wb_data:
        return {}

    df = wb_data['детализация']
    result = {}

    money_col    = 'К перечислению Продавцу за реализованный Товар'
    delivery_col = 'Услуги по доставке товара покупателю'
    storage_col  = 'Хранение'
    fine_col     = 'Общая сумма штрафов'
    logistic_col = 'Возмещение издержек по перемещению и операционной обработке товара'
    obosnov      = 'Обоснование для оплаты'

    mask_sale = df.get(obosnov, pd.Series()) == 'Продажа'
    if money_col in df.columns:
        result['К перечислению за товар (Продажи)'] = df.loc[mask_sale, money_col].sum() if mask_sale.any() else 0
    if delivery_col in df.columns:
        result['Услуги доставки'] = df[delivery_col].sum()
    if storage_col in df.columns:
        result['Хранение'] = df[storage_col].sum()
    if logistic_col in df.columns:
        result['Логистика (перемещение)'] = df[logistic_col].sum()
    if fine_col in df.columns:
        result['Штрафы'] = df[fine_col].sum()
    if money_col in df.columns:
        result['ИТОГО К ПЕРЕЧИСЛЕНИЮ'] = df[money_col].sum()

    return result


def _build_wb_svod(ws, df: pd.DataFrame, period: str, sebes_dict: dict):
    ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:C1')
    c = ws['A1']
    c.value = f'WILDBERRIES — СВОДКА — {period.upper()}'
    c.fill = hfill(COLORS['wb_violet'])
    c.font = hfont(bold=True, white=True, size=12)
    c.alignment = center()
    ws.row_dimensions[1].height = 25

    summary = _extract_wb_summary({'детализация': df})
    row = 3
    for col_idx, h in enumerate(['Статья', 'Сумма, руб.'], 1):
        c = ws.cell(row=row, column=col_idx, value=h)
        c.fill = hfill('BDD7EE')
        c.font = hfont(bold=True)
        c.alignment = center()
        c.border = BORDER_THIN
    row += 1

    for i, (stat, val) in enumerate(summary.items()):
        is_total = 'итого' in stat.lower()
        c1 = ws.cell(row=row, column=1, value=stat)
        c1.fill = hfill(COLORS['total']) if is_total else _row_fill()
        c1.font = hfont(bold=is_total)
        c1.border = BORDER_THIN
        c1.alignment = left()
        c2 = ws.cell(row=row, column=2, value=val)
        c2.fill = hfill(COLORS['total']) if is_total else _row_fill()
        c2.font = hfont(bold=is_total)
        c2.number_format = '#,##0.00'
        c2.border = BORDER_THIN
        c2.alignment = right()
        row += 1

    row += 1
    obosnov = 'Обоснование для оплаты'
    if obosnov in df.columns:
        df_sales = df[df[obosnov] == 'Продажа'].copy()
        if len(df_sales) > 0 and 'Название' in df_sales.columns:
            df_sales['Себес (авто)'] = df_sales['Название'].apply(lambda x: get_sebes(x, sebes_dict))
            grp_cols = ['Название', 'Себес (авто)']
            agg = df_sales.groupby(grp_cols, dropna=False).agg(
                Кол_во=('Кол-во', 'sum'),
                К_перечислению=('К перечислению Продавцу за реализованный Товар', 'sum'),
            ).reset_index()
            agg['Себес итого'] = agg['Себес (авто)'] * agg['Кол_во']

            ws.merge_cells(f'A{row}:F{row}')
            c = ws[f'A{row}']
            c.value = 'ПРОДАЖИ ПО ТОВАРАМ (WB)'
            c.fill = hfill(COLORS['wb_violet'])
            c.font = hfont(bold=True, white=True)
            c.alignment = center()
            row += 1

            prod_headers = ['Название', 'Себес (авто)', 'Кол-во', 'Себес итого', 'К перечислению']
            for col_idx, h in enumerate(prod_headers, 1):
                c = ws.cell(row=row, column=col_idx, value=h)
                c.fill = hfill('BDD7EE')
                c.font = hfont(bold=True)
                c.alignment = center()
                c.border = BORDER_THIN
            row += 1

            for i, (_, ar) in enumerate(agg.iterrows()):
                row_vals = [ar['Название'], ar['Себес (авто)'], ar['Кол_во'],
                            ar['Себес итого'], ar['К_перечислению']]
                for col_idx, v in enumerate(row_vals, 1):
                    v_out = None if (isinstance(v, float) and np.isnan(v)) else v
                    c = ws.cell(row=row, column=col_idx, value=v_out)
                    c.fill = _row_fill()
                    c.border = BORDER_THIN
                    c.alignment = left() if col_idx == 1 else right()
                    if col_idx in (2, 4, 5):
                        c.number_format = '#,##0.00'
                row += 1

    _set_col_widths(ws, {'A': 45, 'B': 14, 'C': 10, 'D': 16, 'E': 16})


def _build_wb_detail(ws, df: pd.DataFrame, sebes_dict: dict):
    ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:J1')
    c = ws['A1']
    c.value = 'WILDBERRIES — ДЕТАЛИЗАЦИЯ (все еженедельные отчёты)'
    c.fill = hfill(COLORS['wb_violet'])
    c.font = hfont(bold=True, white=True, size=12)
    c.alignment = center()
    ws.row_dimensions[1].height = 25

    key_cols = [
        'Название', 'Артикул поставщика', 'Тип документа', 'Обоснование для оплаты',
        'Дата заказа покупателем', 'Дата продажи', 'Кол-во',
        'Вайлдберриз реализовал Товар (Пр)',
        'К перечислению Продавцу за реализованный Товар',
        'Услуги по доставке товара покупателю',
        'Хранение', 'Общая сумма штрафов',
        'Возмещение издержек по перемещению и операционной обработке товара',
    ]
    available = [col for col in key_cols if col in df.columns]
    df_out = df[available].copy()
    if 'Название' in df_out.columns:
        df_out['Себес (авто)'] = df_out['Название'].apply(lambda x: get_sebes(x, sebes_dict))

    row = 3
    for col_idx, h in enumerate(df_out.columns, 1):
        c = ws.cell(row=row, column=col_idx, value=str(h))
        c.fill = hfill(COLORS['wb_violet'])
        c.font = hfont(bold=True, white=True)
        c.alignment = center()
        c.border = BORDER_THIN
    row += 1

    for i, (_, dr) in enumerate(df_out.iterrows()):
        for col_idx, v in enumerate(dr.values, 1):
            v_out = None if (isinstance(v, float) and np.isnan(v)) else v
            c = ws.cell(row=row, column=col_idx, value=v_out)
            c.fill = _row_fill()
            c.border = BORDER_THIN
            c.alignment = left() if col_idx <= 2 else right()
        row += 1

    _set_col_widths(ws, {'A': 40, 'B': 14, 'C': 14, 'D': 30, 'E': 18, 'F': 18,
                         'G': 8, 'H': 14, 'I': 16, 'J': 14, 'K': 12, 'L': 12, 'M': 16})


# ─────────────────────────────────────────────
#  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ─────────────────────────────────────────────
def _build_raw_sheet(ws, df_raw: pd.DataFrame, title: str, color: str):
    ws.sheet_view.showGridLines = False
    max_col = max(df_raw.shape[1], 1)
    end_col = get_column_letter(max_col)
    ws.merge_cells(f'A1:{end_col}1')
    c = ws['A1']
    c.value = title
    c.fill = hfill(color)
    c.font = hfont(bold=True, white=True, size=11)
    c.alignment = center()
    _dump_df_to_ws(ws, df_raw, start_row=3)


def _dump_df_to_ws(ws, df: pd.DataFrame, start_row: int):
    for i, (_, r) in enumerate(df.iterrows()):
        for col_idx, v in enumerate(r.values, 1):
            v_out = None if (isinstance(v, float) and np.isnan(v)) else v
            c = ws.cell(row=start_row + i, column=col_idx, value=v_out)
            c.fill = _row_fill()
            c.border = BORDER_THIN
            c.alignment = left() if col_idx == 1 else right()


# ─────────────────────────────────────────────
#  STREAMLIT UI
# ─────────────────────────────────────────────
def render_sebes_editor():
    """Панель редактирования себестоимости."""
    st.subheader('⚙️ Настройки себестоимости')
    st.caption('Здесь можно изменить себестоимость для любого типа товара. Изменения сохраняются автоматически.')

    sebes = load_sebes()

    col_add, col_reset = st.columns([3, 1])
    with col_reset:
        if st.button('↺ Сбросить к исходным', use_container_width=True):
            save_sebes(dict(DEFAULT_SEBES))
            st.success('Сброшено!')
            st.rerun()

    # Таблица с возможностью редактирования
    df_sebes = pd.DataFrame(
        [{'Тип товара (подстрока в названии)': k, 'Себестоимость, руб.': v}
         for k, v in sebes.items()]
    )

    edited = st.data_editor(
        df_sebes,
        num_rows='dynamic',
        use_container_width=True,
        column_config={
            'Тип товара (подстрока в названии)': st.column_config.TextColumn(
                'Тип товара (подстрока в названии)',
                help='Любая часть названия товара, например: "диффузор 100мл"',
                width='large',
            ),
            'Себестоимость, руб.': st.column_config.NumberColumn(
                'Себестоимость, руб.',
                min_value=0,
                format='%.2f',
                width='medium',
            ),
        },
        hide_index=True,
        key='sebes_editor',
    )

    if st.button('💾 Сохранить себестоимость', type='primary', use_container_width=True):
        new_sebes = {}
        for _, row in edited.iterrows():
            key = str(row['Тип товара (подстрока в названии)']).strip().lower()
            val = row['Себестоимость, руб.']
            if key and not pd.isna(val):
                new_sebes[key] = float(val)
        save_sebes(new_sebes)
        st.success(f'✅ Сохранено {len(new_sebes)} позиций!')
        st.rerun()

    return sebes


def main():
    st.set_page_config(
        page_title='SP CANDLE — Отчёт',
        page_icon='🕯️',
        layout='wide',
    )

    st.title('🕯️ SP CANDLE — Сводный отчёт за месяц')

    # ── Вкладки ──
    tab_main, tab_sebes = st.tabs(['📊 Формирование отчёта', '⚙️ Себестоимость'])

    with tab_sebes:
        sebes_dict = render_sebes_editor()

    with tab_main:
        st.markdown('Загрузите файлы из 1С, Ozon и Wildberries — получите единый Excel-отчёт.')

        # ── Период ──
        col1, col2, col3 = st.columns([1, 1, 2])
        months_ru = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                     'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
        with col1:
            month_num = st.selectbox('Месяц отчёта', range(1, 13),
                                     index=datetime.now().month - 2,
                                     format_func=lambda x: months_ru[x - 1])
        with col2:
            year = st.number_input('Год', value=datetime.now().year, step=1, min_value=2020)
        period = f'{months_ru[month_num - 1]} {year}'

        st.divider()

        # ── WB период фильтрации ──
        with st.expander('📅 Период фильтрации WB-отчётов', expanded=False):
            st.caption(
                'WB-отчёты содержат данные за несколько месяцев. '
                'Укажите период, чтобы оставить только нужные строки.'
            )
            wb_col1, wb_col2 = st.columns(2)
            with wb_col1:
                # Первый день выбранного месяца по умолчанию
                import datetime as dt
                default_from = dt.date(int(year), month_num, 1)
                wb_date_from = st.date_input('Дата С', value=default_from, key='wb_from')
            with wb_col2:
                # Последний день выбранного месяца
                import calendar
                last_day = calendar.monthrange(int(year), month_num)[1]
                default_to = dt.date(int(year), month_num, last_day)
                wb_date_to = st.date_input('Дата ПО', value=default_to, key='wb_to')

        st.divider()

        # ── Загрузка файлов ──
        st.subheader('📂 1С — Прибыльность по товарам (4 XLS-файла)')
        st.caption('Каждый файл — отдельный канал: Интернет, ОПТ, СТМ, Столярный')
        col_a, col_b = st.columns(2)
        with col_a:
            f_internet = st.file_uploader('Интернет магазин', type=['xls'], key='internet')
            f_opt      = st.file_uploader('ОПТ', type=['xls'], key='opt')
        with col_b:
            f_stm      = st.file_uploader('СЗ/СТМ', type=['xls'], key='stm')
            f_stolyar  = st.file_uploader('Магазин Столярный 9', type=['xls'], key='stolyar')

        st.subheader('🔵 Ozon — отчёт начислений (xlsx)')
        f_ozon = st.file_uploader('Отчёт Ozon', type=['xlsx'], key='ozon')

        st.subheader('🟣 Wildberries — еженедельные отчёты (zip-архивы)')
        st.caption('Загрузите все еженедельные отчёты — лишние месяцы будут отфильтрованы по периоду выше')
        f_wb_zips = st.file_uploader('WB отчёты', type=['zip'],
                                      accept_multiple_files=True, key='wb')

        st.divider()

        # ── Кнопка ──
        if st.button('🚀 Сформировать отчёт', type='primary', use_container_width=True):
            # Перечитываем актуальную себестоимость
            sebes_dict = load_sebes()

            with st.spinner('Обрабатываю данные...'):
                ms_data = {}

                channel_files = {
                    'Интернет_сен26':  f_internet,
                    'ОПТ_сен26':       f_opt,
                    'СТМ_сен26':       f_stm,
                    'Столярный_сен26': f_stolyar,
                }
                for key, f in channel_files.items():
                    if f is not None:
                        ch_name = key.split('_')[0]
                        df = parse_1c_file(f, ch_name, sebes_dict)
                        ms_data[key] = df
                        st.success(f'✅ {ch_name}: {len(df)} позиций')

                if not ms_data:
                    st.warning('⚠️ Не загружен ни один файл 1С/МС')

                ozon_data = {}
                if f_ozon is not None:
                    ozon_data = parse_ozon_file(f_ozon)
                    st.success('✅ Ozon загружен')

                wb_data = {}
                if f_wb_zips:
                    wb_data = parse_wb_files(f_wb_zips, wb_date_from, wb_date_to)
                    n = len(wb_data.get('детализация', pd.DataFrame()))
                    st.success(f'✅ WB загружен ({n} строк за выбранный период)')

                if not ms_data and not ozon_data and not wb_data:
                    st.error('Загрузите хотя бы один файл!')
                else:
                    excel_bytes = build_excel(ms_data, ozon_data, wb_data, period, sebes_dict)
                    fname = f'SP_CANDLE_Отчет_{period.replace(" ", "_")}.xlsx'
                    st.download_button(
                        label='⬇️ Скачать Excel-отчёт',
                        data=excel_bytes,
                        file_name=fname,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        type='primary',
                        use_container_width=True,
                    )
                    st.balloons()


if __name__ == '__main__':
    main()
