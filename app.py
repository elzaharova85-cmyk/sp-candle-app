import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
import os, io, zipfile, tempfile, json
from datetime import datetime

# ─────────────────────────────────────────────
#  СЕБЕСТОИМОСТЬ  (по типу товара)
# ─────────────────────────────────────────────
DEFAULT_SEBES = {
    'автодиффузор 7мл':            66,
    'диффузор-спрей 100мл':       233.5,
    'диффузор 100мл':              327,
    'диффузор 50мл':               198,
    'диффузор sign':              1629,
    'диффузор-рефил 5л':           660,
    'диффузор-рефил 500мл':       1288,
    'диффузор-рефил 250мл':        684,
    'диффузор 30мл':               288,
    'срезанный флакон 30мл':       288,
    'тестер шайба':                 31,
    'свеча 320г':                 1334,
    'свеча 220г':                  372,
    'свеча 100г':                  174,
    'свеча pur':                   174,
    'крем для рук дой-пак 1 л':    793,
    'крем для рук 1л':             793,
    'крем для рук 45мл':           191,
    'крем для рук 250мл':          290,
    'мыло для рук 250мл':          201,
    'мыло для рук 500мл':          300,
    'мыло для рук 5 л':           1850,
    'мыло для рук 5л':            1850,
    'набор hand care':             875,
    'набор свечей candle trio':    498,
    'набор: свеча220':            1043,
    'подарочная коробка бордовая': 245,
    'фибровые палочки 3/20':        30,
    'фибровые палочки 4/22':        34,
    'фибровые палочки 5/20':        30,
    'спички каминные':              32.5,
    'стакан agger 350мл':         1260,
    'стакан agger 230мл':         1050,
    'подвес':                      280,
}

SEBES_FILE = os.path.join(os.path.dirname(__file__), 'sebes.json')

def load_sebes() -> dict:
    if os.path.exists(SEBES_FILE):
        try:
            with open(SEBES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return dict(DEFAULT_SEBES)

def save_sebes(d: dict):
    with open(SEBES_FILE, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

def get_sebes(name: str, sebes_dict: dict = None) -> float | None:
    if not isinstance(name, str):
        return None
    d = sebes_dict if sebes_dict is not None else load_sebes()
    n = name.lower()
    for key, val in d.items():
        if key in n:
            return val
    return None

def _to_float(v) -> float:
    try:
        f = float(v)
        return 0.0 if (f != f) else f  # nan → 0
    except (TypeError, ValueError):
        return 0.0


# ─────────────────────────────────────────────
#  КАТЕГОРИЯ ТОВАРА (для группировки)
# ─────────────────────────────────────────────

# Подстроки, которые относят товар к доставке
DELIVERY_KEYWORDS = [
    'сдэк', 'доставка', 'самовывоз', 'курьер', 'пункт выдачи', 'зарубежье',
    'ближнее зарубежье',
]
# Подстроки для основных товаров (типы продукции)
MAIN_KEYWORDS = [
    'диффузор', 'свеча', 'крем для рук', 'мыло для рук', 'набор', 'шампунь',
    'гель для душа', 'автодиффузор', 'рефил', 'срезанный флакон', 'тестер',
]

def classify_item(name: str) -> str:
    """Возвращает 'основные', 'доставка' или 'сопутствующие'."""
    n = name.lower().strip()
    for kw in DELIVERY_KEYWORDS:
        if kw in n:
            return 'доставка'
    for kw in MAIN_KEYWORDS:
        if kw in n:
            return 'основные'
    return 'сопутствующие'


# ─────────────────────────────────────────────
#  ПАРСИНГ ФАЙЛОВ
# ─────────────────────────────────────────────

def parse_1c_file(file_obj, channel_name: str, sebes_dict: dict) -> pd.DataFrame:
    import xlrd
    with tempfile.NamedTemporaryFile(suffix='.xls', delete=False) as tmp:
        tmp.write(file_obj.read())
        tmp_path = tmp.name

    df_raw = pd.read_excel(tmp_path, engine='xlrd', header=None)
    os.unlink(tmp_path)

    # Ищем строку с "Наименование"
    header_row = None
    for i, row in df_raw.iterrows():
        vals = [str(v).strip() for v in row.values if pd.notna(v)]
        if any('наименование' in v.lower() for v in vals):
            header_row = i
            break

    data_start = (header_row + 2) if header_row is not None else 13

    rows = []
    for i in range(data_start, len(df_raw)):
        row = df_raw.iloc[i]
        name      = row.iloc[3]  if len(row) > 3  else None
        qty       = row.iloc[5]  if len(row) > 5  else None
        price     = row.iloc[6]  if len(row) > 6  else None
        sebes_1c  = row.iloc[7]  if len(row) > 7  else None
        sum_sales = row.iloc[9]  if len(row) > 9  else None
        sum_sebes = row.iloc[10] if len(row) > 10 else None
        profit    = row.iloc[20] if len(row) > 20 else None

        if not isinstance(name, str) or len(name.strip()) < 2:
            continue
        if name.strip().lower() in ('наименование', 'итого', ''):
            continue
        try:
            qty_f = float(qty)
        except (TypeError, ValueError):
            continue
        if qty_f == 0:
            continue

        rows.append({
            'Канал':        channel_name,
            'Наименование': name.strip(),
            'Кол-во':       qty_f,
            'Цена':         _to_float(price),
            'Себес_1С':     _to_float(sebes_1c),
            'Себес_авто':   get_sebes(name, sebes_dict),
            'Сумма продаж': _to_float(sum_sales),
            'Сумма себест': _to_float(sum_sebes),
            'Прибыль':      _to_float(profit),
        })

    return pd.DataFrame(rows)


def parse_ozon_file(file_obj) -> dict:
    raw = file_obj.read()
    xl = pd.ExcelFile(io.BytesIO(raw))
    result = {'_sheet_names': xl.sheet_names}

    for sh in xl.sheet_names:
        sh_low = sh.strip().lower()
        # Лист с начислениями/сводкой
        if any(x in sh_low for x in ('начислени', 'сводк', 'summary', 'финанс')):
            result['начисления'] = pd.read_excel(io.BytesIO(raw), sheet_name=sh, header=None)
        # Лист с детализацией по товарам
        elif any(x in sh_low for x in ('детализ', 'detail', 'товар')):
            result['детализация'] = pd.read_excel(io.BytesIO(raw), sheet_name=sh)
        # Лист по типам
        elif any(x in sh_low for x in ('типам', 'тип')):
            result['по_типам'] = pd.read_excel(io.BytesIO(raw), sheet_name=sh, header=None)

    # Fallback по позиции — берём все листы
    if 'начисления' not in result and len(xl.sheet_names) >= 1:
        result['начисления'] = pd.read_excel(io.BytesIO(raw), sheet_name=xl.sheet_names[0], header=None)
    if 'детализация' not in result and len(xl.sheet_names) >= 2:
        result['детализация'] = pd.read_excel(io.BytesIO(raw), sheet_name=xl.sheet_names[1])
    if 'по_типам' not in result and len(xl.sheet_names) >= 3:
        result['по_типам'] = pd.read_excel(io.BytesIO(raw), sheet_name=xl.sheet_names[2], header=None)

    # Обратная совместимость — дублируем под старым ключом
    if 'начисления' in result:
        result['сводка'] = result['начисления']

    return result


def parse_wb_files(zip_files, date_from=None, date_to=None) -> dict:
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


# ─────────────────────────────────────────────
#  СТИЛИ EXCEL
# ─────────────────────────────────────────────
COLORS = {
    'header_dark':   '1F3864',
    'header_blue':   '2E75B6',
    'group_header':  'D9E1F2',  # секция Основные/Сопутствующие/Доставка
    'total_section': 'BDD7EE',  # ИТОГО по секции
    'grand_total':   '1F3864',  # ИТОГО всего
    'row_white':     'FFFFFF',
    'ozon_blue':     '005BFF',
    'wb_violet':     '7B2FBE',
    'yellow':        'FFEB9C',
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

def _write_cell(ws, row, col, value, fill_color, bold=False, white=False,
                fmt=None, align=None, size=10):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = hfill(fill_color)
    c.font = hfont(bold=bold, white=white, size=size)
    c.border = BORDER_THIN
    c.alignment = align or (left() if col == 1 else right())
    if fmt:
        c.number_format = fmt
    return c


# ─────────────────────────────────────────────
#  ГРУППИРОВКА ДАННЫХ КАНАЛА
# ─────────────────────────────────────────────

def get_product_type(name: str, sebes_dict: dict) -> str | None:
    """
    Определяет тип товара по ключам sebes_dict.
    Возвращает ключ с наибольшим совпадением (самый длинный подходящий ключ).
    """
    n = name.lower().strip()
    best = None
    best_len = 0
    for key in sebes_dict:
        if key in n and len(key) > best_len:
            best = key
            best_len = len(key)
    return best


def group_channel_df(df: pd.DataFrame, sebes_dict: dict) -> dict:
    """
    Разбивает DataFrame канала на группы: основные, сопутствующие, доставка.
    ОСНОВНЫЕ товары агрегируются по типу (ключи sebes_dict).
    Возвращает dict:
      'основные'      → список агрегированных строк по типу товара
      'сопутствующие' → список строк по каждому товару (без группировки)
      'доставка'      → список строк по доставке
    """
    # Временные накопители для основных (по типу)
    main_by_type: dict[str, dict] = {}   # key=тип, value=накопленные суммы
    sop_rows = []
    del_rows = []

    for _, r in df.iterrows():
        name = str(r.get('Наименование', '')).strip()
        if not name or name == 'nan':
            continue

        qty_f  = _to_float(r.get('Кол-во'))
        rev    = _to_float(r.get('Сумма продаж'))
        profit = _to_float(r.get('Прибыль'))

        cat = classify_item(name)

        if cat == 'доставка':
            del_rows.append({
                'name':      name,
                'qty':       qty_f,
                'rev':       rev,
                'sebes1':    0,
                'sebes_all': 0,
                'profit':    rev,   # доставка — прибыль = выручка
                'margin':    1.0,
            })
            continue

        # Определяем тип товара через sebes_dict
        product_type = get_product_type(name, sebes_dict)
        sebes1 = sebes_dict.get(product_type, 0) if product_type else 0
        sebes_all = sebes1 * qty_f

        if cat == 'основные' and product_type:
            # Агрегируем по типу
            # Нормализуем название типа: первая буква заглавная
            type_display = product_type.capitalize()
            if type_display not in main_by_type:
                main_by_type[type_display] = {
                    'qty': 0, 'rev': 0, 'sebes_all': 0, 'profit': 0,
                    'sebes1': sebes1,
                }
            main_by_type[type_display]['qty']      += qty_f
            main_by_type[type_display]['rev']      += rev
            main_by_type[type_display]['sebes_all'] += sebes_all
            main_by_type[type_display]['profit']   += profit
        else:
            # Сопутствующие или основные без известного типа → не агрегируем
            sop_rows.append({
                'name':      name,
                'qty':       qty_f,
                'rev':       rev,
                'sebes1':    sebes1,
                'sebes_all': sebes_all,
                'profit':    profit,
                'margin':    (profit / rev) if rev else 0,
            })

    # Преобразуем накопленные основные в список строк
    main_rows = []
    for type_name, t in main_by_type.items():
        margin = (t['profit'] / t['rev']) if t['rev'] else 0
        main_rows.append({
            'name':      type_name,
            'qty':       t['qty'],
            'rev':       t['rev'],
            'sebes1':    t['sebes1'],
            'sebes_all': t['sebes_all'],
            'profit':    t['profit'],
            'margin':    margin,
        })
    # Сортируем по выручке убыванием
    main_rows.sort(key=lambda x: x['rev'], reverse=True)

    return {
        'основные':      main_rows,
        'сопутствующие': sop_rows,
        'доставка':      del_rows,
    }


def calc_group_totals(rows: list) -> dict:
    """Считает ИТОГО по группе строк."""
    qty   = sum(r['qty'] or 0 for r in rows)
    rev   = sum(r['rev'] or 0 for r in rows)
    sebes = sum(r['sebes_all'] or 0 for r in rows)
    prof  = sum(r['profit'] or 0 for r in rows)
    marg  = (prof / rev) if rev else 0
    return {'qty': qty, 'rev': rev, 'sebes_all': sebes, 'profit': prof, 'margin': marg}


# ─────────────────────────────────────────────
#  ПОСТРОЕНИЕ EXCEL-ОТЧЁТА
# ─────────────────────────────────────────────

def build_excel(ms_data: dict, ozon_data: dict, wb_data: dict,
                period: str, sebes_dict: dict) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    ws_svod = wb.create_sheet('СВОДНЫЙ')

    # Сначала строим листы каналов и собираем итоги для сводного
    channel_display = [
        ('Интернет_сен26',  'Интернет',  'SP.Candle интернет магазин'),
        ('ОПТ_сен26',       'ОПТ',       'ОПТ'),
        ('СТМ_сен26',       'СТМ',       'СЗ/СТМ'),
        ('Столярный_сен26', 'Столярный', 'Магазин Столярный 9'),
    ]

    svod_rows = []  # данные для сводного листа
    for key, sheet_name, display_name in channel_display:
        if key not in ms_data:
            continue
        df = ms_data[key]
        groups = group_channel_df(df, sebes_dict)

        ws_ch = wb.create_sheet(f'МС_{sheet_name}')
        _build_channel_sheet(ws_ch, groups, display_name, period)

        # Итоги для сводного (основные + сопутствующие, без доставки)
        main_rows  = groups['основные'] + groups['сопутствующие']
        deliv_rows = groups['доставка']
        t_main  = calc_group_totals(main_rows)
        t_deliv = calc_group_totals(deliv_rows)
        svod_rows.append({
            'name':        display_name,
            'main':        t_main,
            'delivery':    t_deliv,
        })

    ozon_real = {k: v for k, v in ozon_data.items() if not k.startswith('_')}

    _build_svod_sheet(ws_svod, svod_rows, ozon_real, wb_data, period)

    if ozon_real:
        ws_ozon_svod = wb.create_sheet('Ozon_Сводка')
        _build_ozon_svod(ws_ozon_svod, ozon_real, period)

        ws_ozon_types = wb.create_sheet('Ozon_По типам')
        _build_ozon_by_types(ws_ozon_types, ozon_real, period, sebes_dict)

        if 'детализация' in ozon_real:
            ws_ozon_det = wb.create_sheet('Ozon_Детализация')
            _build_ozon_detail(ws_ozon_det, ozon_real['детализация'], sebes_dict)

    if wb_data and 'детализация' in wb_data:
        ws_wb_svod = wb.create_sheet('WB_Сводка')
        _build_wb_svod(ws_wb_svod, wb_data['детализация'], period, sebes_dict)

        ws_wb_types = wb.create_sheet('WB_По типам')
        _build_wb_by_types(ws_wb_types, wb_data['детализация'], period, sebes_dict)

        ws_wb_det = wb.create_sheet('WB_Детализация')
        _build_wb_detail(ws_wb_det, wb_data['детализация'], sebes_dict)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
#  ЛИСТ: КАНАЛ (Интернет / ОПТ / СТМ / Столярный)
# ─────────────────────────────────────────────

def _build_channel_sheet(ws, groups: dict, channel_name: str, period: str):
    ws.sheet_view.showGridLines = False
    COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
    NCOLS = 7

    # Заголовок
    ws.merge_cells('A1:G1')
    c = ws['A1']
    c.value = f'Сводный отчёт — {channel_name} — {period}'
    c.fill = hfill(COLORS['header_blue'])
    c.font = hfont(bold=True, white=True, size=12)
    c.alignment = center()
    ws.row_dimensions[1].height = 25

    # Шапка таблицы
    hdrs = ['Наименование', 'Кол-во', 'Выручка', 'Себес 1 ед', 'Себес итого', 'Прибыль', 'Маржа %']
    row = 2
    for ci, h in enumerate(hdrs, 1):
        _write_cell(ws, row, ci, h, 'BDD7EE', bold=True, align=center())
    row += 1

    section_labels = [
        ('основные',       'Основные товары'),
        ('сопутствующие',  'Сопутствующие товары'),
        ('доставка',       'Доставка и транспорт'),
    ]

    grand_qty = grand_rev = grand_sebes = grand_prof = 0.0

    for cat_key, cat_label in section_labels:
        items = groups.get(cat_key, [])

        # Заголовок секции (синяя строка)
        ws.merge_cells(f'A{row}:G{row}')
        c = ws.cell(row=row, column=1, value=cat_label)
        c.fill = hfill(COLORS['header_blue'])
        c.font = hfont(bold=True, white=True)
        c.alignment = left()
        c.border = BORDER_THIN
        ws.row_dimensions[row].height = 16
        row += 1

        if not items:
            row += 1
            continue

        for r in items:
            sebes1_out = r['sebes1'] if r['sebes1'] else None
            vals = [r['name'], r['qty'], r['rev'], sebes1_out,
                    r['sebes_all'] or None, r['profit'], r['margin']]
            fmts = [None, '#,##0', '#,##0.00', '#,##0.00', '#,##0.00', '#,##0.00', '0.0%']
            for ci, (v, fmt) in enumerate(zip(vals, fmts), 1):
                _write_cell(ws, row, ci, v, COLORS['row_white'], fmt=fmt)
            row += 1

        # ИТОГО секции
        t = calc_group_totals(items)
        label = f'ИТОГО {cat_label.lower()}'
        vals_t = [label, t['qty'], t['rev'], None, t['sebes_all'], t['profit'], t['margin']]
        fmts_t = [None, '#,##0', '#,##0.00', None, '#,##0.00', '#,##0.00', '0.0%']
        for ci, (v, fmt) in enumerate(zip(vals_t, fmts_t), 1):
            _write_cell(ws, row, ci, v, COLORS['total_section'], bold=True, fmt=fmt)
        row += 1

        if cat_key != 'доставка':
            grand_qty   += t['qty']
            grand_rev   += t['rev']
            grand_sebes += t['sebes_all']
            grand_prof  += t['profit']

    # Итоговая строка ИТОГО (тёмно-синяя)
    grand_margin = (grand_prof / grand_rev) if grand_rev else 0
    vals_g = ['ИТОГО', grand_qty, grand_rev, None, grand_sebes, grand_prof, grand_margin]
    fmts_g = [None, '#,##0', '#,##0.00', None, '#,##0.00', '#,##0.00', '0.0%']
    for ci, (v, fmt) in enumerate(zip(vals_g, fmts_g), 1):
        _write_cell(ws, row, ci, v, COLORS['grand_total'],
                    bold=True, white=True, fmt=fmt)

    _set_col_widths(ws, {'A': 38, 'B': 10, 'C': 16, 'D': 16, 'E': 16, 'F': 14, 'G': 10})


# ─────────────────────────────────────────────
#  ЛИСТ: СВОДНЫЙ
# ─────────────────────────────────────────────

def _build_svod_sheet(ws, svod_rows: list, ozon_data: dict, wb_data: dict, period: str):
    ws.sheet_view.showGridLines = False

    # Заголовок
    ws.merge_cells('A1:H1')
    c = ws['A1']
    c.value = f'СВОДНЫЙ ОТЧЁТ SP CANDLE — {period.upper()}'
    c.fill = hfill(COLORS['header_dark'])
    c.font = hfont(bold=True, white=True, size=13)
    c.alignment = center()
    ws.row_dimensions[1].height = 30

    row = 3

    # ── Блок 1С ──
    ws.merge_cells(f'A{row}:H{row}')
    c = ws[f'A{row}']
    c.value = '1С — ПРОДАЖИ ПО КАНАЛАМ'
    c.fill = hfill(COLORS['header_blue'])
    c.font = hfont(bold=True, white=True)
    c.alignment = center()
    row += 1

    hdrs = ['Канал', 'Кол-во', 'Выручка', 'Себес итого', 'Прибыль', 'Маржа %']
    for ci, h in enumerate(hdrs, 1):
        _write_cell(ws, row, ci, h, 'BDD7EE', bold=True, align=center())
    row += 1

    # Секция основных + сопутствующих товаров
    ws.merge_cells(f'A{row}:F{row}')
    c = ws.cell(row=row, column=1, value='Основные и сопутствующие товары')
    c.fill = hfill(COLORS['group_header'])
    c.font = hfont(bold=True)
    c.alignment = left()
    c.border = BORDER_THIN
    row += 1

    tot_qty = tot_rev = tot_sebes = tot_prof = 0.0
    for sr in svod_rows:
        t = sr['main']
        marg = (t['profit'] / t['rev']) if t['rev'] else 0
        vals = [sr['name'], t['qty'], t['rev'], t['sebes_all'], t['profit'], marg]
        fmts = [None, '#,##0', '#,##0.00', '#,##0.00', '#,##0.00', '0.0%']
        for ci, (v, fmt) in enumerate(zip(vals, fmts), 1):
            _write_cell(ws, row, ci, v, COLORS['row_white'], fmt=fmt)
        tot_qty   += t['qty']
        tot_rev   += t['rev']
        tot_sebes += t['sebes_all']
        tot_prof  += t['profit']
        row += 1

    t_marg = (tot_prof / tot_rev) if tot_rev else 0
    vals_t = ['ИТОГО', tot_qty, tot_rev, tot_sebes, tot_prof, t_marg]
    fmts_t = [None, '#,##0', '#,##0.00', '#,##0.00', '#,##0.00', '0.0%']
    for ci, (v, fmt) in enumerate(zip(vals_t, fmts_t), 1):
        _write_cell(ws, row, ci, v, COLORS['total_section'], bold=True, fmt=fmt)
    row += 1

    # Секция доставки
    ws.merge_cells(f'A{row}:F{row}')
    c = ws.cell(row=row, column=1, value='Доставка и транспорт')
    c.fill = hfill(COLORS['group_header'])
    c.font = hfont(bold=True)
    c.alignment = left()
    c.border = BORDER_THIN
    row += 1

    del_qty = del_rev = 0.0
    for sr in svod_rows:
        t = sr['delivery']
        vals = [sr['name'], t['qty'], t['rev'], 0, t['rev'], '1']
        fmts = [None, '#,##0', '#,##0.00', None, '#,##0.00', '0%']
        for ci, (v, fmt) in enumerate(zip([sr['name'], t['qty'], t['rev']], [None, '#,##0', '#,##0.00']), 1):
            _write_cell(ws, row, ci, v, COLORS['row_white'], fmt=fmt)
        del_qty += t['qty']
        del_rev += t['rev']
        row += 1

    vals_d = ['ИТОГО', del_qty, del_rev]
    for ci, (v, fmt) in enumerate(zip(vals_d, [None, '#,##0', '#,##0.00']), 1):
        _write_cell(ws, row, ci, v, COLORS['total_section'], bold=True, fmt=fmt)
    row += 2

    # ── Блок Ozon ──
    ws.merge_cells(f'A{row}:H{row}')
    c = ws[f'A{row}']
    c.value = 'OZON — НАЧИСЛЕНИЯ'
    c.fill = hfill(COLORS['ozon_blue'])
    c.font = hfont(bold=True, white=True)
    c.alignment = center()
    row += 1

    for ci, h in enumerate(['Статья', 'Сумма, руб.'], 1):
        _write_cell(ws, row, ci, h, 'BDD7EE', bold=True, align=center())
    row += 1

    # Берём итоговую сумму Ozon к перечислению
    df_raw_ozon = ozon_data.get('начисления')
    if df_raw_ozon is None:
        df_raw_ozon = ozon_data.get('сводка')
    ozon_itog = None
    if df_raw_ozon is not None and not df_raw_ozon.empty:
        sections = _parse_ozon_nacisleniya(df_raw_ozon)
        prod_total  = sum(v for k, v in sections['продажи'].items()   if 'итого' not in k.lower())
        nach_total  = sum(v for k, v in sections['начисления'].items() if 'итого' not in k.lower())
        # Явное итого или вычисляем
        for k, v in sections['итого'].items():
            ozon_itog = v
        if ozon_itog is None:
            ozon_itog = prod_total + nach_total
        rows_ozon = [
            ('Продажи и возвраты (валовые)', prod_total),
            ('Начисления / удержания Ozon',  nach_total),
            ('ИТОГО К ПЕРЕЧИСЛЕНИЮ ПРОДАВЦУ', ozon_itog),
        ]
    else:
        rows_ozon = []

    for stat, val in rows_ozon:
        is_total = 'итого' in stat.lower()
        fill = COLORS['total_section'] if is_total else COLORS['row_white']
        _write_cell(ws, row, 1, stat, fill, bold=is_total)
        _write_cell(ws, row, 2, val,  fill, bold=is_total, fmt='#,##0.00')
        row += 1
    if not rows_ozon:
        _write_cell(ws, row, 1, '⚠️ Файл Ozon не загружен', COLORS['row_white'])
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

    for ci, h in enumerate(['Статья', 'Сумма, руб.'], 1):
        _write_cell(ws, row, ci, h, 'BDD7EE', bold=True, align=center())
    row += 1

    wb_detail_df = wb_data.get('детализация') if wb_data else None
    if wb_detail_df is not None and not wb_detail_df.empty:
        m = _extract_wb_metrics(wb_detail_df)
        wb_svod_rows = [
            ('Продажи (выручка WB)', m.get('retail_total', 0)),
            ('К перечислению (нетто)', m.get('money_net', 0)),
            ('Комиссия WB', m.get('commission', 0)),
            ('Доставка', m.get('cost_delivery', 0)),
            ('Хранение', m.get('cost_storage', 0)),
            ('Приёмка', m.get('cost_accept', 0)),
            ('Штрафы', m.get('cost_fines', 0)),
            ('Прочие удержания', m.get('cost_other', 0)),
            ('Итого расходы WB', m.get('total_costs', 0)),
            ('ИТОГО К ПОЛУЧЕНИЮ', m.get('total_payment', 0)),
        ]
        for stat, val in wb_svod_rows:
            is_total = 'итого' in stat.lower()
            fill = COLORS['total_section'] if is_total else COLORS['row_white']
            _write_cell(ws, row, 1, stat, fill, bold=is_total)
            _write_cell(ws, row, 2, val, fill, bold=is_total, fmt='#,##0.00')
            row += 1
    else:
        _write_cell(ws, row, 1, '⚠️ Файл WB не загружен', COLORS['row_white'])
        row += 1

    _set_col_widths(ws, {'A': 38, 'B': 14, 'C': 16, 'D': 16, 'E': 14, 'F': 10, 'G': 12, 'H': 12})


# ─────────────────────────────────────────────
#  ЛИСТ: OZON
# ─────────────────────────────────────────────

def _parse_ozon_nacisleniya(df_raw: pd.DataFrame) -> dict:
    """
    Парсит лист начислений Ozon (header=None).
    Ищет строки вида: [label, ..., value] и группирует их по секциям.
    Возвращает dict вида:
      {
        'продажи': {'Выручка': 1221148, 'Программы партнёров': 9310, ...},
        'начисления': {'Вознаграждение Ozon': -1194161, ...},
        'итого': {'ИТОГО К ПЕРЕЧИСЛЕНИЮ ПРОДАВЦУ': 902057},
      }
    """
    sections = {'продажи': {}, 'начисления': {}, 'итого': {}}
    current_section = None

    SECTION_MARKERS = {
        'продажи': ['продаж', 'возврат'],
        'начисления': ['начислени', 'удержани', 'вознаграждени', 'услуги'],
        'итого': ['итого к перечислени', 'итого к оплате'],
    }

    for _, row in df_raw.iterrows():
        vals = [v for v in row.values if pd.notna(v) and str(v).strip() not in ('', 'nan')]
        if not vals:
            continue

        label = str(vals[0]).strip()
        label_low = label.lower()

        # Определяем секцию по метке
        for sec, markers in SECTION_MARKERS.items():
            if any(m in label_low for m in markers):
                current_section = sec
                break

        # Ищем числовое значение в строке (последнее числовое значение в строке)
        num_val = None
        for v in reversed(vals):
            try:
                num_val = float(v)
                break
            except (TypeError, ValueError):
                continue

        if num_val is not None and current_section and label_low not in ('nan',):
            # Пропускаем строки-заголовки секций (без числа они уже попали выше)
            is_section_header = any(
                label_low.startswith(m) for m in
                ['продажи и возвраты', 'начисления', 'итог']
            )
            if not is_section_header:
                sections[current_section][label] = num_val

    # Считаем итоговые суммы если они не нашлись явно
    if sections['продажи'] and 'ИТОГО Продажи и возвраты' not in sections['продажи']:
        s = sum(sections['продажи'].values())
        sections['продажи']['ИТОГО Продажи и возвраты'] = s

    if sections['начисления'] and 'ИТОГО Начисления' not in sections['начисления']:
        s = sum(sections['начисления'].values())
        sections['начисления']['ИТОГО Начисления'] = s

    return sections


def _build_ozon_svod(ws, ozon_data: dict, period: str):
    """Лист Ozon_Сводка — структура как в эталонном отчёте."""
    ws.sheet_view.showGridLines = False

    BLUE   = COLORS['ozon_blue']
    GREEN  = 'E2EFDA'
    YELLOW = 'FFF2CC'
    BEIGE  = 'F2F2F2'
    WHITE  = COLORS['row_white']
    TOTAL  = COLORS['total_section']

    def title_row(r, text, color):
        ws.merge_cells(f'A{r}:B{r}')
        c = ws.cell(row=r, column=1, value=text)
        c.fill = hfill(color)
        c.font = hfont(bold=True, size=11)
        c.alignment = left()
        c.border = BORDER_THIN
        ws.row_dimensions[r].height = 18

    def data_row(r, label, value, bold=False, fill=WHITE, fmt='#,##0'):
        _write_cell(ws, r, 1, label, fill, bold=bold, align=left())
        _write_cell(ws, r, 2, value, fill, bold=bold, fmt=fmt)

    row = 1
    # Заголовок
    ws.merge_cells(f'A{row}:B{row}')
    c = ws.cell(row=row, column=1, value=f'Отчёт Ozon — {period}')
    c.fill = hfill(BLUE)
    c.font = hfont(bold=True, white=True, size=13)
    c.alignment = left()
    c.border = BORDER_THIN
    ws.row_dimensions[row].height = 24
    row += 2

    # Получаем данные из листа начислений
    df_raw = ozon_data.get('начисления')
    if df_raw is None:
        df_raw = ozon_data.get('сводка')

    if df_raw is not None and not df_raw.empty:
        sections = _parse_ozon_nacisleniya(df_raw)

        # ── ПРОДАЖИ И ВОЗВРАТЫ ──
        title_row(row, 'ПРОДАЖИ И ВОЗВРАТЫ', GREEN)
        row += 1
        prod_total = 0.0
        for label, val in sections['продажи'].items():
            is_total = 'итого' in label.lower()
            fill_r = TOTAL if is_total else WHITE
            data_row(row, label, val, bold=is_total, fill=fill_r)
            if not is_total:
                prod_total += val
            row += 1
        # Если ИТОГО не нашлось — пишем сами
        if not any('итого' in l.lower() for l in sections['продажи']):
            data_row(row, 'ИТОГО Продажи и возвраты', prod_total, bold=True, fill=TOTAL)
            row += 1
        row += 1

        # ── НАЧИСЛЕНИЯ (УДЕРЖАНИЯ) ──
        title_row(row, 'НАЧИСЛЕНИЯ (УДЕРЖАНИЯ OZON)', YELLOW)
        row += 1
        nach_total = 0.0
        for label, val in sections['начисления'].items():
            is_total = 'итого' in label.lower()
            fill_r = TOTAL if is_total else WHITE
            data_row(row, label, val, bold=is_total, fill=fill_r)
            if not is_total:
                nach_total += val
            row += 1
        if not any('итого' in l.lower() for l in sections['начисления']):
            data_row(row, 'ИТОГО Начисления', nach_total, bold=True, fill=TOTAL)
            row += 1
        row += 1

        # ── ИТОГ ──
        title_row(row, 'ИТОГ', BEIGE)
        row += 1

        # Ищем итоговую сумму к перечислению
        itog_val = None
        for label, val in sections['итого'].items():
            itog_val = val
        if itog_val is None:
            # Считаем: продажи + начисления
            itog_val = prod_total + nach_total

        data_row(row, 'ИТОГО К ПЕРЕЧИСЛЕНИЮ ПРОДАВЦУ', itog_val, bold=True, fill=TOTAL)
        row += 1

        # % от продаж
        if prod_total:
            pct = itog_val / prod_total
            _write_cell(ws, row, 1, 'Итого к оплате, % от суммы продаж', BEIGE, align=left())
            _write_cell(ws, row, 2, pct, BEIGE, fmt='0.0%')
        row += 1
    else:
        ws.merge_cells(f'A{row}:B{row}')
        c = ws.cell(row=row, column=1, value='⚠️ Данные Ozon не найдены — проверьте структуру файла')
        c.fill = hfill('FCE4D6')
        c.font = hfont(bold=True)

    _set_col_widths(ws, {'A': 45, 'B': 18})


def _build_ozon_by_types(ws, ozon_data: dict, period: str, sebes_dict: dict):
    """
    Лист Ozon_По типам — продажи по типам товара с себестоимостью.
    Если есть лист 'детализация' — группируем по типу из него.
    Иначе — показываем типы из sebes_dict с нулевым кол-вом.
    """
    ws.sheet_view.showGridLines = False

    BLUE  = COLORS['ozon_blue']
    LIGHT = 'D9E8FB'
    WHITE = COLORS['row_white']
    TOTAL = COLORS['total_section']

    # Заголовок
    ws.merge_cells('A1:D1')
    c = ws['A1']
    c.value = f'Продажи по типам товара - Ozon, {period}'
    c.fill = hfill(BLUE)
    c.font = hfont(bold=True, white=True, size=12)
    c.alignment = left()
    c.border = BORDER_THIN
    ws.row_dimensions[1].height = 22
    row = 3

    # Шапка
    for ci, h in enumerate(['Тип товара', 'Кол-во, шт', 'Себестоимость за 1 шт', 'Себес итого'], 1):
        _write_cell(ws, row, ci, h, LIGHT, bold=True, align=center())
    row += 1

    # Группировка по типам из детализации
    type_totals: dict[str, float] = {}

    if 'детализация' in ozon_data:
        df = ozon_data['детализация']
        name_col = None
        qty_col  = None
        for col in df.columns:
            cl = col.lower()
            if 'название' in cl or 'наименование' in cl or 'товар' in cl:
                name_col = col
            if 'кол' in cl or 'количество' in cl:
                qty_col = col

        if name_col:
            for _, r in df.iterrows():
                name = str(r.get(name_col, '')).strip()
                qty  = _to_float(r.get(qty_col, 1) if qty_col else 1)
                ptype = get_product_type(name, sebes_dict)
                if ptype:
                    type_totals[ptype] = type_totals.get(ptype, 0) + qty

    # Если детализации нет — берём все типы из sebes_dict с нулём
    if not type_totals:
        for key in sebes_dict:
            type_totals[key] = 0.0

    grand_qty   = 0.0
    grand_sebes = 0.0
    for ptype, qty in sorted(type_totals.items(), key=lambda x: -x[1]):
        sebes1    = sebes_dict.get(ptype, 0)
        sebes_all = sebes1 * qty
        _write_cell(ws, row, 1, ptype.capitalize(), WHITE, align=left())
        _write_cell(ws, row, 2, qty,       WHITE, fmt='#,##0')
        _write_cell(ws, row, 3, sebes1,    WHITE, fmt='#,##0.00')
        _write_cell(ws, row, 4, sebes_all, WHITE, fmt='#,##0.00')
        grand_qty   += qty
        grand_sebes += sebes_all
        row += 1

    # ИТОГО
    _write_cell(ws, row, 1, 'ИТОГО', TOTAL, bold=True, align=left())
    _write_cell(ws, row, 2, grand_qty,   TOTAL, bold=True, fmt='#,##0')
    _write_cell(ws, row, 3, None,        TOTAL)
    _write_cell(ws, row, 4, grand_sebes, TOTAL, bold=True, fmt='#,##0.00')

    _set_col_widths(ws, {'A': 28, 'B': 14, 'C': 24, 'D': 16})


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
    for ci, h in enumerate(df.columns, 1):
        _write_cell(ws, row, ci, str(h), COLORS['ozon_blue'], bold=True, white=True, align=center())
    row += 1

    for _, dr in df.iterrows():
        for ci, v in enumerate(dr.values, 1):
            v_out = None if pd.isna(v) else v
            _write_cell(ws, row, ci, v_out, COLORS['row_white'],
                        align=left() if ci <= 3 else right())
        row += 1

    _set_col_widths(ws, {'A': 18, 'B': 20, 'C': 22, 'D': 25, 'E': 14, 'F': 14,
                         'G': 40, 'H': 10, 'I': 12, 'J': 16})


# ─────────────────────────────────────────────
#  ЛИСТ: WB
# ─────────────────────────────────────────────

def _get_wb_col(df: pd.DataFrame, *candidates) -> str | None:
    """Ищет первый подходящий столбец из кандидатов."""
    for c in candidates:
        for col in df.columns:
            if c.lower() in col.lower():
                return col
    return None


def _wb_col_sum(df: pd.DataFrame, *candidates) -> float:
    col = _get_wb_col(df, *candidates)
    if col:
        return float(df[col].fillna(0).sum())
    return 0.0


def _extract_wb_metrics(df: pd.DataFrame) -> dict:
    """
    Считает все метрики WB из сырого DataFrame детализации.
    Возвращает dict с именованными числами.
    """
    obosnov_col  = _get_wb_col(df, 'Обоснование для оплаты', 'обоснование')
    qty_col      = _get_wb_col(df, 'Кол-во', 'количество')
    retail_col   = _get_wb_col(df, 'Розничная цена с учётом скидки', 'розничная цена')
    wb_real_col  = _get_wb_col(df, 'Вайлдберриз реализовал Товар', 'вб реализовал')
    money_col    = _get_wb_col(df, 'К перечислению Продавцу за реализованный', 'к перечислению продавцу')
    delivery_col = _get_wb_col(df, 'Услуги по доставке товара покупателю', 'доставк')
    storage_col  = _get_wb_col(df, 'Хранение')
    accept_col   = _get_wb_col(df, 'Стоимость операций при приёмке', 'приёмк', 'приемк')
    fine_col     = _get_wb_col(df, 'Общая сумма штрафов', 'штраф')
    other_col    = _get_wb_col(df, 'Прочие удержания', 'прочие выплаты')
    correct_col  = _get_wb_col(df, 'Корректировка Вознаграждения ВВ', 'корректировка вознаграждения')
    logistic_col = _get_wb_col(df, 'Возмещение издержек по перемещению', 'перемещени')

    # Фильтры по обоснованию
    if obosnov_col:
        mask_sale   = df[obosnov_col].astype(str).str.strip() == 'Продажа'
        mask_return = df[obosnov_col].astype(str).str.strip() == 'Возврат'
    else:
        mask_sale   = pd.Series([True] * len(df), index=df.index)
        mask_return = pd.Series([False] * len(df), index=df.index)

    def col_sum(col, mask=None):
        if col is None:
            return 0.0
        s = df[col].fillna(0)
        if mask is not None:
            s = s[mask]
        return float(s.sum())

    # ПРОДАЖИ
    qty_sold     = col_sum(qty_col, mask_sale)
    qty_returned = abs(col_sum(qty_col, mask_return))
    qty_net      = qty_sold - qty_returned

    retail_total = col_sum(retail_col)        # Розничная цена с учётом скидки (все строки)
    wb_real      = col_sum(wb_real_col)       # WB реализовал товар
    money_net    = col_sum(money_col)         # К перечислению за товар (нетто, все строки)

    # КОМИССИЯ = Розница − К перечислению
    commission   = retail_total - money_net if retail_total else 0.0
    commission_pct = (commission / retail_total) if retail_total else 0.0

    # РАСХОДЫ
    cost_delivery = col_sum(delivery_col)
    cost_storage  = col_sum(storage_col)
    cost_accept   = col_sum(accept_col)
    cost_fines    = col_sum(fine_col)
    cost_other    = col_sum(other_col)
    cost_correct  = col_sum(correct_col)
    cost_logistic = col_sum(logistic_col)
    total_costs   = cost_delivery + cost_storage + cost_accept + cost_fines + cost_other + cost_correct

    # ИТОГ К ОПЛАТЕ = К перечислению − расходы
    total_payment = money_net - total_costs
    payment_pct   = (total_payment / retail_total) if retail_total else 0.0

    return {
        # продажи
        'qty_sold':       qty_sold,
        'qty_returned':   qty_returned,
        'qty_net':        qty_net,
        'retail_total':   retail_total,
        'wb_real':        wb_real,
        'money_net':      money_net,
        # комиссия
        'commission':     commission,
        'commission_pct': commission_pct,
        # расходы
        'cost_delivery':  cost_delivery,
        'cost_storage':   cost_storage,
        'cost_accept':    cost_accept,
        'cost_fines':     cost_fines,
        'cost_other':     cost_other,
        'cost_correct':   cost_correct,
        'total_costs':    total_costs,
        # итог
        'total_payment':  total_payment,
        'payment_pct':    payment_pct,
    }


def _build_wb_svod(ws, df: pd.DataFrame, period: str, sebes_dict: dict):
    """Лист WB_Сводка — структура как в эталонном отчёте."""
    ws.sheet_view.showGridLines = False

    m = _extract_wb_metrics(df)

    VIOLET = COLORS['wb_violet']
    GREEN  = 'E2EFDA'   # светло-зелёный (ПРОДАЖИ)
    YELLOW = 'FFF2CC'   # светло-жёлтый  (КОМИССИЯ)
    PINK   = 'FCE4D6'   # светло-розовый (РАСХОДЫ)
    BEIGE  = 'F2F2F2'   # светло-серый   (ИТОГ)
    WHITE  = COLORS['row_white']

    def title_row(r, text, color, ncols=2):
        ws.merge_cells(f'A{r}:B{r}')
        c = ws.cell(row=r, column=1, value=text)
        c.fill = hfill(color)
        c.font = hfont(bold=True, size=11)
        c.alignment = left()
        c.border = BORDER_THIN
        ws.row_dimensions[r].height = 18

    def data_row(r, label, value, fmt='#,##0.00', bold=False, fill=WHITE):
        _write_cell(ws, r, 1, label, fill, bold=bold, align=left())
        _write_cell(ws, r, 2, value, fill, bold=bold, fmt=fmt)

    row = 1
    # Заголовок
    ws.merge_cells(f'A{row}:B{row}')
    c = ws.cell(row=row, column=1, value=f'Отчёт Wildberries — {period}')
    c.fill = hfill(VIOLET)
    c.font = hfont(bold=True, white=True, size=13)
    c.alignment = left()
    c.border = BORDER_THIN
    ws.row_dimensions[row].height = 24
    row += 1

    # Подзаголовок — период и номера отчётов
    ws.merge_cells(f'A{row}:B{row}')
    c = ws.cell(row=row, column=1,
                value=f'Период: сводный отчёт за {period} (нарастающим итогом)')
    c.fill = hfill(WHITE)
    c.font = Font(italic=True, size=9, name='Calibri')
    c.alignment = left()
    row += 2

    # ── ПРОДАЖИ ──
    title_row(row, 'ПРОДАЖИ', GREEN)
    row += 1
    data_row(row, 'Продано штук (продажи)',                   m['qty_sold'],     '#,##0', fill=WHITE)
    row += 1
    data_row(row, 'Возвращено штук (возвраты)',                m['qty_returned'], '#,##0', fill=WHITE)
    row += 1
    data_row(row, 'Продано нетто (продажи − возвраты), шт.',  m['qty_net'],      '#,##0', fill=WHITE)
    row += 1
    data_row(row, 'Розничная цена с учётом скидки, ₽',        m['retail_total'], fill=WHITE)
    row += 1
    data_row(row, 'WB реализовал товар (продажи клиентам), ₽', m['wb_real'],     fill=WHITE)
    row += 1
    data_row(row, 'К перечислению за товар (нетто), ₽',       m['money_net'],    fill=WHITE)
    row += 2

    # ── КОМИССИЯ ──
    title_row(row, 'КОМИССИЯ WILDBERRIES', YELLOW)
    row += 1
    data_row(row, 'Комиссия WB (розничная − к перечислению), ₽', m['commission'],     fill=WHITE)
    row += 1
    data_row(row, 'Комиссия WB, % от розничной цены со скидкой',  m['commission_pct'], fmt='0.0%', fill=WHITE)
    row += 2

    # ── РАСХОДЫ ──
    title_row(row, 'РАСХОДЫ (УСЛУГИ WB)', PINK)
    row += 1
    data_row(row, 'Стоимость логистики, ₽',             m['cost_delivery'], fill=WHITE)
    row += 1
    data_row(row, 'Стоимость хранения, ₽',              m['cost_storage'],  fill=WHITE)
    row += 1
    data_row(row, 'Стоимость операций при приёмке, ₽',  m['cost_accept'],   fill=WHITE)
    row += 1
    data_row(row, 'Общая сумма штрафов, ₽',             m['cost_fines'],    fill=WHITE)
    row += 1
    data_row(row, 'Прочие удержания/выплаты, ₽',        m['cost_other'],    fill=WHITE)
    row += 1
    data_row(row, 'Корректировка Вознаграждения ВВ, ₽', m['cost_correct'],  fill=WHITE)
    row += 1
    data_row(row, 'ИТОГО РАСХОДОВ, ₽',                  m['total_costs'],   bold=True,
             fill=COLORS['total_section'])
    row += 2

    # ── ИТОГ ──
    title_row(row, 'ИТОГ', BEIGE)
    row += 1
    data_row(row, 'ИТОГО К ОПЛАТЕ ПРОДАВЦУ, ₽',
             m['total_payment'], bold=True, fill=COLORS['total_section'])
    row += 1
    data_row(row, 'Итого к оплате, % от розничной цены со скидкой',
             m['payment_pct'], fmt='0.0%', fill=WHITE)

    _set_col_widths(ws, {'A': 52, 'B': 18})


def _build_wb_by_types(ws, df: pd.DataFrame, period: str, sebes_dict: dict):
    """Лист WB_По типам — продажи сгруппированы по типу товара с себестоимостью."""
    ws.sheet_view.showGridLines = False

    VIOLET = COLORS['wb_violet']
    LIGHT  = 'D9D2E9'
    WHITE  = COLORS['row_white']
    TOTAL  = COLORS['total_section']

    # Заголовок
    ws.merge_cells('A1:D1')
    c = ws['A1']
    c.value = f'Продажи по типам товара - WB, {period}'
    c.fill = hfill(VIOLET)
    c.font = hfont(bold=True, white=True, size=12)
    c.alignment = left()
    c.border = BORDER_THIN
    ws.row_dimensions[1].height = 22
    row = 3

    # Шапка
    hdrs = ['Тип товара', 'Кол-во, шт', 'Себестоимость за 1 шт', 'Себес итого']
    fills_h = [LIGHT, LIGHT, LIGHT, LIGHT]
    for ci, h in enumerate(hdrs, 1):
        _write_cell(ws, row, ci, h, LIGHT, bold=True, align=center())
    row += 1

    # Группировка по типу товара из sebes_dict
    obosnov_col = _get_wb_col(df, 'Обоснование для оплаты')
    name_col    = _get_wb_col(df, 'Название')
    qty_col     = _get_wb_col(df, 'Кол-во', 'количество')

    if obosnov_col:
        df_sales = df[df[obosnov_col].astype(str).str.strip() == 'Продажа'].copy()
    else:
        df_sales = df.copy()

    type_totals: dict[str, float] = {}
    for _, r in df_sales.iterrows():
        name = str(r.get(name_col, '') if name_col else '').strip()
        qty  = _to_float(r.get(qty_col, 0) if qty_col else 0)
        ptype = get_product_type(name, sebes_dict)
        if ptype:
            type_totals[ptype] = type_totals.get(ptype, 0) + qty

    # Записываем строки (сортировка по убыванию кол-ва)
    grand_qty   = 0.0
    grand_sebes = 0.0
    for ptype, qty in sorted(type_totals.items(), key=lambda x: -x[1]):
        sebes1 = sebes_dict.get(ptype, 0)
        sebes_all = sebes1 * qty
        _write_cell(ws, row, 1, ptype.capitalize(), WHITE, align=left())
        _write_cell(ws, row, 2, qty,      WHITE, fmt='#,##0')
        _write_cell(ws, row, 3, sebes1,   WHITE, fmt='#,##0.00')
        _write_cell(ws, row, 4, sebes_all, WHITE, fmt='#,##0.00')
        grand_qty   += qty
        grand_sebes += sebes_all
        row += 1

    # ИТОГО
    _write_cell(ws, row, 1, 'ИТОГО', TOTAL, bold=True, align=left())
    _write_cell(ws, row, 2, grand_qty,   TOTAL, bold=True, fmt='#,##0')
    _write_cell(ws, row, 3, None,        TOTAL)
    _write_cell(ws, row, 4, grand_sebes, TOTAL, bold=True, fmt='#,##0.00')

    _set_col_widths(ws, {'A': 30, 'B': 14, 'C': 22, 'D': 16})


def _build_wb_detail(ws, df: pd.DataFrame, sebes_dict: dict):
    ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:J1')
    c = ws['A1']
    c.value = 'WILDBERRIES — ДЕТАЛИЗАЦИЯ (все еженедельные отчёты)'
    c.fill = hfill(COLORS['wb_violet'])
    c.font = hfont(bold=True, white=True, size=12)
    c.alignment = center()

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
    for ci, h in enumerate(df_out.columns, 1):
        _write_cell(ws, row, ci, str(h), COLORS['wb_violet'], bold=True, white=True, align=center())
    row += 1

    for _, dr in df_out.iterrows():
        for ci, v in enumerate(dr.values, 1):
            v_out = None if (isinstance(v, float) and np.isnan(v)) else v
            _write_cell(ws, row, ci, v_out, COLORS['row_white'],
                        align=left() if ci <= 2 else right())
        row += 1

    _set_col_widths(ws, {'A': 40, 'B': 14, 'C': 14, 'D': 30, 'E': 18, 'F': 18,
                         'G': 8, 'H': 14, 'I': 16, 'J': 14, 'K': 12, 'L': 12, 'M': 16})


# ─────────────────────────────────────────────
#  ВСПОМОГАТЕЛЬНЫЕ
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
        for ci, v in enumerate(r.values, 1):
            v_out = None if (isinstance(v, float) and np.isnan(v)) else v
            c = ws.cell(row=start_row + i, column=ci, value=v_out)
            c.fill = hfill(COLORS['row_white'])
            c.border = BORDER_THIN
            c.alignment = left() if ci == 1 else right()


# ─────────────────────────────────────────────
#  STREAMLIT UI
# ─────────────────────────────────────────────

def render_sebes_editor():
    st.subheader('⚙️ Настройки себестоимости')
    st.caption('Здесь можно изменить себестоимость для любого типа товара.')

    sebes = load_sebes()

    _, col_reset = st.columns([3, 1])
    with col_reset:
        if st.button('↺ Сбросить к исходным', use_container_width=True):
            save_sebes(dict(DEFAULT_SEBES))
            st.success('Сброшено!')
            st.rerun()

    df_sebes = pd.DataFrame(
        [{'Тип товара (подстрока в названии)': k, 'Себестоимость, руб.': v}
         for k, v in sebes.items()]
    )

    edited = st.data_editor(
        df_sebes,
        num_rows='dynamic',
        use_container_width=True,
        column_config={
            'Тип товара (подстрока в названии)': st.column_config.TextColumn(width='large'),
            'Себестоимость, руб.': st.column_config.NumberColumn(min_value=0, format='%.2f', width='medium'),
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

    tab_main, tab_sebes = st.tabs(['📊 Формирование отчёта', '⚙️ Себестоимость'])

    with tab_sebes:
        sebes_dict = render_sebes_editor()

    with tab_main:
        st.markdown('Загрузите файлы из 1С, Ozon и Wildberries — получите единый Excel-отчёт.')

        col1, col2 = st.columns([1, 1])
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

        with st.expander('📅 Период фильтрации WB-отчётов', expanded=False):
            st.caption('WB-отчёты содержат данные за несколько месяцев. Укажите период.')
            import datetime as dt, calendar
            wb_col1, wb_col2 = st.columns(2)
            with wb_col1:
                default_from = dt.date(int(year), month_num, 1)
                wb_date_from = st.date_input('Дата С', value=default_from, key='wb_from')
            with wb_col2:
                last_day = calendar.monthrange(int(year), month_num)[1]
                default_to = dt.date(int(year), month_num, last_day)
                wb_date_to = st.date_input('Дата ПО', value=default_to, key='wb_to')

        st.divider()

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
        st.caption('Загрузите все еженедельные отчёты — лишние месяцы будут отфильтрованы')
        f_wb_zips = st.file_uploader('WB отчёты', type=['zip'],
                                      accept_multiple_files=True, key='wb')

        st.divider()

        if st.button('🚀 Сформировать отчёт', type='primary', use_container_width=True):
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
                    st.warning('⚠️ Не загружен ни один файл 1С')

                ozon_data = {}
                if f_ozon is not None:
                    try:
                        ozon_data = parse_ozon_file(f_ozon)
                        sheets = ozon_data.get('_sheet_names', [])
                        found  = [k for k in ozon_data if not k.startswith('_')]
                        st.success(f'✅ Ozon: листы в файле {sheets} → найдено: {found}')
                    except Exception as e:
                        import traceback
                        st.error(f'❌ Ошибка Ozon: {e}\n{traceback.format_exc()}')

                wb_data = {}
                if f_wb_zips:
                    wb_data = parse_wb_files(f_wb_zips, wb_date_from, wb_date_to)
                    n = len(wb_data.get('детализация', pd.DataFrame()))
                    st.success(f'✅ WB: {n} строк за выбранный период')

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
