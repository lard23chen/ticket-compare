import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

TICKET_MAP = {
    '全票':           '全票',
    '貴賓券':          '貴賓券',
    '公關票':          '公關票',
    '信用卡優惠票':     '信用卡優惠票',
    '啦啦隊票':        '啦啦隊票',
    '眷屬票':          '眷屬票',
    '內野優惠票':       '內野優惠票',
    '內野身心優惠票':   '身心優惠票',
    '外野身心優惠票':   '身心優惠票',
    '內野半票':         '半票',
    '外野半票':         '半票',
    '統一獅會員優惠票':  '會員優惠票(折100)',
}


def parse_ansource(file) -> dict:
    """讀取安源.xls，回傳 {date_label: DataFrame(area, ticket, price)}"""
    raw = pd.read_excel(file, sheet_name='票區票種資料(劃位)',
                        engine='xlrd', header=None)

    sections = []
    for i, row in raw.iterrows():
        for cell in row:
            if pd.notna(cell) and '演出日期' in str(cell):
                date_str = str(cell).replace('演出日期:', '').strip()
                sections.append((i, date_str))
                break

    header_rows = [
        i for i, row in raw.iterrows()
        if '票種編號' in row.values
    ]

    result = {}
    for idx, (sec_start, date_str) in enumerate(sections):
        data_header = next((hr for hr in header_rows if hr > sec_start), None)
        if data_header is None:
            continue
        data_start = data_header + 1
        data_end = sections[idx + 1][0] if idx + 1 < len(sections) else len(raw)

        df = raw.iloc[data_start:data_end].reset_index(drop=True)[[4, 8, 9]].copy()
        df.columns = ['area', 'ticket', 'price']
        df['area'] = df['area'].ffill()
        df = df.dropna(subset=['ticket'])
        df['price'] = pd.to_numeric(df['price'], errors='coerce')
        df = df.dropna(subset=['price'])
        df['area']   = df['area'].astype(str).str.strip()
        df['ticket'] = df['ticket'].astype(str).str.strip().str.rstrip('.')
        df['price']  = df['price'].astype(int)
        df = df[~df['ticket'].isin(['票種', 'nan'])]
        df = df.drop_duplicates()
        result[date_str] = df

    return result


def parse_vendor(file) -> dict:
    """讀取廠商.xlsx，回傳含「票價」的頁籤 {sheet_name: DataFrame(area, ticket_vs, ticket_std, price)}"""
    xl = pd.ExcelFile(file, engine='openpyxl')
    result = {}
    for sheet in xl.sheet_names:
        if '票價' not in sheet:
            continue
        raw = pd.read_excel(file, sheet_name=sheet,
                            engine='openpyxl', header=None)
        # 動態找出 row 2 中 col 2 起有值的票種欄位
        ticket_row = raw.iloc[2]
        ticket_cols = [c for c in range(2, len(ticket_row))
                       if pd.notna(ticket_row[c]) and str(ticket_row[c]).strip() != '']
        ticket_types = [ticket_row[c] for c in ticket_cols]

        df_wide = raw.iloc[3:].copy()
        available_cols = [c for c in [1] + ticket_cols if c in df_wide.columns]
        df_wide = df_wide[available_cols].copy()
        df_wide.columns = ['area'] + ticket_types[:len(available_cols) - 1]
        df_wide = df_wide.dropna(subset=['area'])
        df_wide['area'] = df_wide['area'].astype(str).str.strip()

        df = df_wide.melt(id_vars=['area'], var_name='ticket_vs', value_name='price')
        df = df[df['price'] != '-'].dropna(subset=['price'])
        df['price'] = pd.to_numeric(df['price'], errors='coerce')
        df = df.dropna(subset=['price'])
        df['price'] = df['price'].astype(int)
        df['ticket_std'] = (df['ticket_vs']
                            .str.replace(r'\(不顯示\)', '', regex=True)
                            .str.strip())
        df = df.drop_duplicates()
        result[sheet] = df
    return result


def compare(df_ay: pd.DataFrame, df_vs: pd.DataFrame) -> dict:
    """比對安源與廠商資料，回傳分類結果 dict"""
    df_ay = df_ay.copy()
    df_ay['ticket_std'] = df_ay['ticket'].map(TICKET_MAP)

    ay_m = (df_ay[df_ay['ticket_std'].notna()]
            [['area', 'ticket', 'ticket_std', 'price']]
            .rename(columns={'price': 'price_ay', 'ticket': 'ticket_ay'}))

    vs_m = (df_vs[['area', 'ticket_vs', 'ticket_std', 'price']]
            .rename(columns={'price': 'price_vs'}))

    merged = pd.merge(ay_m, vs_m, on=['area', 'ticket_std'],
                      how='outer', indicator=True)

    both     = merged[merged['_merge'] == 'both'].copy()
    only_ay  = merged[merged['_merge'] == 'left_only'][['area', 'ticket_ay', 'ticket_std', 'price_ay']].copy()
    only_vs  = merged[merged['_merge'] == 'right_only'][['area', 'ticket_vs', 'ticket_std', 'price_vs']].copy()

    return {
        'price_ok':   both[both['price_ay'] == both['price_vs']].copy(),
        'price_diff': both[both['price_ay'] != both['price_vs']].copy(),
        'only_ay':    only_ay,
        'only_vs':    only_vs,
        'unmapped':   df_ay[df_ay['ticket_std'].isna()][['area', 'ticket', 'price']].copy(),
    }


def generate_excel(result: dict) -> bytes:
    """將比對結果產出為 Excel bytes（in-memory，不落地）"""
    wb = Workbook()
    GREEN  = PatternFill('solid', fgColor='C6EFCE')
    RED    = PatternFill('solid', fgColor='FFC7CE')
    ORANGE = PatternFill('solid', fgColor='FFEB9C')
    BLUE   = PatternFill('solid', fgColor='BDD7EE')
    GRAY   = PatternFill('solid', fgColor='D9D9D9')
    HFILL  = PatternFill('solid', fgColor='4472C4')

    def write_sheet(ws, df, cols, col_names, fill, title, col_widths):
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        r = 1
        ws.cell(r, 1, title).font = Font(bold=True, size=11, color='FFFFFF')
        ws.cell(r, 1).fill = HFILL
        ws.merge_cells(start_row=r, start_column=1,
                       end_row=r, end_column=len(cols))
        ws.cell(r, 1).alignment = Alignment(horizontal='center')
        r += 1
        for c, name in enumerate(col_names, 1):
            cell = ws.cell(r, c, name)
            cell.font = Font(bold=True)
            cell.fill = GRAY
            cell.alignment = Alignment(horizontal='center')
        r += 1
        for row in df[cols].itertuples(index=False):
            for c, val in enumerate(row, 1):
                ws.cell(r, c, val).fill = fill
            r += 1

    price_ok   = result['price_ok']
    price_diff = result['price_diff']
    only_ay    = result['only_ay']
    only_vs    = result['only_vs']
    unmapped   = result['unmapped']

    # 摘要
    ws0 = wb.active
    ws0.title = '摘要'
    ws0.column_dimensions['A'].width = 38
    ws0.column_dimensions['B'].width = 12
    for r, (label, count, fill) in enumerate([
        ('✅ 票價完全相符',      len(price_ok),   GREEN),
        ('❌ 票價不符',          len(price_diff), RED),
        ('⚠️ 廠商有、安源沒有',  len(only_vs),    ORANGE),
        ('⚠️ 安源有、廠商沒有',  len(only_ay),    ORANGE),
        ('ℹ️ 安源票種無對應廠商', len(unmapped),  BLUE),
    ], 1):
        ws0.cell(r, 1, label).fill = fill
        ws0.cell(r, 2, count).fill = fill

    if len(price_diff) > 0:
        ws = wb.create_sheet('❌票價不符')
        write_sheet(ws, price_diff.sort_values(['ticket_std', 'area']),
            ['area', 'ticket_vs', 'ticket_ay', 'price_vs', 'price_ay'],
            ['區域名稱', '廠商票種', '安源票種', '廠商票價', '安源票價'],
            RED, f'❌ 票價不符 ({len(price_diff)}筆)', [42, 20, 20, 12, 12])

    if len(only_vs) > 0:
        ws = wb.create_sheet('⚠️廠商有安源無')
        write_sheet(ws, only_vs.sort_values(['ticket_std', 'area']),
            ['area', 'ticket_vs', 'price_vs'],
            ['區域名稱', '廠商票種', '廠商票價'],
            ORANGE, f'⚠️ 廠商有、安源沒有 ({len(only_vs)}筆)', [42, 22, 12])

    if len(only_ay) > 0:
        ws = wb.create_sheet('⚠️安源有廠商無')
        write_sheet(ws, only_ay.sort_values(['area', 'ticket_ay']),
            ['area', 'ticket_ay', 'price_ay'],
            ['區域名稱', '安源票種', '安源票價'],
            ORANGE, f'⚠️ 安源有、廠商沒有 ({len(only_ay)}筆)', [42, 22, 12])

    if len(unmapped) > 0:
        ws = wb.create_sheet('ℹ️安源無對應')
        write_sheet(ws, unmapped.sort_values(['ticket', 'area']),
            ['area', 'ticket', 'price'],
            ['區域名稱', '安源票種', '安源票價'],
            BLUE, f'ℹ️ 安源票種無對應廠商 ({len(unmapped)}筆)', [42, 22, 12])

    if len(price_ok) > 0:
        ws = wb.create_sheet('✅完全相符')
        ok_display = (price_ok.sort_values(['ticket_std', 'area'])
                              [['area', 'ticket_vs', 'ticket_ay', 'price_vs']]
                              .copy()
                              .rename(columns={'price_vs': 'price'}))
        write_sheet(ws, ok_display,
            ['area', 'ticket_vs', 'ticket_ay', 'price'],
            ['區域名稱', '廠商票種', '安源票種', '票價'],
            GREEN, f'✅ 票價完全相符 ({len(price_ok)}筆)', [42, 22, 22, 12])

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
