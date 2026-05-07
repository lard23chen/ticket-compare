import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

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
