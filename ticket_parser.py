import pandas as pd
from io import BytesIO


# ─────────────────────────────────────────────
# 安源固定格式解析
# ─────────────────────────────────────────────

def parse_ansource(file) -> dict:
    """讀取安源 .xls，回傳 {date_label: DataFrame(area, ticket, price)}"""
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


# ─────────────────────────────────────────────
# 廠商彈性格式解析
# ─────────────────────────────────────────────

def list_sheets(file) -> list:
    """回傳 xlsx 所有頁籤名稱。"""
    xl = pd.ExcelFile(file, engine='openpyxl')
    return xl.sheet_names


# Keywords ordered from specific to broad to avoid false matches
_AREA_KW   = ['區域名稱', '區域', 'Area', 'area', '座位區']
_TICKET_KW = ['票種名稱', '票種', 'Ticket', 'ticket', '種類']
_PRICE_KW  = ['票價', '金額', 'Price', 'price', '費用']
_ALL_KW    = _AREA_KW + _TICKET_KW + _PRICE_KW


def _col_type(cell_str: str) -> str | None:
    """Return 'ticket', 'price', 'area', or None — checked in priority order."""
    s = str(cell_str).strip()
    if any(kw in s for kw in _TICKET_KW):
        return 'ticket'
    if any(kw in s for kw in _PRICE_KW):
        return 'price'
    if any(kw in s for kw in _AREA_KW):
        return 'area'
    return None


def detect_header(df_raw: pd.DataFrame) -> dict:
    """Scan the first 20 rows for the header row containing area/ticket/price keywords.

    Returns:
        {
            'header_row': int,          # best candidate row index
            'area_col':   int | None,
            'ticket_col': int | None,
            'price_col':  int | None,
            'columns':    list[str],    # cell values of header row (for UI dropdowns)
        }

    If no keywords are found in the first 20 rows, falls back to row 0 with
    area_col=None, ticket_col=None, price_col=None. The UI will show a warning.
    """
    if len(df_raw) == 0:
        return {'header_row': 0, 'area_col': None, 'ticket_col': None,
                'price_col': None, 'columns': []}

    scan_limit = min(20, len(df_raw))
    best_row, best_score = 0, 0

    for i in range(scan_limit):
        row = df_raw.iloc[i]
        score = sum(
            1 for cell in row
            if pd.notna(cell) and any(kw in str(cell) for kw in _ALL_KW)
        )
        if score > best_score:
            best_score = score
            best_row = i

    header_vals = df_raw.iloc[best_row]
    columns = [str(v).strip() if pd.notna(v) else '' for v in header_vals]

    area_col = ticket_col = price_col = None
    for col_idx, cell_str in enumerate(columns):
        t = _col_type(cell_str)
        if t == 'ticket' and ticket_col is None:
            ticket_col = col_idx
        elif t == 'price' and price_col is None:
            price_col = col_idx
        elif t == 'area' and area_col is None:
            area_col = col_idx

    # Fallback: detect price_col by numeric column content
    if price_col is None:
        data_sample = df_raw.iloc[best_row + 1: best_row + 11]
        for col_idx in range(len(df_raw.columns)):
            col_data = data_sample.iloc[:, col_idx]
            numeric_count = pd.to_numeric(col_data, errors='coerce').notna().sum()
            total = len(col_data.dropna())
            if total > 0 and numeric_count / total >= 0.5:
                price_col = col_idx
                break

    return {
        'header_row': best_row,
        'area_col':   area_col,
        'ticket_col': ticket_col,
        'price_col':  price_col,
        'columns':    columns,
    }


def extract_vendor(
    df_raw: pd.DataFrame,
    header_row: int,
    area_col: int | None,
    ticket_col: int | None,
    price_col: int | None,
    default_ticket: str = '全票',
) -> pd.DataFrame:
    """Extract vendor data from a raw DataFrame using confirmed column positions.

    Args:
        df_raw:         Raw DataFrame loaded with header=None.
        header_row:     Row index of the header row (data starts at header_row+1).
        area_col:       Column index for area name (required — must not be None).
        ticket_col:     Column index for ticket type; None → use default_ticket for all rows.
        price_col:      Column index for price (required — must not be None).
        default_ticket: Ticket type to use when ticket_col is None. Default: '全票'.

    Returns:
        DataFrame with columns: area, ticket_vs, ticket_std, price
    """
    if area_col is None or price_col is None:
        raise ValueError("area_col and price_col are required (must not be None)")
    data = df_raw.iloc[header_row + 1:].copy().reset_index(drop=True)

    df = pd.DataFrame()
    df['area'] = data.iloc[:, area_col]
    df['ticket_vs'] = (data.iloc[:, ticket_col]
                       if ticket_col is not None
                       else default_ticket)
    df['price_raw'] = data.iloc[:, price_col]

    # Fill merged area cells (merged cells appear as NaN in rows below)
    df['area'] = df['area'].ffill()

    # Clean price: coerce non-numeric to NaN then drop
    df['price'] = pd.to_numeric(df['price_raw'], errors='coerce')
    df = df.dropna(subset=['area', 'price'])
    df['price'] = df['price'].astype(int)

    # Strip whitespace and remove sentinel values
    df['area']      = df['area'].astype(str).str.strip()
    df['ticket_vs'] = df['ticket_vs'].astype(str).str.strip()
    df = df[~df['area'].isin(['', 'nan'])]
    df = df[~df['ticket_vs'].isin(['', 'nan'])]

    # ticket_std: strip "(不顯示)" suffix for use as comparison key
    df['ticket_std'] = (df['ticket_vs']
                        .str.replace(r'\(不顯示\)', '', regex=True)
                        .str.strip())

    return df[['area', 'ticket_vs', 'ticket_std', 'price']].drop_duplicates().reset_index(drop=True)
