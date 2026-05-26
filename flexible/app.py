import sys, os
sys.path.insert(0, os.path.dirname(__file__))  # ensure flexible/ is on path when run from repo root

import streamlit as st
import pandas as pd
from io import BytesIO
from ticket_parser import parse_ansource, list_sheets, detect_header, extract_vendor
from compare import compare, generate_excel

st.set_page_config(page_title='彈性票價比對工具', page_icon='🎫', layout='wide')


def _col_label(idx: int, columns: list) -> str:
    """Format column label for dropdown: '欄2（票種）' or '欄2'."""
    val = columns[idx] if idx < len(columns) else ''
    return f'欄{idx}（{val}）' if val and val != 'nan' else f'欄{idx}'


def _file_bytes(uploaded_file) -> BytesIO:
    """Read uploaded file to BytesIO so it can be re-read multiple times."""
    data = uploaded_file.read()
    uploaded_file.seek(0)
    return BytesIO(data)


def main():
    st.title('🎫 彈性票價比對工具')

    # ── ① 上傳檔案 ────────────────────────────────────────────────────────────
    st.subheader('① 上傳檔案')
    col1, col2 = st.columns(2)
    with col1:
        ay_file = st.file_uploader('安源 .xls', type=['xls'])
    with col2:
        vs_file = st.file_uploader('廠商 .xlsx', type=['xlsx'])

    if not ay_file or not vs_file:
        st.info('請上傳兩份 Excel 檔案後繼續')
        return

    # Cache 安源 parse result by filename+size
    ay_cache_key = f'ay_{ay_file.name}_{ay_file.size}'
    if ay_cache_key not in st.session_state:
        with st.spinner('讀取安源檔案...'):
            try:
                st.session_state[ay_cache_key] = parse_ansource(_file_bytes(ay_file))
            except Exception as e:
                st.error(f'安源檔案讀取失敗：{e}')
                return
    ay_data = st.session_state[ay_cache_key]

    # Cache vendor file bytes (avoid re-reading upload stream)
    vs_bytes_key = f'vs_bytes_{vs_file.name}_{vs_file.size}'
    if vs_bytes_key not in st.session_state:
        st.session_state[vs_bytes_key] = vs_file.read()
        vs_file.seek(0)

    # ── ② 廠商欄位確認 ────────────────────────────────────────────────────────
    st.markdown('---')
    st.subheader('② 廠商欄位確認')

    sheets_key = f'sheets_{vs_file.name}_{vs_file.size}'
    if sheets_key not in st.session_state:
        st.session_state[sheets_key] = list_sheets(BytesIO(st.session_state[vs_bytes_key]))
    sheets = st.session_state[sheets_key]

    selected_sheet = st.selectbox('選擇廠商頁籤', sheets)

    # Load raw sheet + run auto-detection (cached per sheet)
    raw_key = f'raw_{vs_file.name}_{vs_file.size}_{selected_sheet}'
    if raw_key not in st.session_state:
        df_raw = pd.read_excel(
            BytesIO(st.session_state[vs_bytes_key]),
            sheet_name=selected_sheet,
            header=None,
            engine='openpyxl',
        )
        detected = detect_header(df_raw)
        st.session_state[raw_key] = df_raw
        st.session_state[f'detected_{raw_key}'] = detected

    df_raw: pd.DataFrame = st.session_state[raw_key]
    detected: dict = st.session_state[f'detected_{raw_key}']
    columns = detected['columns']
    n_cols = len(df_raw.columns)

    # Show preview (header row + up to 10 data rows)
    preview_start = max(0, detected['header_row'])
    preview_df = df_raw.iloc[preview_start: preview_start + 11].copy()
    preview_df.columns = [_col_label(i, columns) for i in range(n_cols)]
    st.dataframe(preview_df, use_container_width=True)

    # Warn if no keywords found (all col indices are None)
    if detected['area_col'] is None or detected['price_col'] is None:
        st.warning('⚠️ 未完整偵測到欄位，請手動確認「區域欄」與「票價欄」')

    # Dropdowns: None option + integer column indices
    NONE_LABEL = '（無）'

    def default_idx(col_idx):
        """Return selectbox list index (0 = NONE_LABEL, 1 = 欄0, 2 = 欄1, ...)."""
        return 0 if col_idx is None else col_idx + 1

    c1, c2, c3 = st.columns(3)
    with c1:
        area_sel = st.selectbox(
            '區域欄 *',
            [NONE_LABEL] + list(range(n_cols)),
            index=default_idx(detected['area_col']),
            format_func=lambda x: NONE_LABEL if x == NONE_LABEL else _col_label(x, columns),
        )
    with c2:
        ticket_sel = st.selectbox(
            '票種欄（選填）',
            [NONE_LABEL] + list(range(n_cols)),
            index=default_idx(detected['ticket_col']),
            format_func=lambda x: NONE_LABEL if x == NONE_LABEL else _col_label(x, columns),
        )
    with c3:
        price_sel = st.selectbox(
            '票價欄 *',
            [NONE_LABEL] + list(range(n_cols)),
            index=default_idx(detected['price_col']),
            format_func=lambda x: NONE_LABEL if x == NONE_LABEL else _col_label(x, columns),
        )

    area_idx   = None if area_sel   == NONE_LABEL else area_sel
    ticket_idx = None if ticket_sel == NONE_LABEL else ticket_sel
    price_idx  = None if price_sel  == NONE_LABEL else price_sel

    if st.button('✅ 確認並提取廠商資料', type='primary'):
        if area_idx is None or price_idx is None:
            st.error('❌ 請至少指定「區域欄」和「票價欄」')
        else:
            with st.spinner('提取中...'):
                try:
                    vs_df = extract_vendor(
                        df_raw, detected['header_row'],
                        area_idx, ticket_idx, price_idx
                    )
                    st.session_state['vs_data'] = vs_df
                    st.session_state['vs_sheet_label'] = selected_sheet
                    # Clear stale compare result when vendor data changes
                    st.session_state.pop('compare_result', None)
                    data_rows = len(df_raw) - detected['header_row'] - 1
                    skipped = max(0, data_rows - len(vs_df))
                    msg = f'✅ 提取成功：{len(vs_df)} 筆有效資料'
                    if skipped > 0:
                        msg += f'（已略過 {skipped} 筆無效列）'
                    st.success(msg)
                except Exception as e:
                    st.error(f'提取失敗：{e}')

    if 'vs_data' not in st.session_state:
        return

    # ── ③ 選擇比對場次 ────────────────────────────────────────────────────────
    st.markdown('---')
    st.subheader('③ 選擇比對場次')
    col_a, col_b = st.columns(2)
    with col_a:
        ay_choice = st.selectbox('安源場次日期', list(ay_data.keys()))
    with col_b:
        st.text_input(
            '廠商資料來源',
            value=st.session_state.get('vs_sheet_label', ''),
            disabled=True,
        )

    # ── ④ 比對 ───────────────────────────────────────────────────────────────
    st.markdown('---')
    if st.button('🔍 開始比對', type='primary'):
        with st.spinner('比對中...'):
            try:
                result = compare(ay_data[ay_choice], st.session_state['vs_data'])
                st.session_state['compare_result'] = result
            except Exception as e:
                st.error(f'比對失敗：{e}')

    if 'compare_result' in st.session_state:
        _show_results(st.session_state['compare_result'])


def _show_results(result: dict):
    st.subheader('④ 比對結果')
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric('✅ 相符',          len(result['price_ok']))
    c2.metric('❌ 票價不符',      len(result['price_diff']))
    c3.metric('⚠️ 廠商有安源無',  len(result['only_vs']))
    c4.metric('⚠️ 安源有廠商無',  len(result['only_ay']))
    c5.metric('ℹ️ 無對應票種',    len(result['unmapped']))

    if len(result['price_diff']) > 0:
        with st.expander(f'❌ 票價不符 — {len(result["price_diff"])} 筆', expanded=True):
            d = result['price_diff'][['area', 'ticket_vs', 'ticket_ay', 'price_vs', 'price_ay']].copy()
            d.columns = ['區域名稱', '廠商票種', '安源票種', '廠商票價', '安源票價']
            st.dataframe(d, use_container_width=True, hide_index=True)

    if len(result['only_vs']) > 0:
        with st.expander(f'⚠️ 廠商有、安源沒有 — {len(result["only_vs"])} 筆'):
            v = result['only_vs'][['area', 'ticket_vs', 'price_vs']].copy()
            v.columns = ['區域名稱', '廠商票種', '廠商票價']
            st.dataframe(v, use_container_width=True, hide_index=True)

    if len(result['only_ay']) > 0:
        with st.expander(f'⚠️ 安源有、廠商沒有 — {len(result["only_ay"])} 筆'):
            a = result['only_ay'][['area', 'ticket_ay', 'price_ay']].copy()
            a.columns = ['區域名稱', '安源票種', '安源票價']
            st.dataframe(a, use_container_width=True, hide_index=True)

    if len(result['unmapped']) > 0:
        with st.expander(f'ℹ️ 安源票種無對應廠商 — {len(result["unmapped"])} 筆'):
            u = result['unmapped'][['area', 'ticket', 'price']].copy()
            u.columns = ['區域名稱', '安源票種', '安源票價']
            st.dataframe(u, use_container_width=True, hide_index=True)

    st.markdown('---')
    try:
        excel_bytes = generate_excel(result)
        st.download_button(
            label='📥 下載 Excel 報告',
            data=excel_bytes,
            file_name='彈性比對結果.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type='primary',
        )
    except Exception as e:
        st.error(f'Excel 報告產生失敗：{e}')


if __name__ == '__main__':
    main()
