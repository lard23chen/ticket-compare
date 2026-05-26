import re

import pandas as pd
import streamlit as st
from io import BytesIO

from ticket_parser import (
    parse_ansource, list_sheets, detect_header,
    extract_vendor, extract_vendor_wide,
)
from compare import compare, generate_excel

st.set_page_config(page_title='彈性票價比對工具', page_icon='🎫', layout='wide')


# ── 區域名正規化工具 ──────────────────────────────────────────────────────────

def _norm_vs_area(s: str) -> str:
    """廠商區域名正規化：去除括號後綴。例：'巨浪席(對號)' → '巨浪席'"""
    return re.sub(r'\s*[（(][^）)]*[）)]\s*$', '', str(s)).strip()


def _norm_ay_area(s: str) -> str:
    """安源區域名正規化：去除尾部數字。例：'巨浪席1' → '巨浪席'"""
    return re.sub(r'[0-9]+$', '', str(s)).strip()


# ── 通用輔助 ─────────────────────────────────────────────────────────────────

def _col_label(idx: int, columns: list) -> str:
    val = columns[idx] if idx < len(columns) else ''
    return f'欄{idx}（{val}）' if val and val != 'nan' else f'欄{idx}'


def _file_bytes(uploaded_file) -> BytesIO:
    data = uploaded_file.read()
    uploaded_file.seek(0)
    return BytesIO(data)


# ── 主程式 ───────────────────────────────────────────────────────────────────

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

    # 安源解析（按檔案快取）
    ay_cache_key = f'ay_{ay_file.name}_{ay_file.size}'
    if ay_cache_key not in st.session_state:
        with st.spinner('讀取安源檔案...'):
            try:
                st.session_state[ay_cache_key] = parse_ansource(_file_bytes(ay_file))
            except Exception as e:
                st.error(f'安源檔案讀取失敗：{e}')
                return
    ay_data = st.session_state[ay_cache_key]

    # 廠商原始 bytes 快取
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

    raw_key = f'raw_{vs_file.name}_{vs_file.size}_{selected_sheet}'
    if raw_key not in st.session_state:
        df_raw = pd.read_excel(
            BytesIO(st.session_state[vs_bytes_key]),
            sheet_name=selected_sheet, header=None, engine='openpyxl',
        )
        detected = detect_header(df_raw)
        st.session_state[raw_key] = df_raw
        st.session_state[f'detected_{raw_key}'] = detected

    df_raw: pd.DataFrame = st.session_state[raw_key]
    detected: dict = st.session_state[f'detected_{raw_key}']
    columns = detected['columns']
    n_cols = len(df_raw.columns)

    # 預覽（標題列 + 最多 10 列資料）
    preview_start = max(0, detected['header_row'])
    preview_df = df_raw.iloc[preview_start: preview_start + 11].copy()
    preview_df.columns = [_col_label(i, columns) for i in range(n_cols)]
    st.dataframe(preview_df, use_container_width=True)

    if detected['area_col'] is None or detected['price_col'] is None:
        st.warning('⚠️ 未完整偵測到欄位，請手動確認「區域欄」與「票價欄」')

    NONE_LABEL = '（無）'

    def default_idx(col_idx):
        return 0 if col_idx is None else col_idx + 1

    c1, c2, c3 = st.columns(3)
    with c1:
        area_sel = st.selectbox(
            '區域欄 *', [NONE_LABEL] + list(range(n_cols)),
            index=default_idx(detected['area_col']),
            format_func=lambda x: NONE_LABEL if x == NONE_LABEL else _col_label(x, columns),
        )
    with c2:
        ticket_sel = st.selectbox(
            '票種欄（選填）', [NONE_LABEL] + list(range(n_cols)),
            index=default_idx(detected['ticket_col']),
            format_func=lambda x: NONE_LABEL if x == NONE_LABEL else _col_label(x, columns),
        )
    with c3:
        price_sel = st.selectbox(
            '票價欄 *', [NONE_LABEL] + list(range(n_cols)),
            index=default_idx(detected['price_col']),
            format_func=lambda x: NONE_LABEL if x == NONE_LABEL else _col_label(x, columns),
        )

    area_idx   = None if area_sel   == NONE_LABEL else area_sel
    ticket_idx = None if ticket_sel == NONE_LABEL else ticket_sel
    price_idx  = None if price_sel  == NONE_LABEL else price_sel

    if st.button('✅ 確認並提取廠商資料', type='primary'):
        is_wide = detected.get('wide_format') and detected.get('sub_header_row') is not None
        if area_idx is None:
            st.error('❌ 請至少指定「區域欄」')
        elif not is_wide and price_idx is None:
            st.error('❌ 請至少指定「區域欄」和「票價欄」')
        else:
            with st.spinner('提取中...'):
                try:
                    if is_wide:
                        vs_df = extract_vendor_wide(
                            df_raw, detected['header_row'],
                            area_idx, detected['sub_header_row']
                        )
                    else:
                        vs_df = extract_vendor(
                            df_raw, detected['header_row'],
                            area_idx, ticket_idx, price_idx
                        )
                    st.session_state['vs_data'] = vs_df
                    st.session_state['vs_sheet_label'] = selected_sheet
                    # 清除下游狀態
                    for k in ['ticket_map', 'area_map', 'compare_result']:
                        st.session_state.pop(k, None)
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

    # ── ③ 票種對應設定 ────────────────────────────────────────────────────────
    st.markdown('---')
    st.subheader('③ 票種對應設定')
    st.caption('將安源票種對應到廠商票種（無對應的票種比對時將歸入「無對應票種」）')

    # 安源所有票種（跨場次聯集）
    all_ay_tickets = sorted({
        t for df in ay_data.values() for t in df['ticket'].unique()
    })
    # 廠商票種
    vs_tickets = sorted(st.session_state['vs_data']['ticket_std'].unique())

    col_info1, col_info2 = st.columns(2)
    col_info1.info(f'安源共 **{len(all_ay_tickets)}** 種票種')
    col_info2.info(f'廠商共 **{len(vs_tickets)}** 種票種')

    NONE_T = '（無對應）'
    COLS_PER_ROW = 3
    for i in range(0, len(all_ay_tickets), COLS_PER_ROW):
        batch = all_ay_tickets[i:i + COLS_PER_ROW]
        row_cols = st.columns(COLS_PER_ROW)
        for col, ay_t in zip(row_cols, batch):
            with col:
                options = [NONE_T] + vs_tickets
                default = ay_t if ay_t in vs_tickets else NONE_T
                st.selectbox(
                    f'安源：{ay_t}',
                    options,
                    index=options.index(default) if default in options else 0,
                    key=f'tmap_{ay_t}',
                )

    if st.button('✅ 確認票種對應', type='primary', key='btn_ticket_map'):
        tmap = {}
        for ay_t in all_ay_tickets:
            sel = st.session_state.get(f'tmap_{ay_t}', NONE_T)
            if sel != NONE_T:
                tmap[ay_t] = sel
        st.session_state['ticket_map'] = tmap
        for k in ['area_map', 'compare_result']:
            st.session_state.pop(k, None)
        unmapped_n = len(all_ay_tickets) - len(tmap)
        msg = f'✅ 已設定 {len(tmap)} 項票種對應'
        if unmapped_n:
            msg += f'，{unmapped_n} 項設為無對應'
        st.success(msg)

    if 'ticket_map' not in st.session_state:
        st.info('⬆️ 請確認票種對應後繼續')
        return

    # 顯示已確認的票種對應摘要
    tmap_confirmed = st.session_state['ticket_map']
    with st.expander('📋 已確認票種對應', expanded=False):
        tmap_rows = [
            {'安源票種': ay_t, '→ 廠商票種': tmap_confirmed.get(ay_t, '（無對應）')}
            for ay_t in all_ay_tickets
        ]
        st.dataframe(pd.DataFrame(tmap_rows), use_container_width=True, hide_index=True)

    # ── ④ 選擇比對場次 ────────────────────────────────────────────────────────
    st.markdown('---')
    st.subheader('④ 選擇比對場次')
    col_a, col_b = st.columns(2)
    with col_a:
        ay_choice = st.selectbox('安源場次日期', list(ay_data.keys()))
    with col_b:
        st.text_input(
            '廠商資料來源',
            value=st.session_state.get('vs_sheet_label', ''),
            disabled=True,
        )

    # 換場次時清除區域對應與比對結果
    if st.session_state.get('_last_ay_choice') != ay_choice:
        st.session_state.pop('area_map', None)
        st.session_state.pop('compare_result', None)
        st.session_state['_last_ay_choice'] = ay_choice

    # ── ⑤ 區域對應設定 ────────────────────────────────────────────────────────
    st.markdown('---')
    st.subheader('⑤ 區域對應設定')
    st.caption('將安源區域名對應到廠商區域名（自動去除括號後綴與尾部數字）')

    # 廠商區域正規化後的唯一清單
    vs_area_keys = sorted({
        _norm_vs_area(a) for a in st.session_state['vs_data']['area']
    })
    # 本場次安源區域
    ay_areas = sorted(ay_data[ay_choice]['area'].unique())

    # 自動對應
    auto_map: dict[str, str] = {}
    need_manual: list[str] = []
    for ay_a in ay_areas:
        if ay_a in vs_area_keys:                    # 完全吻合
            auto_map[ay_a] = ay_a
        elif _norm_ay_area(ay_a) in vs_area_keys:   # 去掉尾部數字後吻合
            auto_map[ay_a] = _norm_ay_area(ay_a)
        else:
            need_manual.append(ay_a)

    # 顯示自動對應結果
    auto_label = f'🔄 自動對應 {len(auto_map)} / {len(ay_areas)} 個區域'
    with st.expander(auto_label, expanded=(len(need_manual) == 0)):
        if auto_map:
            auto_rows = [
                {'安源區域': k, '→ 廠商對應區域': v}
                for k, v in auto_map.items()
            ]
            st.dataframe(pd.DataFrame(auto_rows), use_container_width=True, hide_index=True)
        else:
            st.info('無法自動對應任何區域')

    # 手動指定未配對的區域
    SKIP = '（略過不比對）'
    if need_manual:
        st.warning(f'⚠️ 以下 **{len(need_manual)}** 個區域無法自動對應，請手動指定（或選「略過」）')
        man_cols = st.columns(min(3, len(need_manual)))
        for i, ay_a in enumerate(need_manual):
            with man_cols[i % 3]:
                options = [SKIP] + vs_area_keys
                st.selectbox(f'安源「{ay_a}」→', options, key=f'amap_{ay_a}')

    if st.button('✅ 確認區域對應', type='primary', key='btn_area_map'):
        area_map = dict(auto_map)
        for ay_a in need_manual:
            sel = st.session_state.get(f'amap_{ay_a}', SKIP)
            if sel != SKIP:
                area_map[ay_a] = sel
        st.session_state['area_map'] = area_map
        st.session_state.pop('compare_result', None)
        skipped_n = len(ay_areas) - len(area_map)
        msg = f'✅ 已確認 {len(area_map)} 個區域對應'
        if skipped_n:
            msg += f'，{skipped_n} 個略過不比對'
        st.success(msg)

    if 'area_map' not in st.session_state:
        st.info('⬆️ 請確認區域對應後繼續')
        return

    # ── ⑥ 比對 ───────────────────────────────────────────────────────────────
    st.markdown('---')
    if st.button('🔍 開始比對', type='primary'):
        with st.spinner('比對中...'):
            try:
                # 廠商區域名正規化（去括號後綴），與 area_map 的 key 一致
                vs_data_norm = st.session_state['vs_data'].copy()
                vs_data_norm['area'] = vs_data_norm['area'].apply(_norm_vs_area)

                result = compare(
                    ay_data[ay_choice],
                    vs_data_norm,
                    ticket_map=st.session_state['ticket_map'],
                    ay_area_map=st.session_state['area_map'],
                )
                st.session_state['compare_result'] = result
            except Exception as e:
                st.error(f'比對失敗：{e}')

    if 'compare_result' in st.session_state:
        _show_results(st.session_state['compare_result'])


def _show_results(result: dict):
    st.subheader('⑥ 比對結果')
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
