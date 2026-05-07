import streamlit as st
from compare import parse_ansource, parse_vendor, compare, generate_excel

st.set_page_config(page_title='票價比對工具', page_icon='🎫', layout='wide')


def check_login(username: str, password: str) -> bool:
    return (username == st.secrets['auth']['username'] and
            password == st.secrets['auth']['password'])


def login_page():
    st.title('🎫 票價比對工具')
    st.subheader('請登入')
    with st.form('login_form'):
        username = st.text_input('帳號')
        password = st.text_input('密碼', type='password')
        submitted = st.form_submit_button('登入', type='primary')
    if submitted:
        if check_login(username, password):
            st.session_state['logged_in'] = True
            st.rerun()
        else:
            st.error('帳號或密碼錯誤，請重試')


def main_page():
    col_title, col_logout = st.columns([8, 1])
    with col_title:
        st.title('🎫 票價比對工具')
    with col_logout:
        st.write('')
        if st.button('登出'):
            st.session_state['logged_in'] = False
            st.rerun()

    st.markdown('---')
    st.subheader('① 上傳檔案')
    col1, col2 = st.columns(2)
    with col1:
        ay_file = st.file_uploader('安源 .xls', type=['xls'])
    with col2:
        vs_file = st.file_uploader('廠商 .xlsx', type=['xlsx'])

    if not (ay_file and vs_file):
        st.info('請上傳兩份 Excel 檔案後繼續')
        return

    ay_key = f'ay_{ay_file.name}'
    vs_key = f'vs_{vs_file.name}'
    if ay_key not in st.session_state or vs_key not in st.session_state:
        with st.spinner('讀取檔案中...'):
            try:
                st.session_state[ay_key] = parse_ansource(ay_file)
                st.session_state[vs_key] = parse_vendor(vs_file)
            except Exception as e:
                st.error(f'檔案讀取失敗：{e}')
                return
    ay_data = st.session_state[ay_key]
    vs_data = st.session_state[vs_key]

    if 'compare_result' in st.session_state:
        # Clear cached compare result if files have changed
        result_key = st.session_state.get('compare_files_key')
        current_key = f'{ay_file.name}_{vs_file.name}'
        if result_key != current_key:
            del st.session_state['compare_result']
    st.session_state['compare_files_key'] = f'{ay_file.name}_{vs_file.name}'

    st.markdown('---')
    st.subheader('② 選擇比對場次')
    col3, col4 = st.columns(2)
    with col3:
        ay_choice = st.selectbox('安源場次日期', list(ay_data.keys()))
    with col4:
        vs_choice = st.selectbox('廠商票價頁籤', list(vs_data.keys()))

    st.markdown('---')
    if st.button('🔍 開始比對', type='primary'):
        with st.spinner('比對中...'):
            st.session_state['compare_result'] = compare(ay_data[ay_choice], vs_data[vs_choice])
    if 'compare_result' in st.session_state:
        _show_results(st.session_state['compare_result'])


def _show_results(result: dict):
    st.subheader('③ 比對結果')

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric('✅ 相符', len(result['price_ok']))
    c2.metric('❌ 票價不符', len(result['price_diff']))
    c3.metric('⚠️ 廠商有安源無', len(result['only_vs']))
    c4.metric('⚠️ 安源有廠商無', len(result['only_ay']))
    c5.metric('ℹ️ 無對應票種', len(result['unmapped']))

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
            file_name='票價比對結果.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type='primary',
        )
    except Exception as e:
        st.error(f'Excel 報告產生失敗：{e}')


def main():
    if 'logged_in' not in st.session_state:
        st.session_state['logged_in'] = False
    if st.session_state['logged_in']:
        main_page()
    else:
        login_page()


if __name__ == '__main__':
    main()
