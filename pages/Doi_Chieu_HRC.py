import sys
import os

# 2. Khai báo các thư viện khác
import io
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st


# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN CHUẨN APPLE & THANH TIỆN ÍCH CỐ ĐỊNH
# -------------------------------------------------------------
st.set_page_config(
    page_title='Đối Chiếu Nghiệm Thu HRC - KSNB', page_icon='🛡️', layout='wide'
)

# CSS ẩn menu sidebar và tùy chỉnh giao diện Apple
st.markdown(
    """
<style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }

    .stApp {
        background-color: #f5f5f7;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif;
        color: #1d1d1f;
    }
    .apple-header-box {
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
        padding: 20px 0 10px 0;
        margin-bottom: 5px;
    }
    .apple-main-title {
        font-size: 40px;
        font-weight: 700;
        color: #1d1d1f;
        letter-spacing: -0.025em;
        margin: 0;
    }
    .apple-sub-text {
        font-size: 17px;
        font-weight: 600;
        color: #1d1d1f;
        text-align: right;
        margin: 0;
        line-height: 1.35;
    }
    .apple-sub-text span {
        color: #0071e3;
        font-size: 14px;
        display: block;
        margin-top: 2px;
        font-weight: 500;
    }
    div.stButton > button:first-child {
        background: #0071e3 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding: 10px 20px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2) !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

# Thanh điều hướng
# Thanh điều hướng (Chỉ giữ nút Trang chủ)
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', use_container_width=True):
        st.switch_page('app_ksnb.py')



st.markdown(
    "<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>",
    unsafe_allow_html=True,
)

# Header Trang
st.markdown(
    """
<div class="apple-header-box">
    <div>
        <h1 class="apple-main-title">📑 Đối Chiếu Nghiệm Thu HRC</h1>
    </div>
    <div>
        <div class="apple-sub-text">
            Nền tảng công cụ kiểm soát<br>nghiệm thu thông minh.
            <span>Lê Công Minh Workspace ›</span>
        </div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    "<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>",
    unsafe_allow_html=True,
)


# ==============================================================================
# CHỨC NĂNG: ĐỐI CHIẾU NGHIỆM THU HRC
# ==============================================================================
def doi_chieu_hrc_tung_file():
    st.caption('Kiểm soát nội bộ • Tự động đối soát 4 bên và tô đỏ ô sai phạm dữ liệu')

    col1, col2 = st.columns(2)
    with col1:
        f_bbnt = st.file_uploader('1. File Biên bản nghiệm thu (BBNT)', type=['xlsx', 'xls'])
        f_pl = st.file_uploader('2. File Packing List nhà cung cấp', type=['xlsx', 'xls'])
        f_tc = st.file_uploader('3. File Dữ liệu Trạm cân', type=['xlsx', 'xls'])
    with col2:
        f_bv = st.file_uploader('4. File Dữ liệu Nhập Bảo vệ', type=['xlsx', 'xls'])
        f_opm = st.file_uploader('5. File Báo cáo ERP OPM 105', type=['xlsx', 'xls'])

    if f_bbnt is not None:
        if st.button('🚀 Bắt đầu đối chiếu tự động', use_container_width=True):
            with st.spinner('Đang đối soát dữ liệu và định dạng màu...'):
                try:
                    df_bbnt = pd.read_excel(f_bbnt, skiprows=7)
                    df_bbnt = df_bbnt.dropna(subset=[df_bbnt.columns[3]]).copy()

                    base_cols = df_bbnt.columns[:13].tolist()
                    res_df = df_bbnt[base_cols].copy()

                    col_ma_cuon = df_bbnt.columns[3]
                    col_phieu_can = df_bbnt.columns[4]
                    col_net_ct = df_bbnt.columns[7]
                    col_gross_ct = df_bbnt.columns[8]
                    col_can_ban = df_bbnt.columns[9]
                    col_lech_pt_bbnt = df_bbnt.columns[11]

                    ma_cuon_clean = (
                        df_bbnt[col_ma_cuon]
                        .astype(str)
                        .str.replace(r'\.0$', '', regex=True)
                        .str.strip()
                        .str.upper()
                    )
                    phieu_can_clean = (
                        df_bbnt[col_phieu_can]
                        .astype(str)
                        .str.replace(r'\.0$', '', regex=True)
                        .str.strip()
                    )
                    ma_hoa = ma_cuon_clean + 'A'

                    dict_pl = {}
                    if f_pl is not None:
                        df_pl = pd.read_excel(f_pl, skiprows=6)
                        col_pl_ma = df_pl.columns[7]
                        col_pl_net = df_pl.columns[11]
                        df_pl_clean = df_pl.dropna(subset=[col_pl_ma, col_pl_net]).copy()
                        pl_key = (
                            df_pl_clean[col_pl_ma]
                            .astype(str)
                            .str.replace(r'\.0$', '', regex=True)
                            .str.strip()
                            .str.upper()
                            + 'A'
                        )
                        pl_net_val = pd.to_numeric(df_pl_clean[col_pl_net], errors='coerce').fillna(0)
                        dict_pl = (
                            pd.DataFrame({'Key': pl_key, 'Net': pl_net_val})
                            .groupby('Key')['Net']
                            .sum()
                            .to_dict()
                        )

                    dict_bv_phieu = {}
                    dict_bv_kl = {}
                    if f_bv is not None:
                        df_bv = pd.read_excel(f_bv)
                        for _, r in df_bv.iterrows():
                            txt = f"{r.get('MÃ HÀNG HÓA', '')} {r.get('TÊN HÀNG HÓA', '')}".upper()
                            p_can = str(r.get('MÃ PHIẾU CÂN', '')).replace('.0', '').strip()
                            kl_can = r.get('KHỐI LƯỢNG PHIẾU CÂN', 0)
                            for m in ma_cuon_clean:
                                if m in txt and m not in dict_bv_phieu:
                                    dict_bv_phieu[m] = p_can
                                    dict_bv_kl[m] = kl_can

                    dict_tc_kl = {}
                    dict_tc_phanloai = {}
                    if f_tc is not None:
                        df_tc = pd.read_excel(f_tc)
                        tc_p = df_tc['PHIẾU CÂN'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        tc_kl = pd.to_numeric(df_tc['KHỐI LƯỢNG PHIẾU CÂN'], errors='coerce').fillna(0)
                        tc_pl = (
                            df_tc['PHÂN LOẠI'].fillna('Nhập hàng').astype(str).str.strip()
                            if 'PHÂN LOẠI' in df_tc.columns
                            else 'Nhập hàng'
                        )
                        dict_tc_kl = dict(zip(tc_p, tc_kl))
                        if isinstance(tc_pl, str):
                            dict_tc_phanloai = {k: tc_pl for k in tc_p}
                        else:
                            dict_tc_phanloai = dict(zip(tc_p, tc_pl))

                    opm_035_po, opm_035_rtv, opm_031, opm_03g = {}, {}, {}, {}
                    if f_opm is not None:
                        df_opm = pd.read_excel(f_opm, skiprows=8)
                        col_opm_ma = df_opm.columns[0]
                        col_opm_loai = df_opm.columns[1]
                        col_opm_kl = df_opm.columns[10]
                        col_opm_kho = df_opm.columns[35]

                        df_opm_clean = df_opm.dropna(subset=[col_opm_ma, col_opm_loai, col_opm_kl]).copy()
                        df_opm_clean['Mã Lô'] = df_opm_clean[col_opm_ma].astype(str).str.strip().str.upper()
                        df_opm_clean['Khối Lượng'] = pd.to_numeric(df_opm_clean[col_opm_kl], errors='coerce').fillna(0)
                        df_opm_clean['Mã Kho'] = df_opm_clean[col_opm_kho].astype(str).str.strip()

                        opm_035_po = (
                            df_opm_clean[(df_opm_clean['Mã Kho'] == '035') & (df_opm_clean[col_opm_loai] == 'PO Receipt')]
                            .groupby('Mã Lô')['Khối Lượng']
                            .sum()
                            .to_dict()
                        )
                        opm_035_rtv = (
                            df_opm_clean[(df_opm_clean['Mã Kho'] == '035') & (df_opm_clean[col_opm_loai] == 'Return to Vendor')]
                            .groupby('Mã Lô')['Khối Lượng']
                            .sum()
                            .to_dict()
                        )
                        opm_031 = (
                            df_opm_clean[(df_opm_clean['Mã Kho'].isin(['03G', '03H'])) & (df_opm_clean[col_opm_loai] == 'Intransit Receipt')]
                            .groupby('Mã Lô')['Khối Lượng']
                            .sum()
                            .to_dict()
                        )
                        opm_03g = (
                            df_opm_clean[(df_opm_clean['Mã Kho'] == '03G') & (df_opm_clean[col_opm_loai] == 'PO Receipt')]
                            .groupby('Mã Lô')['Khối Lượng']
                            .sum()
                            .to_dict()
                        )

                    res_df['NHẬP/XUẤT'] = phieu_can_clean.map(lambda x: dict_tc_phanloai.get(x, 'Nhập hàng'))
                    res_df['Mã hóa'] = ma_hoa
                    res_df['Khối lượng net theo PL'] = ma_hoa.map(dict_pl).fillna(0)

                    net_ct_val = pd.to_numeric(df_bbnt[col_net_ct], errors='coerce').fillna(0)
                    res_df['Đối chiếu'] = (net_ct_val - res_df['Khối lượng net theo PL']).astype(float)

                    res_df['KL theo bc nhập hàng'] = ma_cuon_clean.map(dict_bv_kl).fillna(0)
                    res_df['Số phiếu cân'] = ma_cuon_clean.map(dict_bv_phieu).fillna('')
                    res_df['Đối chiếu số phiếu cân'] = phieu_can_clean == res_df['Số phiếu cân'].astype(str)

                    res_df['KL theo bc trạm cân'] = phieu_can_clean.map(dict_tc_kl).fillna(0).astype(float)
                    res_df['Đối chiếu KL TC & KL BCBV'] = res_df['KL theo bc trạm cân'] == pd.to_numeric(res_df['KL theo bc nhập hàng'], errors='coerce')

                    can_ban_val = pd.to_numeric(df_bbnt[col_can_ban], errors='coerce').fillna(0).astype(float)
                    res_df['Đối chiếu KL TC & KL BBNT'] = res_df['KL theo bc trạm cân'] == can_ban_val

                    gross_ct_val = pd.to_numeric(df_bbnt[col_gross_ct], errors='coerce').replace(0, pd.NA).fillna(1).astype(float)
                    res_df['Chênh lệch (kg)'] = res_df['KL theo bc trạm cân'] - gross_ct_val
                    res_df['Chênh lệch dung sai'] = res_df['Chênh lệch (kg)'] / gross_ct_val

                    lech_pt_val = pd.to_numeric(df_bbnt[col_lech_pt_bbnt], errors='coerce').fillna(0)
                    res_df['Đối chiếu dung sai'] = res_df['Chênh lệch dung sai'].round(6) == lech_pt_val.round(6)
                    res_df['So với quy định ( >+-0.3%)'] = res_df['Chênh lệch dung sai'].apply(lambda x: 'VUOT DUNG SAI' if abs(x) > 0.003 else 'OK')

                    res_df['KL Nhập kho 035'] = ma_hoa.map(lambda x: opm_035_po.get(x, 0) - abs(opm_035_rtv.get(x, 0)))
                    res_df['Chênh lệch '] = res_df['Khối lượng net theo PL'] - res_df['KL Nhập kho 035']

                    res_df['KL Nhập kho 031'] = ma_hoa.map(lambda x: opm_031.get(x, 0))
                    res_df['Chênh lệch  '] = res_df.apply(
                        lambda r: 0 if r['KL Nhập kho 031'] == 0 else r['Khối lượng net theo PL'] - r['KL Nhập kho 031'],
                        axis=1,
                    )

                    res_df['KL Nhập kho 03G-3H'] = ma_hoa.map(lambda x: opm_03g.get(x, 0))
                    res_df['Chênh lệch   '] = res_df['Khối lượng net theo PL'] - res_df['KL Nhập kho 03G-3H']

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        res_df.to_excel(writer, index=False, sheet_name='BBNT_Doi_Chieu')
                        ws = writer.sheets['BBNT_Doi_Chieu']

                        yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
                        red_fill = PatternFill(start_color='FFFFC7CE', end_color='FFFFC7CE', fill_type='solid')
                        red_font = Font(color='9C0006', bold=True)
                        border_thin = Border(
                            left=Side(style='thin', color='D9D9D9'),
                            right=Side(style='thin', color='D9D9D9'),
                            top=Side(style='thin', color='D9D9D9'),
                            bottom=Side(style='thin', color='D9D9D9'),
                        )

                        for c in range(14, 34):
                            cell = ws.cell(row=1, column=c)
                            cell.fill = yellow_fill
                            cell.font = Font(bold=True)

                        col_names = list(res_df.columns)
                        for r_idx in range(2, len(res_df) + 2):
                            for c_idx in range(1, len(col_names) + 1):
                                cell = ws.cell(row=r_idx, column=c_idx)
                                c_name = col_names[c_idx - 1]
                                val = cell.value

                                cell.border = border_thin
                                is_error = False

                                if c_name in [
                                    'Đối chiếu số phiếu cân',
                                    'Đối chiếu KL TC & KL BCBV',
                                    'Đối chiếu KL TC & KL BBNT',
                                    'Đối chiếu dung sai',
                                ] and val is False:
                                    is_error = True
                                elif c_name == 'So với quy định ( >+-0.3%)' and str(val).strip() == 'VUOT DUNG SAI':
                                    is_error = True
                                elif c_name in ['Đối chiếu', 'Chênh lệch '] and isinstance(val, (int, float)) and val != 0:
                                    is_error = True

                                if is_error:
                                    cell.fill = red_fill
                                    cell.font = red_font

                    st.success('✅ Đã đối soát xong!')
                    st.dataframe(res_df, use_container_width=True)

                    st.download_button(
                        '📥 Tải File Kết Quả Nghiệm Thu (Đã Tô Đỏ Lỗi)',
                        data=output.getvalue(),
                        file_name='BBNT_Nghiem_Thu_Da_Doi_Chieu.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    )

                except Exception as e:
                    st.error(f'Lỗi xử lý: {e}')
    else:
        st.info('👉 Hãy nạp tối thiểu File BBNT để bắt đầu đối chiếu.')

