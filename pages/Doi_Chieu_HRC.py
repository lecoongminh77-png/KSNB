import io
import re
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN CHUẨN APPLE & THANH TIỆN ÍCH CỐ ĐỊNH
# -------------------------------------------------------------
st.set_page_config(
    page_title='Đối Chiếu Nghiệm Thu HRC - KSNB',
    page_icon='🛡️',
    layout='wide',
)

# CSS ẩn menu sidebar và tùy chỉnh giao diện Apple
st.markdown(
    """
<style>
    /* Ẩn sidebar */
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }

    /* Nền trang */
    .stApp {
        background-color: #f5f5f7;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif;
        color: #1d1d1f;
    }
    
    /* Header */
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
    
    /* Nút bấm Apple Style */
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
    div.stButton > button:first-child:hover {
        background: #0077ed !important;
        color: #ffffff !important;
    }

    /* Container Streamlit Thẻ Apple */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff;
        border-radius: 16px;
        padding: 15px;
        border: 1px solid rgba(0, 0, 0, 0.08);
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.03);
    }

    /* Thành phần chữ trong thẻ */
    .upload-title {
        font-size: 18px;
        font-weight: 700;
        color: #1d1d1f;
        margin-bottom: 16px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .guide-box {
        background: #fbfbfd;
        border-radius: 12px;
        padding: 16px;
        margin-bottom: 16px;
        border: 1px solid rgba(0,0,0,0.04);
    }
    .guide-header {
        font-weight: 700;
        color: #0071e3;
        font-size: 14px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .guide-text {
        font-size: 13px;
        color: #515154;
        line-height: 1.6;
        margin: 0;
    }
</style>
""",
    unsafe_allow_html=True,
)

# Thanh điều hướng (Nút Trang chủ)
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
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    "<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>",
    unsafe_allow_html=True,
)


# -------------------------------------------------------------
# CÁC HÀM XỬ LÝ DỮ LIỆU TỰ ĐỘNG
# -------------------------------------------------------------
def read_auto_header(file_obj, keywords, sheet_name=0, max_scan=25, match_all=False):
    """Tự động tìm dòng header thực tế dựa trên danh sách từ khóa."""
    file_obj.seek(0)
    df_raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None, nrows=max_scan)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
        if match_all:
            if all(k.upper() in row_str for k in keywords):
                header_idx = idx
                break
        else:
            if any(k.upper() in row_str for k in keywords):
                header_idx = idx
                break
    file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name=sheet_name, skiprows=header_idx)
    return df


def doi_chieu_mot_sheet(f_bbnt, sheet_bbnt, f_pl, sheet_pl, dict_tc_kl, dict_tc_phanloai, dict_bv_phieu, dict_bv_kl, opm_035_po, opm_035_rtv, opm_031, opm_03g_po, opm_03g_rtv):
    # Chuẩn hóa mã để tránh mã cuộn/phiếu cân bị lệch do Excel lưu dạng số.
    def clean_value(value):
        if pd.isna(value):
            return ''
        return re.sub(r'\.0+$', '', str(value).strip())

    def lot_key(value):
        value = clean_value(value).upper()
        if not value:
            return ''
        return value if value.endswith('A') else value + 'A'

    # 1. ĐỌC BBNT
    df_bbnt = read_auto_header(
        f_bbnt,
        ['MÃ SỐ CUỘN', 'MÃ CUỘN', 'SỐ PHIẾU CÂN'],
        sheet_name=sheet_bbnt,
    )
    col_cuon = next(
        (c for c in df_bbnt.columns if 'CUỘN' in str(c).upper()),
        df_bbnt.columns[3],
    )
    col_pcan = next(
        (c for c in df_bbnt.columns if 'PHIẾU CÂN' in str(c).upper()),
        df_bbnt.columns[4],
    )
    col_net = next(
        (
            c for c in df_bbnt.columns
            if 'CHỨNG TỪ GIAO HÀNG' in str(c).upper()
            and 'GROSS' not in str(c).upper()
        ),
        df_bbnt.columns[7],
    )
    col_gross = next(
        (c for c in df_bbnt.columns if 'GROSS' in str(c).upper()),
        df_bbnt.columns[8],
    )
    col_canban = next(
        (c for c in df_bbnt.columns if 'CÂN BÀN' in str(c).upper()),
        df_bbnt.columns[9],
    )
    col_dsai = next(
        (
            c for c in df_bbnt.columns
            if 'CHÊNH LỆCH' in str(c).upper() and '%' in str(c)
        ),
        df_bbnt.columns[11],
    )

    df_bbnt = df_bbnt.dropna(subset=[col_cuon]).copy()
    df_bbnt = df_bbnt[
        df_bbnt[col_cuon].astype(str).str.strip().ne('')
    ].copy()
    df_bbnt = df_bbnt[
        ~df_bbnt[df_bbnt.columns[0]].astype(str).str.contains(
            'Tổng cộng|Kết luận',
            case=False,
            na=False,
        )
    ].copy()

    base_cols = df_bbnt.columns[:13].tolist()
    if len(base_cols) != 13:
        raise ValueError(
            f"Sheet '{sheet_bbnt}' cần có 13 cột BBNT gốc A:M, "
            f"nhưng hiện có {len(base_cols)} cột."
        )

    res_df = df_bbnt[base_cols].copy()

    ma_cuon = df_bbnt[col_cuon].map(
        lambda value: clean_value(value).upper()
    )
    phieu_can_bbnt = df_bbnt[col_pcan].map(clean_value)
    ma_hoa = ma_cuon.map(lot_key)

    # 2. ĐỌC PACKING LIST
    dict_pl = {}
    if f_pl is not None and sheet_pl is not None:
        try:
            df_pl = read_auto_header(
                f_pl,
                ['MÃ SỐ NSX', 'MÃ SỐ HSPM', 'KL(NET)'],
                sheet_name=sheet_pl,
            )
            c_pl_ma = next(
                (
                    c for c in df_pl.columns
                    if 'MÃ SỐ NSX' in str(c).upper()
                    or 'MÃ SỐ' in str(c).upper()
                ),
                None,
            )
            c_pl_kl = next(
                (c for c in df_pl.columns if 'NET' in str(c).upper()),
                None,
            )

            if c_pl_ma is not None and c_pl_kl is not None:
                df_pl_clean = df_pl.dropna(
                    subset=[c_pl_ma, c_pl_kl]
                ).copy()
                pl_key = df_pl_clean[c_pl_ma].map(lot_key)
                pl_val = pd.to_numeric(
                    df_pl_clean[c_pl_kl],
                    errors='coerce',
                ).fillna(0)
                dict_pl = (
                    pd.DataFrame({'Key': pl_key, 'Net': pl_val})
                    .groupby('Key')['Net']
                    .sum()
                    .to_dict()
                )
        except Exception:
            dict_pl = {}

    # 3. TÍNH CÁC CỘT ĐỐI SOÁT
    nhap_xuat = phieu_can_bbnt.map(
        lambda value: dict_tc_phanloai.get(value, 'Nhập hàng')
    )
    net_pl = ma_hoa.map(dict_pl).fillna(0).astype(float)
    net_ct = pd.to_numeric(
        df_bbnt[col_net],
        errors='coerce',
    ).fillna(0).astype(float)
    doi_chieu_pl = net_ct - net_pl

    bv_weight = pd.to_numeric(
        ma_cuon.map(dict_bv_kl),
        errors='coerce',
    ).fillna(0).astype(float)
    bv_ticket = ma_cuon.map(dict_bv_phieu).map(clean_value).fillna('')
    chk_pcan = bv_ticket.ne('') & phieu_can_bbnt.eq(bv_ticket)

    scale_weight = pd.to_numeric(
        bv_ticket.map(dict_tc_kl),
        errors='coerce',
    ).fillna(0).astype(float)
    chk_tc_bv = (
        bv_weight.gt(0)
        & (scale_weight.sub(bv_weight).abs() <= 0.000001)
    )

    can_ban = pd.to_numeric(
        df_bbnt[col_canban],
        errors='coerce',
    ).fillna(0).astype(float)
    chk_tc_bbnt = (
        can_ban.gt(0)
        & (scale_weight.sub(can_ban).abs() <= 0.000001)
    )

    gross_ct = pd.to_numeric(
        df_bbnt[col_gross],
        errors='coerce',
    ).fillna(0).astype(float)
    chenh_lech_kg = scale_weight - gross_ct
    gross_safe = gross_ct.where(gross_ct.ne(0))
    chenh_lech_ds = (
        chenh_lech_kg.div(gross_safe)
        .replace([float('inf'), float('-inf')], 0)
        .fillna(0)
    )

    dsai_bbnt = pd.to_numeric(
        df_bbnt[col_dsai],
        errors='coerce',
    ).fillna(0).astype(float)
    chk_dsai = chenh_lech_ds.round(6) == dsai_bbnt.round(6)
    so_quy_dinh = chenh_lech_ds.apply(
        lambda value: 'VUOT DUNG SAI' if abs(value) > 0.003 else 'OK'
    )

    kl_035 = ma_hoa.map(
        lambda key: opm_035_po.get(key, 0)
        - abs(opm_035_rtv.get(key, 0))
    )
    chenh_035 = net_pl - kl_035

    kl_031 = ma_hoa.map(lambda key: opm_031.get(key, 0))
    chenh_031 = (net_pl - kl_031).where(kl_031.ne(0), 0)

    kl_03g = ma_hoa.map(
        lambda key: opm_03g_po.get(key, 0)
        - abs(opm_03g_rtv.get(key, 0))
    )
    chenh_03g = net_pl - kl_03g

    # Cột V trong mẫu có tiêu đề trống nhưng vẫn chứa phiếu cân + hậu tố A.
    # Ghép theo vị trí để giữ cột "Số phiếu cân" gốc ở E
    # và cột "Số phiếu cân" kết quả ở S.
    result_columns = [
        ('NHẬP/XUẤT', nhap_xuat),
        ('Mã hóa', ma_hoa),
        ('Khối lượng net theo PL', net_pl),
        ('Đối chiếu', doi_chieu_pl),
        ('KL theo bc nhập hàng', bv_weight),
        ('Số phiếu cân', bv_ticket),
        ('Đối chiếu số phiếu cân', chk_pcan),
        ('KL theo bc trạm cân', scale_weight),
        ('', bv_ticket.map(lambda value: value + 'A' if value else '')),
        ('Đối chiếu KL TC & KL BCBV', chk_tc_bv),
        ('Đối chiếu KL TC & KL BBNT', chk_tc_bbnt),
        ('Chênh lệch (kg)', chenh_lech_kg),
        ('Chênh lệch dung sai', chenh_lech_ds),
        ('Đối chiếu dung sai', chk_dsai),
        ('So với quy định ( >+-0.3%)', so_quy_dinh),
        ('KL Nhập kho 035', kl_035),
        ('Chênh lệch ', chenh_035),
        ('KL Nhập kho 031', kl_031),
        ('Chênh lệch  ', chenh_031),
        ('KL Nhập kho 03G-3H', kl_03g),
        ('Chênh lệch   ', chenh_03g),
    ]

    result_df = pd.concat(
        [values.rename(name) for name, values in result_columns],
        axis=1,
    )
    res_df = pd.concat([res_df, result_df], axis=1)

    return res_df


# -------------------------------------------------------------
# CHỨC NĂNG CHÍNH
# -------------------------------------------------------------
def doi_chieu_hrc_tung_file():
    st.caption('Kiểm soát nội bộ • Tự động đối soát 4 bên và tô đỏ ô sai phạm dữ liệu')
    st.markdown("### 📥 Nạp Dữ Liệu Nguồn")

    col_upload1, col_upload2 = st.columns(2, gap="large")

    with col_upload1:
        with st.container(border=True):
            st.markdown("""
<div class="upload-title">1️⃣ Dữ liệu Nghiệm thu & Trạm cân</div>
<div class="guide-box">
<div class="guide-header">📌 1. Biên bản nghiệm thu (BBNT)</div>
<div class="guide-text">Đường dẫn: Nguyen Thi Thu Thuong > Biên bản nghiệm thu (BBNT)</div>
</div>
""", unsafe_allow_html=True)

            f_bbnt = st.file_uploader('1. File Biên bản nghiệm thu (BBNT)', type=['xlsx', 'xls'])

            # Ô CHỌN NHIỀU SHEET
            selected_sheets = []
            if f_bbnt is not None:
                try:
                    f_bbnt.seek(0)
                    xl_bb = pd.ExcelFile(f_bbnt)
                    sheet_names_bb = xl_bb.sheet_names
                    selected_sheets = st.multiselect(
                        '🎯 Chọn Sheet cần đối chiếu:',
                        options=sheet_names_bb,
                        default=[sheet_names_bb[0]] if sheet_names_bb else [],
                        help='Các sheet không được chọn sẽ tự động bị loại bỏ khỏi file kết quả'
                    )
                except Exception as e:
                    st.error(f"Lỗi đọc danh sách sheet: {e}")
            else:
                st.caption("ℹ️ Nạp file BBNT để hiện danh sách sheet.")

            st.markdown("""
<div class="guide-box" style="margin-top: 15px;">
<div class="guide-header">📌 2. Packing List nhà cung cấp</div>
<div class="guide-text">Đường dẫn: Nguyen Thi Thu Thuong > Packing List nhà cung cấp</div>
</div>
""", unsafe_allow_html=True)
            f_pl = st.file_uploader('2. File Packing List nhà cung cấp', type=['xlsx', 'xls'])

            st.markdown("""
<div class="guide-box">
<div class="guide-header">📌 3. Dữ liệu Trạm cân</div>
<div class="guide-text">
Truy cập hệ thống trạm cân nội bộ (thông tin đăng nhập lưu riêng)<br>
Quản lý truy cập > Tra cứu trạm cân<br>
Thời gian: 1 tháng rưỡi
</div>
</div>
""", unsafe_allow_html=True)
            f_tc = st.file_uploader('3. File Dữ liệu Trạm cân', type=['xlsx', 'xls'])

    with col_upload2:
        with st.container(border=True):
            st.markdown("""
<div class="upload-title">2️⃣ Dữ liệu Bảo vệ & ERP</div>
<div class="guide-box">
<div class="guide-header">📌 4. Dữ liệu Nhập Bảo vệ</div>
<div class="guide-text">
Quản lý truy cập > Tra cứu xuất nhập hàng > Chọn Nhập hàng<br>
Thời gian: Giống mốc Tra cứu trạm cân
</div>
</div>
<div class="guide-box">
<div class="guide-header">📌 5. Báo cáo ERP OPM 105</div>
<div class="guide-text">
Chọn: HSG_OPM_105_BCKiemSoatThongTinGDTheoLo_GD2<br>
Tích chọn kho: 035, 03G, 03H<br>
Nghiệp vụ: Inventory, Job or Schedule, Purchase order<br>
Thời gian: 1 tháng cần kiểm tra
</div>
</div>
""", unsafe_allow_html=True)
            f_bv = st.file_uploader('4. File Dữ liệu Nhập Bảo vệ', type=['xlsx', 'xls'])
            f_opm = st.file_uploader('5. File Báo cáo ERP OPM 105', type=['xlsx', 'xls'])

    # TIẾN HÀNH XỬ LÝ
    st.write("")
    if f_bbnt is not None:
        if not selected_sheets:
            st.warning("⚠️ Vui lòng chọn ít nhất 1 sheet ở khung bên trái để bắt đầu đối chiếu.")
        else:
            if st.button('🚀 Bắt đầu đối chiếu tự động', use_container_width=True):
                with st.spinner(f'Đang đối soát {len(selected_sheets)} sheet dữ liệu và định dạng màu...'):
                    try:
                        # 1. QUÉT DỮ LIỆU BẢO VỆ
                        bv_raw_list = []
                        if f_bv is not None:
                            try:
                                df_bv = read_auto_header(f_bv, ['MÃ PHIẾU CÂN', 'KHỐI LƯỢNG PHIẾU CÂN'])
                                col_bv_hang = next((c for c in df_bv.columns if 'TÊN HÀNG HÓA' in str(c).upper() or 'MÃ HÀNG HÓA' in str(c).upper()), None)
                                col_bv_pcan = next((c for c in df_bv.columns if 'MÃ PHIẾU CÂN' in str(c).upper() or 'PHIẾU CÂN' in str(c).upper()), None)
                                col_bv_kl = next((c for c in df_bv.columns if 'KHỐI LƯỢNG' in str(c).upper()), None)

                                if col_bv_hang and col_bv_pcan and col_bv_kl:
                                    for _, r in df_bv.iterrows():
                                        txt = f"{r.get('MÃ HÀNG HÓA', '')} {r.get('TÊN HÀNG HÓA', '')}"
                                        p = str(r.get(col_bv_pcan, '')).replace('.0', '').strip()
                                        kl = r.get(col_bv_kl, 0)
                                        bv_raw_list.append((txt, p, kl))
                            except Exception:
                                bv_raw_list = []

                        # 2. QUÉT DỮ LIỆU TRẠM CÂN
                        dict_tc_kl, dict_tc_phanloai = {}, {}
                        if f_tc is not None:
                            try:
                                df_tc = read_auto_header(f_tc, ['PHIẾU CÂN', 'KHỐI LƯỢNG PHIẾU CÂN'])
                                col_tc_p = next((c for c in df_tc.columns if str(c).strip().upper() == 'PHIẾU CÂN'), df_tc.columns[0])
                                col_tc_kl = next((c for c in df_tc.columns if 'KHỐI LƯỢNG' in str(c).upper()), None)
                                col_tc_pl = next((c for c in df_tc.columns if 'PHÂN LOẠI' in str(c).upper()), None)

                                tc_p = df_tc[col_tc_p].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                                tc_kl = pd.to_numeric(df_tc[col_tc_kl], errors='coerce').fillna(0) if col_tc_kl else 0
                                dict_tc_kl = dict(zip(tc_p, tc_kl))
                                if col_tc_pl:
                                    dict_tc_phanloai = dict(zip(tc_p, df_tc[col_tc_pl].fillna('Nhập hàng').astype(str)))
                            except Exception:
                                dict_tc_kl, dict_tc_phanloai = {}, {}

                        # 3. QUÉT DỮ LIỆU ERP OPM 105 (Đồng nhất tên biến)
                        opm_035_po, opm_035_rtv, opm_031, opm_03g_po, opm_03g_rtv = {}, {}, {}, {}, {}
                        if f_opm is not None:
                            try:
                                df_opm = read_auto_header(f_opm, ['LOẠI GIAO DỊCH', 'MÃ LÔ'], match_all=True)
                                if len(df_opm) > 0 and '[' in str(df_opm.iloc[0, 1]):
                                    df_opm = df_opm.iloc[1:].copy()

                                c_opm_loai = next((c for c in df_opm.columns if 'LOẠI GIAO DỊCH' in str(c).upper()), None)
                                c_opm_malo = next((c for c in df_opm.columns if 'MÃ LÔ' in str(c).upper()), None)
                                c_opm_kl = next((c for c in df_opm.columns if 'KHỐI LƯỢNG' in str(c).upper() and 'GROSS' not in str(c).upper()), None)
                                c_opm_kho = next((c for c in df_opm.columns if any(k in str(c).upper() for k in ['MĂ KHO', 'MÃ KHO', 'MA KHO', 'KHO']) and 'KHU VỰC' not in str(c).upper()), None)

                                if c_opm_loai and c_opm_malo and c_opm_kl and c_opm_kho:
                                    df_opm_clean = df_opm.dropna(subset=[c_opm_malo, c_opm_loai]).copy()
                                    lo_str = df_opm_clean[c_opm_malo].astype(str).str.strip().str.upper()
                                    df_opm_clean['Key'] = lo_str.apply(lambda x: x if x.endswith('A') else x + 'A')
                                    df_opm_clean['KL'] = pd.to_numeric(df_opm_clean[c_opm_kl], errors='coerce').fillna(0)
                                    df_opm_clean['Kho_Clean'] = df_opm_clean[c_opm_kho].astype(str).str.strip().str.upper()

                                    opm_035_po = df_opm_clean[(df_opm_clean['Kho_Clean'] == '035') & (df_opm_clean[c_opm_loai].str.contains('PO Receipt', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
                                    opm_035_rtv = df_opm_clean[(df_opm_clean['Kho_Clean'] == '035') & (df_opm_clean[c_opm_loai].str.contains('Return to Vendor', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
                                    opm_031 = df_opm_clean[(df_opm_clean['Kho_Clean'].isin(['03G', '03H'])) & (df_opm_clean[c_opm_loai].str.contains('Intransit Receipt', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
                                    opm_03g_po = df_opm_clean[(df_opm_clean['Kho_Clean'] == '03G') & (df_opm_clean[c_opm_loai].str.contains('PO Receipt', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
                                    opm_03g_rtv = df_opm_clean[(df_opm_clean['Kho_Clean'] == '03G') & (df_opm_clean[c_opm_loai].str.contains('Return to Vendor', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
                            except Exception:
                                opm_035_po, opm_035_rtv, opm_031, opm_03g_po, opm_03g_rtv = {}, {}, {}, {}, {}

                        # 4. DANH SÁCH SHEET CỦA PACKING LIST ĐỂ BẮT CẶP
                        pl_sheet_names = []
                        if f_pl is not None:
                            try:
                                f_pl.seek(0)
                                pl_sheet_names = pd.ExcelFile(f_pl).sheet_names
                            except Exception:
                                pl_sheet_names = []

                        # 5. TIẾN HÀNH XỬ LÝ VÀ GHI TỪNG SHEET VÀO EXCEL
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            for s_name in selected_sheets:
                                matched_pl = next((p for p in pl_sheet_names if p.strip().upper() == s_name.strip().upper()), None)
                                
                                f_bbnt.seek(0)
                                df_temp = read_auto_header(f_bbnt, ['MÃ SỐ CUỘN', 'MÃ CUỘN'], sheet_name=s_name)
                                col_c_temp = next((c for c in df_temp.columns if 'CUỘN' in str(c).upper()), df_temp.columns[3])
                                cuon_list = df_temp[col_c_temp].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper().tolist()

                                dict_bv_p_sheet = {}
                                dict_bv_kl_sheet = {}
                                if bv_raw_list:
                                    for txt, p, kl in bv_raw_list:
                                        for m in cuon_list:
                                            if m in txt and m not in dict_bv_p_sheet:
                                                dict_bv_p_sheet[m] = p
                                                dict_bv_kl_sheet[m] = kl

                                res_df = doi_chieu_mot_sheet(
                                    f_bbnt, s_name, f_pl, matched_pl,
                                    dict_tc_kl, dict_tc_phanloai,
                                    dict_bv_p_sheet, dict_bv_kl_sheet,
                                    opm_035_po, opm_035_rtv, opm_031, opm_03g_po, opm_03g_rtv
                                )

                                clean_sname = re.sub(r'[\\/*?:\[\]]', '_', s_name)[:31]
                                res_df.to_excel(writer, index=False, sheet_name=clean_sname)
                                ws = writer.sheets[clean_sname]

                                # Định dạng màu
                                yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
                                red_fill = PatternFill(start_color='FFFFC7CE', end_color='FFFFC7CE', fill_type='solid')
                                red_font = Font(color='9C0006', bold=True)
                                border_thin = Border(
                                    left=Side(style='thin', color='D9D9D9'),
                                    right=Side(style='thin', color='D9D9D9'),
                                    top=Side(style='thin', color='D9D9D9'),
                                    bottom=Side(style='thin', color='D9D9D9'),
                                )

                                for c in range(14, len(res_df.columns) + 1):
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

                        st.success(f'✅ Đã đối soát thành công {len(selected_sheets)} sheet đã chọn!')
                        st.download_button(
                            f'📥 Tải File Kết Quả Nghiệm Thu ({len(selected_sheets)} Sheet)',
                            data=output.getvalue(),
                            file_name='BBNT_Nghiem_Thu_Da_Doi_Chieu.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f'Lỗi xử lý: {e}')
    else:
        st.info('👉 Hãy nạp tối thiểu File BBNT để bắt đầu đối chiếu.')


doi_chieu_hrc_tung_file()
