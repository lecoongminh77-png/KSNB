import sys
import os
import io
import time
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import traceback
import urllib.request

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN CHUẨN APPLE & THANH TIỆN ÍCH CỐ ĐỊNH
# -------------------------------------------------------------
st.set_page_config(
    page_title='Đối Chiếu Nghiệm Thu HRC - KSNB', page_icon='🛡️', layout='wide'
)

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
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff;
        border-radius: 16px;
        padding: 15px;
        border: 1px solid rgba(0, 0, 0, 0.08);
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.03);
    }
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

col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', use_container_width=True):
        st.switch_page('app_ksnb.py')

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)
st.markdown(
    """
<div class="apple-header-box">
    <div>
        <h1 class="apple-main-title">📑 Đối Chiếu Nghiệm Thu HRC</h1>
    </div>
</div>
""", unsafe_allow_html=True)
st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>", unsafe_allow_html=True)

# ==============================================================================
# HÀM BỔ TRỢ ÉP KIỂU SỐ AN TOÀN
# ==============================================================================
def to_float(v):
    if v is None: return 0.0
    v_str = str(v).strip().upper()
    if v_str in ['', 'NONE', '#N/A', 'NAN', 'N/A']: return 0.0
    try: return float(v)
    except: return 0.0

# ==============================================================================
# CHỨC NĂNG: ĐỐI CHIẾU NGHIỆM THU HRC LÕI PANDAS TẠO MỚI (CÓ LOGO & KIỂM TRA SUM)
# ==============================================================================
def doi_chieu_hrc_tung_file():
    st.caption('Kiểm soát nội bộ • Đọc Pandas tốc độ cao, tái tạo Form & kiểm toán dòng Tổng cộng')
    
    st.markdown("### 📥 Nạp Dữ Liệu Nguồn")
    
    col_upload1, col_upload2 = st.columns(2, gap="large")
    
    with col_upload1:
        with st.container(border=True):
            st.markdown("""
<div class="upload-title">1️⃣ Dữ liệu Nghiệm thu & Trạm cân</div>
<div class="guide-box">
<div class="guide-header">📌 1. Biên bản nghiệm thu (BBNT)</div>
</div>
""", unsafe_allow_html=True)
            
            f_bbnt = st.file_uploader('1. File Biên bản nghiệm thu (BBNT)', type=['xlsx', 'xls'])
            
            selected_sheets = []
            if f_bbnt is not None:
                try:
                    xls_bbnt = pd.ExcelFile(f_bbnt)
                    sheet_names = xls_bbnt.sheet_names
                    selected_sheets = st.multiselect(
                        '📌 Chọn các Lô / Sheet BBNT cần đối chiếu:',
                        options=sheet_names,
                        default=[sheet_names[0]] if sheet_names else []
                    )
                except Exception as e:
                    st.error(f"Lỗi đọc danh sách sheet: {e}")
                    
            st.markdown("""
<div class="guide-box">
<div class="guide-header">📌 2. Packing List nhà cung cấp & 3. Trạm Cân</div>
</div>
""", unsafe_allow_html=True)
            f_pl = st.file_uploader('2. File Packing List nhà cung cấp', type=['xlsx', 'xls'])
            f_tc = st.file_uploader('3. File Dữ liệu Trạm cân', type=['xlsx', 'xls'])

    with col_upload2:
        with st.container(border=True):
            st.markdown("""
<div class="upload-title">2️⃣ Dữ liệu Bảo vệ & ERP</div>
<div class="guide-box">
<div class="guide-header">📌 4. Dữ liệu Nhập Bảo vệ</div>
</div>
<div class="guide-box">
<div class="guide-header">📌 5. Báo cáo ERP OPM 105</div>
</div>
""", unsafe_allow_html=True)
            f_bv = st.file_uploader('4. File Dữ liệu Nhập Bảo vệ', type=['xlsx', 'xls'])
            f_opm = st.file_uploader('5. File Báo cáo ERP OPM 105', type=['xlsx', 'xls'])

    if f_bbnt is not None:
        if len(selected_sheets) == 0:
            st.warning("⚠️ Vui lòng chọn ít nhất 1 sheet BBNT để bắt đầu.")
        else:
            if st.button('🚀 Bắt đầu đối chiếu tự động', use_container_width=True):
                my_bar = st.progress(0, text="🔄 Bắt đầu nạp hệ thống...")
                try:
                    # =========================================================
                    # BƯỚC 1: XÂY DỰNG KHO DỮ LIỆU BỔ TRỢ (CHẠY 1 LẦN)
                    # =========================================================
                    my_bar.progress(10, text="📦 Đang xử lý Packing List...")
                    dict_pl_master = {}
                    if f_pl is not None:
                        xls_pl = pd.ExcelFile(f_pl)
                        pl_sheets = xls_pl.sheet_names
                        for s in selected_sheets:
                            if s in pl_sheets:
                                df_pl = pd.read_excel(f_pl, sheet_name=s, skiprows=6)
                                if len(df_pl.columns) > 11:
                                    col_pl_ma = df_pl.columns[7]
                                    col_pl_net = df_pl.columns[11]
                                    df_pl_clean = df_pl.dropna(subset=[col_pl_ma, col_pl_net]).copy()
                                    pl_key = df_pl_clean[col_pl_ma].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper() + 'A'
                                    pl_net_val = pd.to_numeric(df_pl_clean[col_pl_net], errors='coerce').fillna(0)
                                    dict_pl_master[s] = dict(zip(pl_key, pl_net_val))

                    my_bar.progress(25, text="⚖️ Đang xử lý Dữ liệu Trạm cân & Bảo vệ...")
                    dict_tc_kl, dict_tc_phanloai = {}, {}
                    if f_tc is not None:
                        df_tc = pd.read_excel(f_tc)
                        if 'PHIẾU CÂN' in df_tc.columns and 'KHỐI LƯỢNG PHIẾU CÂN' in df_tc.columns:
                            tc_p = df_tc['PHIẾU CÂN'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                            tc_kl = pd.to_numeric(df_tc['KHỐI LƯỢNG PHIẾU CÂN'], errors='coerce').fillna(0)
                            tc_pl = df_tc['PHÂN LOẠI'].fillna('Nhập hàng').astype(str).str.strip() if 'PHÂN LOẠI' in df_tc.columns else 'Nhập hàng'
                            dict_tc_kl = dict(zip(tc_p, tc_kl))
                            dict_tc_phanloai = dict(zip(tc_p, tc_pl)) if not isinstance(tc_pl, str) else {k: tc_pl for k in tc_p}

                    dict_bv_phieu, dict_bv_kl = {}, {}
                    if f_bv is not None:
                        df_bv = pd.read_excel(f_bv)
                        for _, r in df_bv.iterrows():
                            p_can = str(r.get('MÃ PHIẾU CÂN', '')).replace('.0', '').strip()
                            txt = f"{r.get('MÃ HÀNG HÓA', '')} {r.get('TÊN HÀNG HÓA', '')}".upper()
                            dict_bv_phieu[txt] = p_can 
                            dict_bv_kl[txt] = r.get('KHỐI LƯỢNG PHIẾU CÂN', 0)

                    my_bar.progress(40, text="🏭 Đang xử lý Báo cáo ERP OPM...")
                    opm_035_po, opm_035_rtv, opm_031, opm_03g = {}, {}, {}, {}
                    if f_opm is not None:
                        df_opm = pd.read_excel(f_opm, skiprows=8)
                        if len(df_opm.columns) > 35:
                            df_opm_clean = df_opm.dropna(subset=[df_opm.columns[0], df_opm.columns[1], df_opm.columns[10]]).copy()
                            df_opm_clean['Mã Lô'] = df_opm_clean[df_opm.columns[0]].astype(str).str.strip().str.upper()
                            df_opm_clean['Khối Lượng'] = pd.to_numeric(df_opm_clean[df_opm.columns[10]], errors='coerce').fillna(0)
                            df_opm_clean['Mã Kho'] = df_opm_clean[df_opm.columns[35]].astype(str).str.strip()
                            opm_035_po = df_opm_clean[(df_opm_clean['Mã Kho'] == '035') & (df_opm_clean[df_opm.columns[1]] == 'PO Receipt')].groupby('Mã Lô')['Khối Lượng'].sum().to_dict()
                            opm_035_rtv = df_opm_clean[(df_opm_clean['Mã Kho'] == '035') & (df_opm_clean[df_opm.columns[1]] == 'Return to Vendor')].groupby('Mã Lô')['Khối Lượng'].sum().to_dict()
                            opm_031 = df_opm_clean[(df_opm_clean['Mã Kho'].isin(['03G', '03H'])) & (df_opm_clean[df_opm.columns[1]] == 'Intransit Receipt')].groupby('Mã Lô')['Khối Lượng'].sum().to_dict()
                            opm_03g = df_opm_clean[(df_opm_clean['Mã Kho'] == '03G') & (df_opm_clean[df_opm.columns[1]] == 'PO Receipt')].groupby('Mã Lô')['Khối Lượng'].sum().to_dict()

                    # =========================================================
                    # BƯỚC 2: TẠO FILE EXCEL MỚI - CHẠY PANDAS - VẼ LẠI MẪU
                    # =========================================================
                    my_bar.progress(50, text="🧠 Đang tạo File Báo cáo và chạy Đối chiếu...")
                    output = io.BytesIO()
                    
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        total_sheets = len(selected_sheets)
                        
                        for idx, s in enumerate(selected_sheets):
                            current_progress = 50 + int(((idx + 1) / total_sheets) * 40)
                            current_progress = min(current_progress, 95)
                            my_bar.progress(current_progress, text=f"🔍 Đang đối chiếu Sheet: {s}...")
                            
                            # --- ĐỌC DỮ LIỆU ĐỂ KIỂM TOÁN TỔNG CỘNG ---
                            df_raw = pd.read_excel(f_bbnt, sheet_name=s, header=None)
                            
                            # Lưu giữ 7 dòng tiêu đề gốc để lát vẽ lại
                            header_data = df_raw.iloc[0:7].values.tolist()
                            
                            orig_sums = {'net':0, 'gross':0, 'canban':0, 'chenhlech':0, 'pct':0}
                            for r_idx, row_vals in df_raw.iterrows():
                                row_str = ' '.join([str(x).lower() for x in row_vals if pd.notna(x)])
                                if 'tổng cộng' in row_str:
                                    orig_sums['net'] = to_float(row_vals[7])
                                    orig_sums['gross'] = to_float(row_vals[8])
                                    orig_sums['canban'] = to_float(row_vals[9])
                                    orig_sums['chenhlech'] = to_float(row_vals[10])
                                    orig_sums['pct'] = to_float(row_vals[11])
                                    break
                                    
                            # --- ĐỌC DỮ LIỆU CỐT LÕI (Từ Code Cũ của bạn) ---
                            df_bbnt = pd.read_excel(f_bbnt, sheet_name=s, skiprows=7)
                            # Loại bỏ các dòng rỗng mã cuộn hoặc dòng chứa chữ Tổng Cộng ở mã cuộn
                            df_bbnt = df_bbnt.dropna(subset=[df_bbnt.columns[3]]).copy()
                            df_bbnt = df_bbnt[~df_bbnt[df_bbnt.columns[3]].astype(str).str.lower().str.contains('tổng cộng', na=False)]

                            base_cols = df_bbnt.columns[:13].tolist()
                            res_df = df_bbnt[base_cols].copy()

                            col_ma_cuon = df_bbnt.columns[3]
                            col_phieu_can = df_bbnt.columns[4]
                            col_net_ct = df_bbnt.columns[7]
                            col_gross_ct = df_bbnt.columns[8]
                            col_can_ban = df_bbnt.columns[9]
                            col_lech_pt_bbnt = df_bbnt.columns[11]

                            ma_cuon_clean = df_bbnt[col_ma_cuon].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
                            phieu_can_clean = df_bbnt[col_phieu_can].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                            ma_hoa = ma_cuon_clean + 'A'
                            
                            dict_pl = dict_pl_master.get(s, {})

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
                            res_df['Chênh lệch  '] = res_df.apply(lambda r: 0 if r['KL Nhập kho 031'] == 0 else r['Khối lượng net theo PL'] - r['KL Nhập kho 031'], axis=1)

                            res_df['KL Nhập kho 03G-3H'] = ma_hoa.map(lambda x: opm_03g.get(x, 0))
                            res_df['Chênh lệch   '] = res_df['Khối lượng net theo PL'] - res_df['KL Nhập kho 03G-3H']

                            # TÍNH TỔNG CỘNG THỰC TẾ
                            chenh_lech_kg_file = pd.to_numeric(res_df.iloc[:, 10], errors='coerce').fillna(0) # Col 10 là chênh lệch Kg (cũ)
                            actual_net = net_ct_val.sum()
                            actual_gross = gross_ct_val.sum()
                            actual_canban = can_ban_val.sum()
                            actual_chenhlech = chenh_lech_kg_file.sum()
                            actual_pct = (actual_chenhlech / actual_gross) if actual_gross != 0 else 0
                            
                            # Kiểm tra xem file gốc có tính sai không (Lệch > 0.1)
                            has_sum_error = False
                            if (abs(orig_sums['net'] - actual_net) > 0.1 or 
                                abs(orig_sums['gross'] - actual_gross) > 0.1 or 
                                abs(orig_sums['canban'] - actual_canban) > 0.1 or 
                                abs(orig_sums['chenhlech'] - actual_chenhlech) > 0.1 or 
                                abs(orig_sums['pct'] - actual_pct) > 0.0001):
                                has_sum_error = True

                            # --- XUẤT RA EXCEL TẠI DÒNG SỐ 8 ---
                            res_df.to_excel(writer, index=False, sheet_name=s, startrow=7)
                            ws = writer.sheets[s]

                            # 1. VẼ LẠI 7 DÒNG HEADER & TẢI LOGO HOA SEN
                            for r_i, row_data in enumerate(header_data):
                                for c_i, val in enumerate(row_data):
                                    if pd.notna(val) and str(val).strip() != '':
                                        cell = ws.cell(row=r_i+1, column=c_i+1)
                                        cell.value = str(val)
                                        if r_i == 1: # Dòng Tiêu đề bôi đậm
                                            cell.font = Font(bold=True, size=12)

                            # Tải và chèn Logo Hoa Sen (Lấy hình trực tiếp từ web)
                            try:
                                req = urllib.request.Request('https://hoasengroup.vn/wp-content/uploads/2023/05/Logo-HSG-01.png', headers={'User-Agent': 'Mozilla/5.0'})
                                img_data = urllib.request.urlopen(req).read()
                                img = openpyxl.drawing.image.Image(io.BytesIO(img_data))
                                img.width, img.height = 140, 45 # Căn chỉnh cho đẹp
                                ws.add_image(img, 'A1')
                            except Exception:
                                pass # Nếu không có mạng thì bỏ qua logo, form chữ vẫn còn nguyên

                            # 2. ĐỊNH DẠNG MÀU SẮC (Code Cũ)
                            yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
                            red_fill = PatternFill(start_color='FFFFC7CE', end_color='FFFFC7CE', fill_type='solid')
                            red_font = Font(color='9C0006', bold=True)
                            border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), 
                                                 top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

                            # Tô vàng Header mới
                            for c in range(14, 34):
                                cell = ws.cell(row=8, column=c) # Header hiện ở dòng 8
                                cell.fill = yellow_fill
                                cell.font = Font(bold=True)
                                cell.border = border_thin

                            # Rà soát lỗi và viền bảng
                            col_names = list(res_df.columns)
                            for r_idx in range(9, len(res_df) + 9): # Dữ liệu từ dòng 9
                                for c_idx in range(1, len(col_names) + 1):
                                    cell = ws.cell(row=r_idx, column=c_idx)
                                    c_name = col_names[c_idx - 1]
                                    val = cell.value

                                    cell.border = border_thin
                                    is_error = False

                                    if c_name in ['Đối chiếu số phiếu cân', 'Đối chiếu KL TC & KL BCBV', 'Đối chiếu KL TC & KL BBNT', 'Đối chiếu dung sai'] and val is False:
                                        is_error = True
                                    elif c_name == 'So với quy định ( >+-0.3%)' and str(val).strip() == 'VUOT DUNG SAI':
                                        is_error = True
                                    elif c_name in ['Đối chiếu', 'Chênh lệch ', 'Chênh lệch  ', 'Chênh lệch   '] and isinstance(val, (int, float)) and val != 0:
                                        is_error = True

                                    if is_error:
                                        cell.fill = red_fill
                                        cell.font = red_font
                                        
                            # 3. CHỐT HẠ: CHÈN DÒNG TỔNG CỘNG ĐÃ KIỂM TOÁN VÀO CUỐI BẢNG
                            last_row = 9 + len(res_df)
                            ws.cell(row=last_row, column=1, value='Tổng cộng').font = Font(bold=True)
                            ws.cell(row=last_row, column=8, value=actual_net).font = Font(bold=True)
                            ws.cell(row=last_row, column=9, value=actual_gross).font = Font(bold=True)
                            ws.cell(row=last_row, column=10, value=actual_canban).font = Font(bold=True)
                            ws.cell(row=last_row, column=11, value=actual_chenhlech).font = Font(bold=True)
                            ws.cell(row=last_row, column=12, value=actual_pct).font = Font(bold=True)
                            
                            # Nếu file gốc tính bậy, tô đỏ rực toàn bộ dòng tổng mới này!
                            for c_idx in range(1, 14):
                                c_cell = ws.cell(row=last_row, column=c_idx)
                                c_cell.border = border_thin
                                if has_sum_error:
                                    c_cell.fill = red_fill
                                    c_cell.font = red_font

                    # --- LƯU FILE KẾT QUẢ ---
                    my_bar.progress(100, text="✅ Đã tải Logo Hoa Sen và hoàn tất đóng gói!")
                    st.success(f'🎉 Kế toán đã rà soát xong {len(selected_sheets)} lô hàng (Sử dụng lõi Pandas siêu chuẩn)!')
                    st.download_button(
                        '📥 Tải File Biên Bản Đối Chiếu',
                        data=output.getvalue(),
                        file_name='BBNT_Nghiem_Thu_Doi_Chieu_Tong_Hop.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    )

                except Exception as e:
                    my_bar.empty()
                    st.error(f'🛑 Lỗi Cốt Lõi: {e}')
                    with st.expander("🔍 Chi tiết lỗi (Dành cho kỹ thuật)"):
                        st.code(traceback.format_exc())
    else:
        st.info('👉 Hãy nạp tối thiểu File BBNT để bắt đầu.')

doi_chieu_hrc_tung_file()
