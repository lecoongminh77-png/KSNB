import sys
import os
import io
import time
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side

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
# LOGIC KSNB TỔNG HỢP: ĐỌC 2 LỚP, BẢO LƯU TEMPLATE & THANH TIẾN TRÌNH %
# ==============================================================================
def doi_chieu_hrc_tung_file():
    st.caption('Kiểm soát nội bộ • Đọc thông minh nhiều sheet, đối soát 4 bên & phát hiện gian lận công thức')
    
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
                        default=[sheet_names[0]],
                        help="Có thể chọn cùng lúc nhiều nhà cung cấp."
                    )
                except Exception as e:
                    st.error(f"Lỗi đọc danh sách sheet BBNT: {e}")
                    
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

    # --- HÀM ÉP KIỂU SỐ AN TOÀN ---
    def to_float(v):
        if v is None: return 0.0
        v_str = str(v).strip().upper()
        if v_str in ['', 'NONE', '#N/A', 'NAN']: return 0.0
        try: return float(v)
        except: return 0.0

    # --- NÚT XỬ LÝ ---
    if f_bbnt is not None:
        if len(selected_sheets) == 0:
            st.warning("⚠️ Vui lòng chọn ít nhất 1 sheet BBNT để bắt đầu đối chiếu.")
        else:
            if st.button('🚀 Bắt đầu đối chiếu tự động', use_container_width=True):
                # Khởi tạo thanh Progress Bar
                progress_text = "🔄 Bắt đầu nạp hệ thống..."
                my_bar = st.progress(0, text=progress_text)
                
                try:
                    # --- BƯỚC 1: XÂY DỰNG KHO DỮ LIỆU ---
                    my_bar.progress(10, text="📦 Đang xử lý Packing List...")
                    dict_pl_master = {}
                    if f_pl is not None:
                        xls_pl = pd.ExcelFile(f_pl)
                        pl_sheets = xls_pl.sheet_names
                        for s in selected_sheets:
                            if s in pl_sheets:
                                df_pl = pd.read_excel(f_pl, sheet_name=s, skiprows=6)
                                col_pl_ma = df_pl.columns[7]
                                col_pl_net = df_pl.columns[11]
                                df_pl_clean = df_pl.dropna(subset=[col_pl_ma, col_pl_net]).copy()
                                pl_key = df_pl_clean[col_pl_ma].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper() + 'A'
                                pl_net_val = pd.to_numeric(df_pl_clean[col_pl_net], errors='coerce').fillna(0)
                                dict_pl_master[s] = dict(zip(pl_key, pl_net_val))
                            else:
                                st.warning(f"⚠️ Bỏ qua PL cho lô '{s}' vì không tìm thấy sheet cùng tên trong file Packing List.")

                    my_bar.progress(25, text="⚖️ Đang xử lý Dữ liệu Trạm cân & Bảo vệ...")
                    dict_tc_kl, dict_tc_phanloai = {}, {}
                    if f_tc is not None:
                        df_tc = pd.read_excel(f_tc)
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
                        df_opm_clean = df_opm.dropna(subset=[df_opm.columns[0], df_opm.columns[1], df_opm.columns[10]]).copy()
                        df_opm_clean['Mã Lô'] = df_opm_clean[df_opm.columns[0]].astype(str).str.strip().str.upper()
                        df_opm_clean['Khối Lượng'] = pd.to_numeric(df_opm_clean[df_opm.columns[10]], errors='coerce').fillna(0)
                        df_opm_clean['Mã Kho'] = df_opm_clean[df_opm.columns[35]].astype(str).str.strip()
                        opm_035_po = df_opm_clean[(df_opm_clean['Mã Kho'] == '035') & (df_opm_clean[df_opm.columns[1]] == 'PO Receipt')].groupby('Mã Lô')['Khối Lượng'].sum().to_dict()
                        opm_035_rtv = df_opm_clean[(df_opm_clean['Mã Kho'] == '035') & (df_opm_clean[df_opm.columns[1]] == 'Return to Vendor')].groupby('Mã Lô')['Khối Lượng'].sum().to_dict()
                        opm_031 = df_opm_clean[(df_opm_clean['Mã Kho'].isin(['03G', '03H'])) & (df_opm_clean[df_opm.columns[1]] == 'Intransit Receipt')].groupby('Mã Lô')['Khối Lượng'].sum().to_dict()
                        opm_03g = df_opm_clean[(df_opm_clean['Mã Kho'] == '03G') & (df_opm_clean[df_opm.columns[1]] == 'PO Receipt')].groupby('Mã Lô')['Khối Lượng'].sum().to_dict()

                    # --- BƯỚC 2: MỞ FILE 2 LỚP & ĐỐI CHIẾU ---
                    my_bar.progress(50, text="🧠 Đang phân tích kép biểu mẫu BBNT (Values & Formulas)...")
                    
                    # Bản 1: Đọc giá trị để bắt số thật (tránh lỗi VLOOKUP)
                    wb_values = openpyxl.load_workbook(f_bbnt, data_only=True)
                    # Bản 2: Đọc công thức để tô màu, viết kết quả và check dòng SUM
                    wb_formulas = openpyxl.load_workbook(f_bbnt, data_only=False) 
                    
                    yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
                    red_fill = PatternFill(start_color='FFFFC7CE', end_color='FFFFC7CE', fill_type='solid')
                    red_font = Font(color='9C0006', bold=True)
                    border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

                    new_cols = [
                        'NHẬP/XUẤT', 'Mã hóa', 'Khối lượng net theo PL', 'Đối chiếu', 
                        'KL theo bc nhập hàng', 'Số phiếu cân', 'Đối chiếu số phiếu cân', 
                        'KL theo bc trạm cân', 'Đối chiếu KL TC & KL BCBV', 'Đối chiếu KL TC & KL BBNT', 
                        'Chênh lệch (kg)', 'Chênh lệch dung sai', 'Đối chiếu dung sai', 
                        'So với quy định ( >+-0.3%)', 'KL Nhập kho 035', 'Chênh lệch ', 
                        'KL Nhập kho 031', 'Chênh lệch  ', 'KL Nhập kho 03G-3H', 'Chênh lệch   '
                    ]

                    total_sheets = len(selected_sheets)
                    for idx, sheet_name in enumerate(wb_formulas.sheetnames):
                        if sheet_name not in selected_sheets:
                            wb_formulas.remove(wb_formulas[sheet_name]) # Xóa sheet thừa ở bản output
                            continue
                            
                        # Cập nhật thanh % cho từng sheet
                        current_progress = 50 + int((idx / len(wb_formulas.sheetnames)) * 40)
                        my_bar.progress(current_progress, text=f"🔍 Đang đối chiếu Sheet: {sheet_name}...")

                        ws_val = wb_values[sheet_name]
                        ws_form = wb_formulas[sheet_name]
                        
                        # Tìm tọa độ dòng "Tổng cộng" bằng bản Values
                        tong_row = -1
                        for row in range(8, ws_val.max_row + 2):
                            for col in range(1, 5):
                                val = ws_val.cell(row=row, column=col).value
                                if val and isinstance(val, str) and 'tổng cộng' in val.lower():
                                    tong_row = row
                                    break
                            if tong_row != -1: break
                        
                        if tong_row == -1: tong_row = ws_val.max_row + 1

                        # Tô vàng Header Cột đối chiếu 
                        for i, col_name in enumerate(new_cols):
                            cell = ws_form.cell(row=8, column=14+i)
                            cell.value = col_name
                            cell.fill = yellow_fill
                            cell.font = Font(bold=True)
                            cell.border = border_thin

                        sum_net, sum_gross, sum_can_ban, sum_chenh_lech = 0.0, 0.0, 0.0, 0.0
                        
                        # Xử lý từng dòng dữ liệu bằng Data Thực (ws_val)
                        for r in range(10, tong_row):
                            raw_ma_cuon = ws_val.cell(r, 4).value
                            if raw_ma_cuon is None or str(raw_ma_cuon).strip().upper() == 'NONE' or str(raw_ma_cuon).strip() == '':
                                continue
                                
                            ma_cuon = str(raw_ma_cuon).replace('.0', '').strip().upper()
                            phieu_can = str(ws_val.cell(r, 5).value or '').replace('.0', '').strip()
                            ma_hoa = ma_cuon + 'A'
                            
                            net_ct = to_float(ws_val.cell(r, 8).value)
                            gross_ct = to_float(ws_val.cell(r, 9).value)
                            can_ban = to_float(ws_val.cell(r, 10).value)
                            chenh_lech_kg_file = to_float(ws_val.cell(r, 11).value)
                            lech_pt_bbnt = to_float(ws_val.cell(r, 12).value)
                            
                            sum_net += net_ct
                            sum_gross += gross_ct
                            sum_can_ban += can_ban
                            sum_chenh_lech += chenh_lech_kg_file

                            # Logic KSNB 4 bên
                            tc_phanloai = dict_tc_phanloai.get(phieu_can, 'Nhập hàng')
                            kl_net_pl = dict_pl_master.get(sheet_name, {}).get(ma_hoa, 0)
                            doi_chieu = net_ct - kl_net_pl
                            
                            kl_bcbv, spc_bcbv = 0, ''
                            for k_txt, v_phieu in dict_bv_phieu.items():
                                if ma_cuon in k_txt:
                                    spc_bcbv = v_phieu
                                    kl_bcbv = dict_bv_kl.get(k_txt, 0)
                                    break
                            
                            dc_spc = (phieu_can == str(spc_bcbv)) if phieu_can else False
                            kl_tc = dict_tc_kl.get(phieu_can, 0.0)
                            dc_tc_bv = (kl_tc == to_float(kl_bcbv))
                            dc_tc_bbnt = (kl_tc == can_ban)
                            
                            gross_ct_div = gross_ct if gross_ct != 0 else 1
                            chenh_lech_kg = kl_tc - gross_ct
                            chenh_lech_ds = chenh_lech_kg / gross_ct_div
                            
                            dc_ds = (round(chenh_lech_ds, 6) == round(lech_pt_bbnt, 6))
                            so_voi_qd = 'VUOT DUNG SAI' if abs(chenh_lech_ds) > 0.003 else 'OK'
                            
                            kl_035 = opm_035_po.get(ma_hoa, 0) - abs(opm_035_rtv.get(ma_hoa, 0))
                            cl_035 = kl_net_pl - kl_035
                            
                            kl_031 = opm_031.get(ma_hoa, 0)
                            cl_031 = 0 if kl_031 == 0 else (kl_net_pl - kl_031)
                            
                            kl_03g = opm_03g.get(ma_hoa, 0)
                            cl_03g = kl_net_pl - kl_03g
                            
                            row_data = [
                                tc_phanloai, ma_hoa, kl_net_pl, doi_chieu,
                                kl_bcbv, spc_bcbv, dc_spc, kl_tc, dc_tc_bv, dc_tc_bbnt,
                                chenh_lech_kg, chenh_lech_ds, dc_ds,
                                so_voi_qd, kl_035, cl_035, kl_031, cl_031, kl_03g, cl_03g
                            ]
                            
                            # Ghi kết quả vào bản Formulas
                            for i, val in enumerate(row_data):
                                cell = ws_form.cell(row=r, column=14+i)
                                cell.value = val
                                cell.border = border_thin
                                
                                c_name = new_cols[i]
                                is_error = False
                                if c_name in ['Đối chiếu số phiếu cân', 'Đối chiếu KL TC & KL BCBV', 'Đối chiếu KL TC & KL BBNT', 'Đối chiếu dung sai'] and val is False:
                                    is_error = True
                                elif c_name == 'So với quy định ( >+-0.3%)' and str(val) == 'VUOT DUNG SAI':
                                    is_error = True
                                elif c_name in ['Đối chiếu', 'Chênh lệch ', 'Chênh lệch  ', 'Chênh lệch   '] and isinstance(val, (int, float)) and val != 0:
                                    is_error = True
                                    
                                if is_error:
                                    cell.fill = red_fill
                                    cell.font = red_font

                        # --- KIỂM TOÁN DÒNG TỔNG CỘNG ---
                        if tong_row != -1 and tong_row > 10:
                            cols_to_check = {8: ('H', sum_net), 9: ('I', sum_gross), 10: ('J', sum_can_ban), 11: ('K', sum_chenh_lech)}
                            
                            # Check 4 cột SUM
                            for col_idx, (col_letter, manual_sum) in cols_to_check.items():
                                cell_form = ws_form.cell(row=tong_row, column=col_idx)
                                cell_val = ws_val.cell(row=tong_row, column=col_idx)
                                
                                val_f = cell_form.value
                                is_error = False
                                
                                if isinstance(val_f, str) and val_f.startswith('='):
                                    expected_formula = f"=SUM({col_letter}10:{col_letter}{tong_row-1})"
                                    if val_f.replace(' ', '').replace('$', '').upper() != expected_formula:
                                        is_error = True 
                                else:
                                    if abs(to_float(cell_val.value) - manual_sum) > 0.1:
                                        is_error = True 
                                        
                                if is_error:
                                    cell_form.fill = red_fill
                                    cell_form.font = red_font
                            
                            # Check cột %
                            cell_l_form = ws_form.cell(row=tong_row, column=12)
                            cell_l_val = ws_val.cell(row=tong_row, column=12)
                            val_l = cell_l_form.value
                            is_error_l = False
                            
                            if isinstance(val_l, str) and val_l.startswith('='):
                                val_upper = val_l.replace(' ', '').replace('$', '').upper()
                                if f"K{tong_row}/I{tong_row}" not in val_upper:
                                    is_error_l = True 
                            else:
                                expected_pct = sum_chenh_lech / sum_gross if sum_gross != 0 else 0
                                if abs(to_float(cell_l_val.value) - expected_pct) > 0.0001:
                                    is_error_l = True
                                    
                            if is_error_l:
                                cell_l_form.fill = red_fill
                                cell_l_form.font = red_font

                    # --- LƯU FILE KẾT QUẢ ---
                    my_bar.progress(95, text="💾 Đang đóng gói File Excel...")
                    output = io.BytesIO()
                    wb_formulas.save(output)
                    
                    my_bar.progress(100, text="✅ Hoàn tất 100%!")
                    st.success(f'🎉 Kế toán đã rà soát xong {len(selected_sheets)} lô hàng (Đã tự động loại bỏ lỗi VLOOKUP)!')
                    st.download_button(
                        '📥 Tải File Biên Bản Đối Chiếu (Bảo lưu Form & Tô đỏ)',
                        data=output.getvalue(),
                        file_name='BBNT_Nghiem_Thu_Doi_Chieu_Tong_Hop.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    )

                except Exception as e:
                    my_bar.empty() # Ẩn thanh % nếu lỗi
                    st.error(f'⚠️ Lỗi hệ thống khi xử lý dữ liệu: {e}')
    else:
        st.info('👉 Hãy nạp tối thiểu File BBNT để bắt đầu.')

doi_chieu_hrc_tung_file()
