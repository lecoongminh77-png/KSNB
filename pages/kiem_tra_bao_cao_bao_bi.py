import sys
import os
import io
import time
import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill, Font, Border, Side

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN CHUẨN APPLE & THANH TIỆN ÍCH CỐ ĐỊNH
# -------------------------------------------------------------
st.set_page_config(
    page_title='Kiểm Soát Báo Cáo Cắt Tấm', page_icon='🛡️', layout='wide'
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
    
    /* Nút Download có màu xanh lá */
    div.stDownloadButton > button:first-child {
        background: #28a745 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding: 10px 20px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        box-shadow: 0 4px 12px rgba(40, 167, 69, 0.2) !important;
    }
    div.stDownloadButton > button:first-child:hover {
        background: #218838 !important;
    }

    /* Bắt ép Container của Streamlit biến thành Thẻ Apple */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff;
        border-radius: 16px;
        padding: 15px;
        border: 1px solid rgba(0, 0, 0, 0.08);
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.03);
    }

    /* Dashboard Metrics Card */
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 15px;
        border: 1px solid rgba(0,0,0,0.05);
        box-shadow: 0 2px 8px rgba(0,0,0,0.02);
        text-align: center;
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

# Thanh điều hướng (Chỉ giữ nút Trang chủ)
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', use_container_width=True):
        st.switch_page('app_ksnb.py')

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)

st.markdown(
    """
<div class="apple-header-box">
    <div>
        <h1 class="apple-main-title">🏭 Kiểm Soát Báo Cáo Bao Bì & Cắt Tấm</h1>
    </div>
</div>
""", unsafe_allow_html=True)
st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>", unsafe_allow_html=True)


# ==============================================================================
# HÀM HỖ TRỢ XỬ LÝ SỐ ÂM & LÀM SẠCH DỮ LIỆU
# ==============================================================================
def is_negative(val):
    if pd.isna(val): return False
    if isinstance(val, (int, float)): return val < 0
    if isinstance(val, str):
        val = val.strip()
        if (val.startswith('(') and val.endswith(')')) or val.startswith('-'):
            return True
    return False

def clean_ma_lo(val):
    if pd.isna(val): return ""
    return str(val).strip().upper()

# ==============================================================================
# GIAO DIỆN CHÍNH
# ==============================================================================
def kiem_soat_san_xuat():
    st.caption('Kiểm soát nội bộ • Tự động đối soát dữ liệu Báo cáo cắt tấm & OPM 105')
    
    st.markdown("### 📥 Nạp Dữ Liệu Nguồn")
    
    col_t1, col_t2, _ = st.columns([2, 2, 6])
    with col_t1:
        thang_chon = st.selectbox('Chọn Tháng', list(range(1, 13)), index=9)
    with col_t2:
        nam_chon = st.selectbox('Chọn Năm', [2024, 2025, 2026, 2027], index=2)

    chuoi_yymm = f"{str(nam_chon)[-2:]}{thang_chon:02d}"
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_upload1, col_upload2 = st.columns(2, gap="large")
    
    with col_upload1:
        with st.container(border=True):
            st.markdown("""
<div class="upload-title">1️⃣ Dữ liệu Báo cáo Cắt tấm</div>
<div class="guide-box">
<div class="guide-header">📌 Báo cáo cắt tấm (BCSX)</div>
<div class="guide-text">Hệ thống sẽ quét sheet: NXT_TAM_LOT, QLSX NIENG, NHUNG... và NXT NGUYEN LIEU.</div>
</div>
""", unsafe_allow_html=True)
            f_bc = st.file_uploader('Tải lên File BÁO CÁO CẮT TẤM', type=['xlsx', 'xls'])

    with col_upload2:
        with st.container(border=True):
            st.markdown("""
<div class="upload-title">2️⃣ Dữ liệu Giao dịch ERP 105</div>
<div class="guide-box">
<div class="guide-header">📌 Báo cáo OPM 105</div>
<div class="guide-text">
Tích chọn Kho: <b>031</b> | Giao dịch: <b>Move order</b><br>
Thời gian: Chọn đúng tháng đang kiểm tra.
</div>
</div>
""", unsafe_allow_html=True)
            f_105 = st.file_uploader('Tải lên File OPM 105', type=['xlsx', 'xls'])

    if st.button('🚀 Bắt đầu đối chiếu tự động', use_container_width=True):
        if not f_bc and not f_105:
            st.warning("⚠️ Vui lòng tải lên ít nhất một file để kiểm tra!")
            return

        progress_bar = st.progress(0, text="Khởi động tiến trình kiểm tra...")
        
        bc_sheets = []
        if f_bc:
            xl_bc = pd.ExcelFile(f_bc)
            bc_sheets = xl_bc.sheet_names

        # ==============================================================
        # PHẦN 1: KIỂM TRA BÁO CÁO CẮT TẤM
        # ==============================================================
        if f_bc is not None:
            if 'NXT_TAM_LOT' not in bc_sheets:
                st.error("🚨 Không tìm thấy sheet 'NXT_TAM_LOT' trong file Báo Cáo. Có thể bạn đã tải nhầm file!")
            else:
                progress_bar.progress(10, text="Đang đọc dữ liệu NXT_TAM_LOT...")
                st.markdown("### 📊 KẾT QUẢ: BÁO CÁO CẮT TẤM")
                try:
                    df_nxt_tam_lot = pd.read_excel(f_bc, sheet_name='NXT_TAM_LOT', header=None, dtype=str)
                    
                    ma_kien_am = []
                    for idx, row in df_nxt_tam_lot.iterrows():
                        ma_kien = str(row[1]).strip() if pd.notna(row[1]) else ""
                        so_tam = row[11]
                        if ma_kien != "" and ma_kien.lower() != "nan" and idx > 1:
                            if is_negative(so_tam):
                                ma_kien_am.append((ma_kien, so_tam))
                    
                    if len(ma_kien_am) > 0:
                        st.error(f"🚨 PHÁT HIỆN SỐ TẤM BỊ ÂM TRONG SHEET 'NXT_TAM_LOT':")
                        for mk, st_am in ma_kien_am:
                            st.write(f"- Mã kiện: **{mk}** | Số tấm: **{st_am}**")
                    else:
                        st.success("✅ Sheet 'NXT_TAM_LOT': Không có trường hợp nào số tấm bị âm.")

                    progress_bar.progress(30, text="Đang quét các sheet QLSX...")

                    ds_sheets = {
                        'QLSX NIENG': 'SXNI',
                        'QLSX NHUNG': 'SXNH',
                        'QLSX_CAT_TAM MAT TRANG': 'CTMT',
                        'QLSX_CAT_TAM LOT NGOAI': 'CTLN',
                        'QLSX_CAT_TAM LOT TRONG': 'CTLT',
                        'QLSX TP MAT TRANG': 'SXMT'
                    }

                    nxt_ma_kiens = df_nxt_tam_lot[1].dropna().astype(str).str.strip().tolist()
                    step_progress, progress_increment = 30, 40 / len(ds_sheets)

                    for sheet_name, prefix in ds_sheets.items():
                        step_progress += progress_increment
                        progress_bar.progress(int(step_progress), text=f"Đang kiểm tra: {sheet_name}...")
                        
                        sheet_real_name = next((s for s in bc_sheets if sheet_name.lower() in s.lower()), None)
                        
                        if sheet_real_name:
                            try:
                                df_qlsx = pd.read_excel(f_bc, sheet_name=sheet_real_name, header=None, dtype=str)
                                max_seq = 0
                                for val in df_qlsx[12].dropna():
                                    val_str = str(val).strip()
                                    if val_str.startswith(f"{prefix}{chuoi_yymm}"):
                                        try:
                                            seq = int(val_str[-3:])
                                            if seq > max_seq: max_seq = seq
                                        except ValueError:
                                            pass
                                
                                with st.expander(f"Kết quả đối chiếu: {sheet_real_name} ({prefix}{chuoi_yymm})"):
                                    if max_seq == 0:
                                        st.info(f"Không có mã kiện thuộc tháng {thang_chon}/{nam_chon}.")
                                    else:
                                        st.write(f"Số lượng mã kiện lớn nhất trong tháng: **{max_seq}** kiện.")
                                        thieu = [f"{prefix}{chuoi_yymm}{i:03d}" for i in range(1, max_seq + 1) if f"{prefix}{chuoi_yymm}{i:03d}" not in nxt_ma_kiens]
                                        
                                        if len(thieu) > 0:
                                            st.error(f"❌ Phát hiện thiếu **{len(thieu)}** mã kiện bên NXT_TAM_LOT:")
                                            st.write(", ".join(thieu))
                                        else:
                                            st.success("✅ Đủ số lượng, không rớt mã lô nào.")
                            except Exception as e:
                                st.error(f"Lỗi đọc {sheet_real_name}: {e}")
                        else:
                            st.warning(f"Không tìm thấy sheet '{sheet_name}'.")
                except Exception as e:
                    st.error(f"Lỗi hệ thống: {e}")

        # ==============================================================
        # PHẦN 2: ĐỐI CHIẾU OPM 105 VÀ NXT NGUYÊN LIỆU (CỘNG ÂM DƯƠNG = 0)
        # ==============================================================
        if f_105 is not None and f_bc is not None:
            progress_bar.progress(80, text="Đang đồng bộ và gom nhóm dữ liệu ERP 105...")
            st.markdown("### 📊 DASHBOARD: ĐỐI CHIẾU ERP 105 & NXT NGUYÊN LIỆU")
            
            try:
                # 1. Chuẩn bị dữ liệu 105 
                # Chống mất số 0 ở đầu bằng cách đọc ép kiểu String
                df_105_raw = pd.read_excel(f_105, header=None, dtype=str)
                df_105_raw['Mã Lô 105'] = df_105_raw[5].apply(clean_ma_lo)
                df_105_raw['Khối Lượng 105'] = pd.to_numeric(df_105_raw[9], errors='coerce').fillna(0)
                
                # Lọc lấy các dòng có Mã Lô dài hơn 3 ký tự để loại bỏ dòng tiêu đề nhiễu
                df_105_filtered = df_105_raw[(df_105_raw['Mã Lô 105'] != "") & (df_105_raw['Mã Lô 105'].str.len() > 3)]
                df_105_grouped = df_105_filtered.groupby('Mã Lô 105', as_index=False)['Khối Lượng 105'].sum()

                # 2. Chuẩn bị dữ liệu NXT NGUYÊN LIỆU (LẤY CỘT H DƯƠNG)
                nl_sheet_name = next((s for s in bc_sheets if "NXT NGUYEN LIEU" in s.upper()), None)
                if not nl_sheet_name:
                    st.error("🚨 Không tìm thấy sheet có tên 'NXT NGUYEN LIEU' trong Báo cáo.")
                else:
                    df_nl = pd.read_excel(f_bc, sheet_name=nl_sheet_name, header=None, dtype=str)
                    dict_nl = {}
                    for idx, row in df_nl.iterrows():
                        ma_lo_nxt = clean_ma_lo(row[1])
                        # Lấy cột H (Index 7) - Số dương
                        kl_nxt = pd.to_numeric(row[7], errors='coerce') if pd.notna(row[7]) else 0
                        
                        if ma_lo_nxt != "" and len(ma_lo_nxt) > 3:
                            dict_nl[ma_lo_nxt] = dict_nl.get(ma_lo_nxt, 0) + kl_nxt

                    # 3. Đối chiếu và tính Delta
                    ket_qua_105 = []
                    tong_so_lo = len(df_105_grouped)
                    so_dat = so_lech = so_thieu = 0

                    for idx, row in df_105_grouped.iterrows():
                        ma_105 = row['Mã Lô 105']
                        kl_105 = row['Khối Lượng 105'] # Số âm
                        
                        if ma_105 in dict_nl:
                            kl_nxt = dict_nl[ma_105] # Số dương
                            
                            # Cộng hai số lại, nếu bằng 0 là Đạt
                            chenh_lech = kl_105 + kl_nxt 
                            
                            if round(chenh_lech, 2) == 0:
                                trang_thai = "Đạt"
                                so_dat += 1
                            else:
                                trang_thai = "Lệch khối lượng"
                                so_lech += 1
                        else:
                            kl_nxt = 0
                            chenh_lech = kl_105
                            trang_thai = "Không có thông tin trên ERP (NXT)"
                            so_thieu += 1
                            
                        ket_qua_105.append({
                            'Mã Lô': ma_105,
                            'KL ERP 105 (Âm)': kl_105,
                            'KL Thực NXT (Dương)': kl_nxt,
                            'Chênh Lệch (kg)': round(chenh_lech, 2),
                            'Trạng Thái': trang_thai
                        })

                    # 4. Hiển thị Dashboard 4 Chỉ số
                    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                    col_m1.metric("📦 Tổng Số Lô", f"{tong_so_lo} Lô")
                    col_m2.metric("✅ Khớp Dữ Liệu", f"{so_dat} Lô")
                    col_m3.metric("⚠️ Lệch Khối Lượng", f"{so_lech} Lô", delta=f"{so_lech} lỗi", delta_color="inverse")
                    col_m4.metric("🚨 Không Thông Tin", f"{so_thieu} Lô", delta=f"{so_thieu} lỗi", delta_color="inverse")

                    # 5. Hiển thị Bảng Dữ Liệu
                    if tong_so_lo > 0:
                        df_kq = pd.DataFrame(ket_qua_105)
                        st.dataframe(df_kq, use_container_width=True)

                        # 6. TẠO FILE EXCEL VỚI MÀU SẮC
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_kq.to_excel(writer, index=False, sheet_name='Ket_Qua_105')
                            worksheet = writer.sheets['Ket_Qua_105']
                            
                            fill_green = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
                            font_green = Font(color='006100')
                            fill_yellow = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
                            font_yellow = Font(color='9C5700')
                            fill_red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                            font_red = Font(color='9C0006')
                            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                            for row in range(2, len(df_kq) + 2):
                                trang_thai_val = worksheet.cell(row=row, column=5).value
                                for col in range(1, 6):
                                    cell = worksheet.cell(row=row, column=col)
                                    cell.border = thin_border
                                    if trang_thai_val == "Đạt":
                                        cell.fill = fill_green
                                        cell.font = font_green
                                    elif trang_thai_val == "Lệch khối lượng":
                                        cell.fill = fill_yellow
                                        cell.font = font_yellow
                                    elif "Không có thông tin" in str(trang_thai_val):
                                        cell.fill = fill_red
                                        cell.font = font_red
                                        
                            for col in worksheet.columns:
                                max_length = 0
                                column = col[0].column_letter
                                for cell in col:
                                    try:
                                        if len(str(cell.value)) > max_length: max_length = len(cell.value)
                                    except: pass
                                adjusted_width = (max_length + 2)
                                worksheet.column_dimensions[column].width = adjusted_width

                        st.download_button(
                            label="📥 Tải Báo Cáo Đối Chiếu OPM 105 (Đã Tô Màu)",
                            data=output.getvalue(),
                            file_name=f"Ket_Qua_Doi_Chieu_105_{chuoi_yymm}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.info("Không tìm thấy dữ liệu hợp lệ trên file OPM 105. Vui lòng kiểm tra lại cấu trúc file!")

            except Exception as e:
                st.error(f"Lỗi hệ thống khi xử lý OPM 105: {e}")

        elif f_105 is not None and f_bc is None:
            st.warning("⚠️ Để đối chiếu file OPM 105, bạn phải tải thêm file Báo Cáo Cắt Tấm để đối soát với NXT Nguyên Liệu.")

        progress_bar.progress(100, text="✅ Hoàn tất kiểm tra toàn bộ dữ liệu!")
        time.sleep(1)
        progress_bar.empty()

kiem_soat_san_xuat()
