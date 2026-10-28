import sys
import os
import io
import re
import pandas as pd
import streamlit as st

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

    /* Bắt ép Container của Streamlit biến thành Thẻ Apple */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff;
        border-radius: 16px;
        padding: 15px;
        border: 1px solid rgba(0, 0, 0, 0.08);
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.03);
    }

    /* Các thành phần chữ bên trong thẻ */
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

# Thanh điều hướng (Chỉ giữ nút Trang chủ)
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', use_container_width=True):
        st.write("Chức năng điều hướng trang chủ...") # Đổi thành lệnh switch page của bạn nếu cần

st.markdown(
    "<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>",
    unsafe_allow_html=True,
)

# Header Trang
st.markdown(
    """
<div class="apple-header-box">
    <div>
        <h1 class="apple-main-title">📊 Kiểm Soát Dữ Liệu Sản Xuất & ERP</h1>
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
# HÀM HỖ TRỢ XỬ LÝ SỐ ÂM
# ==============================================================================
def is_negative(val):
    if pd.isna(val):
        return False
    if isinstance(val, (int, float)):
        return val < 0
    if isinstance(val, str):
        val = val.strip()
        # Nếu có dấu ngoặc đơn báo số âm (VD: (340)) hoặc dấu trừ (-340)
        if (val.startswith('(') and val.endswith(')')) or val.startswith('-'):
            return True
    return False

# ==============================================================================
# GIAO DIỆN CHÍNH
# ==============================================================================
def kiem_soat_san_xuat():
    st.caption('Kiểm soát nội bộ • Tự động đối soát dữ liệu Báo cáo cắt tấm & OPM 105')
    
    st.markdown("### 📥 Nạp Dữ Liệu Nguồn")
    
    # ---------------------------------------------------------
    # KHU VỰC CHỌN THỜI GIAN
    # ---------------------------------------------------------
    col_t1, col_t2, _ = st.columns([2, 2, 6])
    with col_t1:
        thang_chon = st.selectbox('Chọn Tháng', list(range(1, 13)), index=9) # Mặc định tháng 10
    with col_t2:
        nam_chon = st.selectbox('Chọn Năm', [2024, 2025, 2026, 2027], index=2) # Mặc định 2026

    # Tạo chuỗi thời gian dạng YYMM (VD: 2610)
    chuoi_yymm = f"{str(nam_chon)[-2:]}{thang_chon:02d}"

    st.markdown("<br>", unsafe_allow_html=True)
    
    # ---------------------------------------------------------
    # KHU VỰC HƯỚNG DẪN & UPLOAD (SỬ DỤNG CONTAINER)
    # ---------------------------------------------------------
    col_upload1, col_upload2 = st.columns(2, gap="large")
    
    with col_upload1:
        with st.container(border=True):
            st.markdown("""
<div class="upload-title">1️⃣ Dữ liệu Báo cáo Cắt tấm</div>
<div class="guide-box">
<div class="guide-header">📌 1. Báo cáo cắt tấm (BCSX)</div>
<div class="guide-text">Bao gồm các Sheet: NXT_TAM_LOT, QLSX NIENG, QLSX NHUNG, QLSX_CAT_TAM MAT TRANG, QLSX_CAT_TAM LOT NGOAI, QLSX_CAT_TAM LOT TRONG, QLSX TP MAT TRANG...</div>
</div>
""", unsafe_allow_html=True)
            f_bc = st.file_uploader('Tải lên File BÁO CÁO CẮT TẤM', type=['xlsx', 'xls'])

    with col_upload2:
        with st.container(border=True):
            st.markdown("""
<div class="upload-title">2️⃣ Dữ liệu Giao dịch ERP 105</div>
<div class="guide-box">
<div class="guide-header">📌 2. Báo cáo OPM 105</div>
<div class="guide-text">
File mẫu: HSG_OPM_105_BCKiemSoatThongTinGDTheoLo_GD2<br>
<b>Lưu ý khi xuất hệ thống:</b><br>
- Tích chọn Kho: <b>031</b><br>
- Tích chọn: <b>Move order</b><br>
- Thời gian: <b>Chọn đúng 1 tháng cần kiểm tra (Ví dụ tháng đang chọn bên trên)</b>
</div>
</div>
""", unsafe_allow_html=True)
            f_105 = st.file_uploader('Tải lên File OPM 105', type=['xlsx', 'xls'])

    # ---------------------------------------------------------
    # XỬ LÝ LOGIC ĐỐI CHIẾU
    # ---------------------------------------------------------
    if st.button('🚀 Bắt đầu đối chiếu tự động', use_container_width=True):
        if not f_bc and not f_105:
            st.warning("⚠️ Vui lòng tải lên ít nhất một file để kiểm tra!")
            return

        with st.spinner('Đang tiến hành đối soát dữ liệu...'):
            # ==============================================================
            # PHẦN 1: KIỂM TRA BÁO CÁO CẮT TẤM (NXT_TAM_LOT & CÁC SHEET QLSX)
            # ==============================================================
            if f_bc is not None:
                st.markdown("### 📊 KẾT QUẢ: BÁO CÁO CẮT TẤM")
                try:
                    # 1. Quét số âm trong NXT_TAM_LOT
                    df_nxt_tam_lot = pd.read_excel(f_bc, sheet_name='NXT_TAM_LOT', header=None)
                    
                    ma_kien_am = []
                    # Giả định: Cột B (Mã số kiện) là index 1, Cột L (Số tấm) là index 11
                    for idx, row in df_nxt_tam_lot.iterrows():
                        ma_kien = str(row[1]).strip() if pd.notna(row[1]) else ""
                        so_tam = row[11]
                        
                        if ma_kien != "" and ma_kien.lower() != "nan" and idx > 1: # Bỏ qua header
                            if is_negative(so_tam):
                                ma_kien_am.append((ma_kien, so_tam))
                    
                    if len(ma_kien_am) > 0:
                        st.error(f"🚨 PHÁT HIỆN SỐ TẤM BỊ ÂM TRONG SHEET 'NXT_TAM_LOT':")
                        for mk, st_am in ma_kien_am:
                            st.write(f"- Mã kiện: **{mk}** | Số tấm: **{st_am}**")
                    else:
                        st.success("✅ Sheet 'NXT_TAM_LOT': Không có trường hợp nào số tấm bị âm.")

                    # 2. Đối chiếu số lượng giữa QLSX và NXT_TAM_LOT
                    ds_sheets = {
                        'QLSX NIENG': 'SXNI',
                        'QLSX NHUNG': 'SXNH',
                        'QLSX_CAT_TAM MAT TRANG': 'CTMT',
                        'QLSX_CAT_TAM LOT NGOAI': 'CTLN',
                        'QLSX_CAT_TAM LOT TRONG': 'CTLT',
                        'QLSX TP MAT TRANG': 'SXMT'
                    }

                    # Tạo danh sách các mã kiện có bên trong NXT_TAM_LOT để so sánh nhanh
                    nxt_ma_kiens = df_nxt_tam_lot[1].dropna().astype(str).str.strip().tolist()

                    for sheet_name, prefix in ds_sheets.items():
                        try:
                            df_qlsx = pd.read_excel(f_bc, sheet_name=sheet_name, header=None)
                            
                            # Tìm Max Seq trong tháng cho sheet này
                            max_seq = 0
                            # Cột M (Mã số kiện) = index 12
                            for val in df_qlsx[12].dropna():
                                val_str = str(val).strip()
                                # Kiểm tra xem mã có khớp với Prefix + Tháng/Năm hiện tại không (VD: SXNI2610...)
                                if val_str.startswith(f"{prefix}{chuoi_yymm}"):
                                    try:
                                        # Lấy 3 số cuối
                                        seq = int(val_str[-3:])
                                        if seq > max_seq:
                                            max_seq = seq
                                    except ValueError:
                                        pass
                            
                            with st.expander(f"Kết quả đối chiếu: {sheet_name} ({prefix}{chuoi_yymm})"):
                                if max_seq == 0:
                                    st.info(f"Không phát hiện mã kiện nào thuộc tháng {thang_chon}/{nam_chon} trong sheet này.")
                                else:
                                    st.write(f"Số lượng mã kiện ghi nhận lớn nhất trong tháng: **{max_seq}** kiện.")
                                    thieu = []
                                    # Kiểm tra xem từ 001 đến max_seq có đủ mặt bên NXT_TAM_LOT không
                                    for i in range(1, max_seq + 1):
                                        ma_can_tim = f"{prefix}{chuoi_yymm}{i:03d}"
                                        if ma_can_tim not in nxt_ma_kiens:
                                            thieu.append(ma_can_tim)
                                    
                                    if len(thieu) > 0:
                                        st.error(f"❌ Phát hiện thiếu **{len(thieu)}** mã kiện bên sheet NXT_TAM_LOT:")
                                        st.write(", ".join(thieu))
                                    else:
                                        st.success("✅ Đã đủ số lượng mã kiện, không có mã nào bị thiếu khi đối chiếu với NXT_TAM_LOT.")

                        except ValueError as ve:
                            st.warning(f"Không tìm thấy sheet '{sheet_name}' trong file Báo cáo.")
                        except Exception as e:
                            st.error(f"Lỗi khi đọc {sheet_name}: {e}")

                except Exception as e:
                    st.error(f"Lỗi hệ thống khi xử lý File Báo Cáo Cắt Tấm: {e}")


            # ==============================================================
            # PHẦN 2: KIỂM TRA FILE 105 VÀ NXT NGUYEN LIEU
            # ==============================================================
            if f_105 is not None and f_bc is not None:
                st.markdown("### 📊 KẾT QUẢ: ĐỐI CHIẾU OPM 105 & NXT NGUYÊN LIỆU")
                try:
                    # Đọc file 105 - Dữ liệu bắt đầu từ dòng 12 (skiprows=11)
                    df_105 = pd.read_excel(f_105, header=None, skiprows=11)
                    
                    # Đọc sheet NXT NGUYEN LIEU từ file BCSX (Cột B: index 1, Cột Q: index 16)
                    # Lưu ý: Sheet name có thể có dấu cách ở cuối 'NXT NGUYEN LIEU '
                    try:
                        df_nl = pd.read_excel(f_bc, sheet_name='NXT NGUYEN LIEU ', header=None)
                    except:
                        df_nl = pd.read_excel(f_bc, sheet_name='NXT NGUYEN LIEU', header=None)

                    # Tạo Dict mapping (Mã Lô -> Khối Lượng) từ sheet NXT NGUYEN LIEU
                    dict_nl = {}
                    for idx, row in df_nl.iterrows():
                        ma_lo = str(row[1]).strip() if pd.notna(row[1]) else ""
                        kl_nl = pd.to_numeric(row[16], errors='coerce')
                        if ma_lo != "" and ma_lo.lower() != "nan":
                            dict_nl[ma_lo] = kl_nl

                    ket_qua_105 = []
                    # Quét file 105 (Cột F / Mã Lô = index 5, Cột J / Khối Lượng = index 9)
                    for idx, row in df_105.iterrows():
                        ma_lo_105 = str(row[5]).strip() if pd.notna(row[5]) else ""
                        kl_105 = abs(pd.to_numeric(row[9], errors='coerce')) # Lấy trị tuyệt đối vì trên 105 move issue thường hiển thị âm
                        
                        if ma_lo_105 != "" and ma_lo_105.lower() != "nan":
                            if ma_lo_105 in dict_nl:
                                kl_nl = dict_nl[ma_lo_105]
                                
                                # Đối chiếu
                                if pd.isna(kl_nl):
                                    trang_thai = "Thiếu dữ liệu KL bên NXT"
                                elif round(kl_105, 2) == round(abs(kl_nl), 2):
                                    trang_thai = "Đạt"
                                else:
                                    trang_thai = f"Lệch (105: {kl_105} | NXT: {kl_nl})"
                            else:
                                trang_thai = "Không có thông tin trên ERP (NXT)"
                            
                            ket_qua_105.append({
                                'Mã Lô 105': ma_lo_105,
                                'Khối Lượng 105': kl_105,
                                'Trạng Thái Đối Chiếu': trang_thai
                            })

                    if len(ket_qua_105) > 0:
                        df_ket_qua_105 = pd.DataFrame(ket_qua_105)
                        
                        so_loi = len(df_ket_qua_105[df_ket_qua_105['Trạng Thái Đối Chiếu'] != 'Đạt'])
                        if so_loi == 0:
                            st.success("✅ Toàn bộ lô trên file 105 đều khớp khối lượng với sheet NXT NGUYEN LIEU (Báo cáo đạt).")
                        else:
                            st.warning(f"⚠️ Phát hiện {so_loi} lô có vấn đề (lệch khối lượng hoặc không có thông tin):")
                        
                        st.dataframe(df_ket_qua_105, use_container_width=True)
                    else:
                        st.info("Không tìm thấy dữ liệu mã lô hợp lệ trên file OPM 105.")

                except Exception as e:
                    st.error(f"Lỗi hệ thống khi xử lý File OPM 105 & Nguyên Liệu: {e}")
            elif f_105 is not None and f_bc is None:
                st.warning("⚠️ Để đối chiếu file OPM 105, bạn phải tải thêm file Báo Cáo Cắt Tấm để lấy sheet NXT NGUYEN LIEU.")

kiem_soat_san_xuat()
