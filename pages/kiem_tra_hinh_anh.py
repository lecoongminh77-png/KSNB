import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image
import io
import time

try:
    import easyocr
except ImportError:
    st.error("⚠️ Bạn chưa cài thư viện EasyOCR. Hãy mở Terminal và gõ: pip install easyocr opencv-python-headless")

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN
# -------------------------------------------------------------
st.set_page_config(page_title='Kiểm Kê Bằng Hình Ảnh', page_icon='📸', layout='wide')

st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    .stApp { background-color: #f5f5f7; font-family: -apple-system, sans-serif; color: #1d1d1f; }
    .apple-main-title { font-size: 36px; font-weight: 700; color: #1d1d1f; margin-bottom: 20px; text-align: center; }
    div.stButton > button:first-child { background: #0071e3 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; padding: 10px 20px !important; font-weight: 600 !important; }
    div.stButton > button:first-child:hover { background: #0077ed !important; }
    [data-testid="stVerticalBlockBorderWrapper"] { background: #ffffff; border-radius: 16px; padding: 15px; border: 1px solid rgba(0,0,0,0.08); box-shadow: 0 2px 12px rgba(0,0,0,0.03); }
    div[data-testid="stMetric"] { background-color: #ffffff; border-radius: 12px; padding: 15px; border: 1px solid rgba(0,0,0,0.05); box-shadow: 0 2px 8px rgba(0,0,0,0.02); text-align: center; }
    .upload-title { font-size: 18px; font-weight: 700; margin-bottom: 10px; color: #0071e3;}
    .list-box { background: #fff; padding: 15px; border-radius: 10px; border: 1px solid #e2e8f0; height: 100%; }
</style>
""", unsafe_allow_html=True)

# Nút quay về
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', use_container_width=True):
        st.switch_page('app_ksnb.py')

st.markdown("<h1 class='apple-main-title'>📸 AI Kiểm Kê Bằng Hình Ảnh</h1>", unsafe_allow_html=True)
st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin-bottom: 30px;'>", unsafe_allow_html=True)


# -------------------------------------------------------------
# HÀM HỖ TRỢ & AI (CACHE ĐỂ CHẠY NHANH)
# -------------------------------------------------------------
@st.cache_resource
def load_ocr_reader():
    """Tải model AI một lần duy nhất vào bộ nhớ RAM để tăng tốc"""
    return easyocr.Reader(['vi', 'en'], gpu=False)

def parse_and_check_date(val, t_month, t_year):
    """Kiểm tra ngày tháng xem có khớp với điều kiện lọc không"""
    if pd.isna(val) or str(val).strip() == "": return False
    try:
        dt = pd.to_datetime(str(val), errors='coerce')
        if pd.notna(dt) and dt.month == t_month and dt.year == t_year: return True
    except: pass
    return False

# -------------------------------------------------------------
# KHU VỰC TẢI DỮ LIỆU
# -------------------------------------------------------------
st.markdown("### 📥 1. Thiết Lập Mốc Thời Gian Kiểm Kê")
col_t1, col_t2, _ = st.columns([2, 2, 6])
with col_t1: thang_chon = st.selectbox('Chọn Tháng', list(range(1, 13)), index=9)
with col_t2: nam_chon = st.selectbox('Chọn Năm', [2024, 2025, 2026, 2027], index=2)

st.markdown("<br>", unsafe_allow_html=True)

col_upload1, col_upload2 = st.columns(2, gap="large")
with col_upload1:
    with st.container(border=True):
        st.markdown("<div class='upload-title'>📄 2. Nạp File Excel (Kế hoạch)</div>", unsafe_allow_html=True)
        st.write("Hệ thống sẽ lấy danh sách mã lô từ sheet GIAO NHAN TP GCBB và NXT NGUYEN LIEU.")
        f_excel = st.file_uploader("Tải file Excel", type=['xlsx', 'xls'], label_visibility="collapsed")

with col_upload2:
    with st.container(border=True):
        st.markdown("<div class='upload-title'>📸 3. Nạp Hình Ảnh (Thực tế)</div>", unsafe_allow_html=True)
        st.write("Bạn có thể chọn **nhiều hình ảnh** cùng một lúc để máy quét hàng loạt.")
        f_images = st.file_uploader("Tải hình ảnh", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True, label_visibility="collapsed")

# -------------------------------------------------------------
# LOGIC XỬ LÝ & ĐỐI CHIẾU
# -------------------------------------------------------------
if st.button('🔍 Bắt Đầu Phân Tích Bằng AI', use_container_width=True):
    if not f_excel or not f_images:
        st.warning("⚠️ Vui lòng tải lên ĐỦ file Excel và ít nhất 1 hình ảnh để tiến hành kiểm kê!")
    else:
        progress_bar = st.progress(0, text="Đang bóc tách danh sách mã lô từ File Excel...")
        time.sleep(0.5)

        try:
            # ==========================================
            # BƯỚC 1: LỌC DANH SÁCH MÃ TỪ EXCEL
            # ==========================================
            ds_ma_can_tim = set()
            xls = pd.ExcelFile(f_excel)
            sheet_names = xls.sheet_names
            
            # Nguồn 1: GIAO NHAN TP GCBB (Cột B thời gian, Cột J mã)
            if 'GIAO NHAN TP GCBB' in sheet_names:
                df1 = pd.read_excel(xls, sheet_name='GIAO NHAN TP GCBB', header=None, dtype=str)
                for idx, row in df1.iterrows():
                    if parse_and_check_date(row[1], thang_chon, nam_chon): # Cột B là index 1
                        ma = str(row[9]).strip().upper() # Cột J là index 9
                        if ma and ma != 'NAN' and len(ma) > 3:
                            ds_ma_can_tim.add(ma)

            # Nguồn 2: NXT NGUYEN LIEU (Cột N thời gian, Cột B mã)
            sheet2 = next((s for s in sheet_names if 'NXT NGUYEN LIEU' in s.upper()), None)
            if sheet2:
                df2 = pd.read_excel(xls, sheet_name=sheet2, header=None, dtype=str)
                for idx, row in df2.iterrows():
                    if parse_and_check_date(row[13], thang_chon, nam_chon): # Cột N là index 13
                        ma = str(row[1]).strip().upper() # Cột B là index 1
                        if ma and ma != 'NAN' and len(ma) > 3:
                            ds_ma_can_tim.add(ma)

            if len(ds_ma_can_tim) == 0:
                progress_bar.empty()
                st.error(f"🚨 Không tìm thấy bất kỳ mã lô nào thuộc tháng {thang_chon}/{nam_chon} trong file Excel!")
                st.stop()

            # ==========================================
            # BƯỚC 2: AI ĐỌC HÌNH ẢNH HÀNG LOẠT
            # ==========================================
            reader = load_ocr_reader()
            
            ma_da_thay = set()
            hinh_nghi_ngo = [] # Lưu các hình không có mã nào khớp
            
            tong_hinh = len(f_images)
            for i, img_file in enumerate(f_images):
                # Cập nhật thanh tiến trình
                ptram = int(20 + (80 * (i + 1) / tong_hinh))
                progress_bar.progress(ptram, text=f"🤖 AI đang quét tấm ảnh thứ {i+1}/{tong_hinh}: {img_file.name}...")
                
                # Mở hình và đọc chữ
                image = Image.open(img_file)
                image_np = np.array(image)
                ocr_results = reader.readtext(image_np, detail=0)
                chuoi_hinh_anh = "".join(ocr_results).upper().replace(" ", "").replace("-", "")
                
                # Dò tìm
                tim_thay_trong_hinh = False
                for ma in ds_ma_can_tim:
                    ma_clean = ma.replace(" ", "").replace("-", "")
                    if ma_clean in chuoi_hinh_anh:
                        ma_da_thay.add(ma)
                        tim_thay_trong_hinh = True
                
                # Nếu không tìm thấy mã nào hợp lệ trong tấm ảnh này
                if not tim_thay_trong_hinh:
                    hinh_nghi_ngo.append({
                        'tên': img_file.name,
                        'hình': image,
                        'chữ_nhận_diện': ocr_results
                    })

            progress_bar.progress(100, text="✅ Hoàn tất phân tích!")
            time.sleep(1)
            progress_bar.empty()

            # ==========================================
            # BƯỚC 3: BÁO CÁO KẾT QUẢ KIỂM KÊ
            # ==========================================
            ma_thieu = ds_ma_can_tim - ma_da_thay
            
            st.markdown("### 📊 BẢNG TỔNG KẾT KIỂM KÊ")
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("📋 Tổng Mã Cần Kiểm Kê", f"{len(ds_ma_can_tim)} Lô")
            col_m2.metric("✅ Đã Có Hình", f"{len(ma_da_thay)} Lô")
            col_m3.metric("❌ Thiếu Hình", f"{len(ma_thieu)} Lô", delta=f"{len(ma_thieu)} lô thiếu", delta_color="inverse")
            col_m4.metric("⚠️ Ảnh Nghi Ngờ / Lỗi", f"{len(hinh_nghi_ngo)} Ảnh", delta=f"{len(hinh_nghi_ngo)} lỗi", delta_color="inverse")

            # Chia 3 cột chi tiết
            st.markdown("<br>", unsafe_allow_html=True)
            col_c1, col_c2, col_c3 = st.columns(3, gap="medium")
            
            with col_c1:
                st.markdown("<div class='list-box'><h4 style='color:#1e7e34; margin-top:0;'>✅ CÁC MÃ ĐÃ CÓ HÌNH</h4>", unsafe_allow_html=True)
                if len(ma_da_thay) > 0:
                    for ma in sorted(list(ma_da_thay)):
                        st.markdown(f"- **{ma}**")
                else:
                    st.info("Chưa tìm thấy mã nào.")
                st.markdown("</div>", unsafe_allow_html=True)

            with col_c2:
                st.markdown("<div class='list-box'><h4 style='color:#dc3545; margin-top:0;'>❌ CÁC MÃ BỊ THIẾU</h4>", unsafe_allow_html=True)
                if len(ma_thieu) > 0:
                    for ma in sorted(list(ma_thieu)):
                        st.markdown(f"- **{ma}**")
                else:
                    st.success("Tuyệt vời! Đã chụp đủ 100% hàng.")
                st.markdown("</div>", unsafe_allow_html=True)

            with col_c3:
                st.markdown("<div class='list-box'><h4 style='color:#e37100; margin-top:0;'>⚠️ HÌNH ẢNH NGHI NGỜ / MỜ</h4>", unsafe_allow_html=True)
                if len(hinh_nghi_ngo) > 0:
                    st.warning("Các ảnh dưới đây không chứa mã nào nằm trong danh sách của tháng. Vui lòng kiểm tra lại tem!")
                    for item in hinh_nghi_ngo:
                        with st.expander(f"📷 {item['tên']}"):
                            st.image(item['hình'], use_container_width=True)
                            st.caption("AI đọc được các chữ sau (không khớp):")
                            st.write(item['chữ_nhận_diện'])
                else:
                    st.success("Tất cả hình tải lên đều nhận diện thành công.")
                st.markdown("</div>", unsafe_allow_html=True)

        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình phân tích: {e}")
