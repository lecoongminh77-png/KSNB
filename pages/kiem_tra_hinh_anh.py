import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image
import io
import time

# Thư viện AI đọc chữ từ hình ảnh
try:
    import easyocr
except ImportError:
    st.error("⚠️ Bạn chưa cài thư viện EasyOCR. Hãy mở Terminal và gõ: pip install easyocr")

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
    
    div.stButton > button:first-child {
        background: #0071e3 !important; color: #ffffff !important;
        border-radius: 980px !important; border: none !important;
        padding: 10px 20px !important; font-weight: 600 !important;
    }
    div.stButton > button:first-child:hover { background: #0077ed !important; }
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff; border-radius: 16px; padding: 15px; border: 1px solid rgba(0,0,0,0.08); box-shadow: 0 2px 12px rgba(0,0,0,0.03);
    }
    .upload-title { font-size: 18px; font-weight: 700; margin-bottom: 10px; color: #0071e3;}
    
    /* Khung kết quả */
    .result-box-success {
        background-color: #e3f8e5; border-left: 5px solid #28a745; padding: 15px; border-radius: 8px; margin-top: 15px;
    }
    .result-box-error {
        background-color: #ffe6e6; border-left: 5px solid #dc3545; padding: 15px; border-radius: 8px; margin-top: 15px;
    }
</style>
""", unsafe_allow_html=True)

# Nút quay về
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', use_container_width=True):
        st.switch_page('app_ksnb.py') # Trỏ về file trang chủ của bạn

st.markdown("<h1 class='apple-main-title'>📸 Đối Chiếu Hình Ảnh & File Giao Nhận</h1>", unsafe_allow_html=True)
st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin-bottom: 30px;'>", unsafe_allow_html=True)

# -------------------------------------------------------------
# KHU VỰC TẢI DỮ LIỆU
# -------------------------------------------------------------
col_upload1, col_upload2 = st.columns(2, gap="large")

with col_upload1:
    with st.container(border=True):
        st.markdown("<div class='upload-title'>📄 1. Nạp File Giao Nhận (Excel)</div>", unsafe_allow_html=True)
        st.write("File chứa danh sách mã kiện (Hệ thống sẽ tự quét toàn bộ các cột trong file này).")
        f_excel = st.file_uploader("Tải file Excel", type=['xlsx', 'xls'], label_visibility="collapsed")

with col_upload2:
    with st.container(border=True):
        st.markdown("<div class='upload-title'>📸 2. Nạp Hình Ảnh Tem Nhãn</div>", unsafe_allow_html=True)
        st.write("Chụp hoặc tải lên hình ảnh tem dán trên cuộn/kiện hàng.")
        f_image = st.file_uploader("Tải hình ảnh", type=['jpg', 'jpeg', 'png'], label_visibility="collapsed")

# -------------------------------------------------------------
# LOGIC XỬ LÝ & ĐỐI CHIẾU
# -------------------------------------------------------------
if st.button('🔍 Bắt Đầu Quét & Đối Chiếu', use_container_width=True):
    if not f_excel or not f_image:
        st.warning("⚠️ Vui lòng tải lên ĐỦ 2 ô: File Giao Nhận và Hình Ảnh để đối chiếu!")
    else:
        # 1. Khởi tạo thanh tiến trình
        progress_bar = st.progress(0, text="Khởi động hệ thống nhận diện...")
        time.sleep(0.5)

        try:
            # 2. Đọc File Giao Nhận
            progress_bar.progress(20, text="Đang đọc và bóc tách dữ liệu từ File Excel...")
            df = pd.read_excel(f_excel)
            ds_ma_excel = df.astype(str).apply(lambda x: x.str.strip().str.upper()).values.flatten()
            ds_ma_excel = set([m for m in ds_ma_excel if m != 'NAN' and len(m) > 4])

            # 3. Đọc hình ảnh bằng AI
            progress_bar.progress(50, text="🤖 AI đang phân tích và nhận diện chữ trên hình ảnh...")
            image = Image.open(f_image)
            image_np = np.array(image)
            
            reader = easyocr.Reader(['vi', 'en'], gpu=False)
            ocr_results = reader.readtext(image_np, detail=0) 
            
            chuoi_hinh_anh = "".join(ocr_results).upper().replace(" ", "").replace("-", "")

            # 4. Đối chiếu
            progress_bar.progress(85, text="Đang dò tìm mã kiện trong cơ sở dữ liệu...")
            time.sleep(0.5)
            
            ma_tim_thay = []
            for ma in ds_ma_excel:
                ma_clean = ma.replace(" ", "").replace("-", "")
                if ma_clean in chuoi_hinh_anh:
                    ma_tim_thay.append(ma)

            # 5. Hoàn tất & Hiển thị
            progress_bar.progress(100, text="✅ Hoàn tất kiểm tra!")
            time.sleep(1)
            progress_bar.empty() # Ẩn thanh chạy đi cho đẹp

            # TRÌNH BÀY KẾT QUẢ RA MÀN HÌNH
            st.markdown("### 📋 KẾT QUẢ KIỂM TRA")
            col_res1, col_res2 = st.columns([1, 1.5])
            
            with col_res1:
                st.image(image, caption="Hình ảnh thực tế từ kho", use_container_width=True)
            
            with col_res2:
                if len(ma_tim_thay) > 0:
                    st.markdown(f"""
                    <div class="result-box-success">
                        <h4 style="color: #1e7e34; margin-top:0;">✅ HỢP LỆ</h4>
                        <p style="margin-bottom:5px;">Mã trên ảnh <b>CÓ TỒN TẠI</b> trong file giao nhận.</p>
                        Mã nhận diện được: <b>{", ".join(ma_tim_thay)}</b>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown("""
                    <div class="result-box-error">
                        <h4 style="color: #dc3545; margin-top:0;">🚨 CẢNH BÁO</h4>
                        <p style="margin-bottom:0;">Mã trên hình ảnh <b>KHÔNG TỒN TẠI</b> trong file giao nhận hoặc do ảnh quá mờ!</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    with st.expander("👁️ Xem chi tiết các chữ mà AI đọc được từ ảnh:"):
                        st.write(ocr_results)

        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
