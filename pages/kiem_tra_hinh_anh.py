import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image
import io

# Thư viện AI đọc chữ từ hình ảnh (cần chạy: pip install easyocr)
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
</style>
""", unsafe_allow_html=True)

# Nút quay về
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', use_container_width=True):
        st.switch_page('app_ksnb.py') # Trỏ về trang chủ của bạn

st.markdown("<h1 class='apple-main-title'>📸 Đối Chiếu Hình Ảnh & File Giao Nhận</h1>", unsafe_allow_html=True)
st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin-bottom: 30px;'>", unsafe_allow_html=True)

# -------------------------------------------------------------
# KHU VỰC TẢI DỮ LIỆU (CHIA 2 Ô RÕ RÀNG)
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
        with st.spinner("🤖 Trí tuệ nhân tạo đang đọc hình ảnh và dò tìm mã..."):
            try:
                # 1. Lấy toàn bộ mã từ File Giao Nhận
                df = pd.read_excel(f_excel)
                # Biến toàn bộ file Excel thành 1 list, xóa khoảng trắng, viết hoa
                ds_ma_excel = df.astype(str).apply(lambda x: x.str.strip().str.upper()).values.flatten()
                # Lọc bỏ các giá trị nhiễu hoặc nan, chỉ lấy các mã có ý nghĩa (độ dài > 4)
                ds_ma_excel = set([m for m in ds_ma_excel if m != 'NAN' and len(m) > 4])

                # 2. Đọc chữ từ Hình Ảnh bằng EasyOCR
                image = Image.open(f_image)
                image_np = np.array(image)
                
                # Khởi tạo bộ đọc chữ (Tiếng Việt & Tiếng Anh)
                reader = easyocr.Reader(['vi', 'en'], gpu=False)
                ocr_results = reader.readtext(image_np, detail=0) # detail=0 chỉ lấy list chữ
                
                # THỦ THUẬT: Ghép toàn bộ chữ đọc được thành 1 cục dài, xóa hết khoảng trắng
                chuoi_hinh_anh = "".join(ocr_results).upper().replace(" ", "").replace("-", "")

                # 3. Đối chiếu (Kiểm tra xem mã Excel có nằm trong chuỗi hình ảnh không)
                ma_tim_thay = []
                for ma in ds_ma_excel:
                    # Gỡ khoảng trắng của mã excel luôn cho an toàn
                    ma_clean = ma.replace(" ", "").replace("-", "")
                    if ma_clean in chuoi_hinh_anh:
                        ma_tim_thay.append(ma)

                # 4. TRẢ KẾT QUẢ RA MÀN HÌNH
                st.markdown("### 📋 KẾT QUẢ KIỂM TRA")
                col_res1, col_res2 = st.columns([1, 1.5])
                
                with col_res1:
                    st.image(image, caption="Hình ảnh thực tế", use_container_width=True)
                
                with col_res2:
                    if len(ma_tim_thay) > 0:
                        st.success("✅ HỢP LỆ: Đã tìm thấy mã trên ảnh khớp với File Giao nhận!")
                        st.write("🎯 **Các mã nhận diện được:**")
                        for m in ma_tim_thay:
                            st.markdown(f"- <b style='font-size:20px; color:#1e7e34;'>{m}</b>", unsafe_allow_html=True)
                    else:
                        st.error("🚨 KHÔNG HỢP LỆ: Mã trên hình ảnh KHÔNG TỒN TẠI trong file giao nhận!")
                        with st.expander("Bấm vào đây để xem máy tính đã đọc được những chữ gì từ ảnh (để kiểm tra xem máy đọc sai hay hàng sai thật):"):
                            st.write(ocr_results)

            except Exception as e:
                st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
                      
