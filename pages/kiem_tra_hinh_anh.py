import streamlit as st
import pandas as pd
from PIL import Image, ImageEnhance, ImageFilter
import io
import time
import pytesseract

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN & MÀU SẮC
# -------------------------------------------------------------
st.set_page_config(page_title='Kiểm Kê Bằng Hình Ảnh', page_icon='📸', layout='wide')

st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    .stApp { background-color: #f5f5f7; font-family: -apple-system, sans-serif; color: #1d1d1f; }
    .apple-main-title { font-size: 36px; font-weight: 700; color: #1d1d1f; margin-bottom: 20px; text-align: center; }
    [data-testid="stVerticalBlockBorderWrapper"] { background: #ffffff; border-radius: 16px; padding: 15px; border: 1px solid rgba(0,0,0,0.08); box-shadow: 0 2px 12px rgba(0,0,0,0.03); }
    div[data-testid="stMetric"] { background-color: #ffffff; border-radius: 12px; padding: 15px; border: 1px solid rgba(0,0,0,0.05); box-shadow: 0 2px 8px rgba(0,0,0,0.02); text-align: center; }
    .upload-title { font-size: 18px; font-weight: 700; margin-bottom: 10px; color: #0071e3;}
    .list-box { background: #fff; padding: 15px; border-radius: 10px; border: 1px solid #e2e8f0; height: 100%; }
    
    button[kind="primary"] {
        background-color: #007aff !important;
        color: white !important;
        border-color: #007aff !important;
    }
    button[kind="primary"]:hover {
        background-color: #0056b3 !important;
        border-color: #0056b3 !important;
    }
</style>
""", unsafe_allow_html=True)

# Nút quay về Trang chủ
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', type='primary', use_container_width=True):
        st.switch_page('app_ksnb.py')

st.markdown("<h1 class='apple-main-title'>📸 AI Kiểm Kê Bằng Hình Ảnh (Bản Siêu Nét)</h1>", unsafe_allow_html=True)
st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin-bottom: 20px;'>", unsafe_allow_html=True)

# -------------------------------------------------------------
# HÀM HỖ TRỢ
# -------------------------------------------------------------
def parse_and_check_date(val, t_month, t_year):
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

chuoi_yymm = f"{str(nam_chon)[-2:]}{thang_chon:02d}"
st.markdown("<br>", unsafe_allow_html=True)

col_upload1, col_upload2 = st.columns(2, gap="large")
with col_upload1:
    with st.container(border=True):
        st.markdown("<div class='upload-title'>📄 2. Nạp File Excel (Kế hoạch)</div>", unsafe_allow_html=True)
        f_excel = st.file_uploader("Tải file Excel", type=['xlsx', 'xls'], label_visibility="collapsed")

with col_upload2:
    with st.container(border=True):
        st.markdown("<div class='upload-title'>📸 3. Nạp Hình Ảnh (Thực tế)</div>", unsafe_allow_html=True)
        f_images = st.file_uploader("Tải hình ảnh", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True, label_visibility="collapsed")

# -------------------------------------------------------------
# LOGIC XỬ LÝ (TESSERACT ĐƯỢC TĂNG CƯỜNG THỊ GIÁC)
# -------------------------------------------------------------
if not f_excel or not f_images:
    st.info("👆 Hãy tải lên File Excel và Hình ảnh để quét mã vạch.")
else:
    progress_bar = st.progress(0, text="Đang bóc tách danh sách mã lô từ File Excel...")
    time.sleep(0.5)

    try:
        # BƯỚC 1: LỌC DANH SÁCH MÃ TỪ EXCEL
        ds_ma_can_tim = set()
        xls = pd.ExcelFile(f_excel)
        sheet_names = xls.sheet_names
        
        if 'GIAO NHAN TP GCBB' in sheet_names:
            df1 = pd.read_excel(xls, sheet_name='GIAO NHAN TP GCBB', header=None, dtype=str)
            for idx, row in df1.iterrows():
                if parse_and_check_date(row[1], thang_chon, nam_chon):
                    ma = str(row[9]).strip().upper()
                    if ma and ma != 'NAN' and len(ma) > 3: ds_ma_can_tim.add(ma)

        sheet2 = next((s for s in sheet_names if 'NXT NGUYEN LIEU' in s.upper()), None)
        if sheet2:
            df2 = pd.read_excel(xls, sheet_name=sheet2, header=None, dtype=str)
            for idx, row in df2.iterrows():
                if parse_and_check_date(row[13], thang_chon, nam_chon):
                    ma = str(row[1]).strip().upper()
                    if ma and ma != 'NAN' and len(ma) > 3: ds_ma_can_tim.add(ma)

        if len(ds_ma_can_tim) == 0:
            progress_bar.empty()
            st.error(f"🚨 Không tìm thấy mã lô nào thuộc tháng {thang_chon}/{nam_chon} trong file Excel!")
            st.stop()

        # BƯỚC 2: QUÉT ẢNH VỚI TESSERACT "ĐỘ"
        ma_da_thay = {} 
        hinh_nghi_ngo = [] 
        
        tong_hinh = len(f_images)
        for i, img_file in enumerate(f_images):
            ptram = int(20 + (80 * (i + 1) / tong_hinh))
            progress_bar.progress(ptram, text=f"🔍 Đang rà soát tấm ảnh {i+1}/{tong_hinh}: {img_file.name}...")
            
            # --- BỘ LỌC XỬ LÝ ẢNH TRƯỚC KHI ĐỌC ---
            image_goc = Image.open(img_file)
            
            # 1. Phóng to ảnh x2 (giúp chữ to rõ hơn)
            w, h = image_goc.size
            img_to = image_goc.resize((w * 2, h * 2), Image.Resampling.LANCZOS)
            
            # 2. Chuyển sang ảnh Trắng Đen
            img_gray = img_to.convert('L')
            
            # 3. Kéo Max Độ tương phản (chữ đen nền trắng)
            enhancer = ImageEnhance.Contrast(img_gray)
            img_contrast = enhancer.enhance(2.5) 
            
            # 4. Làm nét chữ
            sharpener = ImageEnhance.Sharpness(img_contrast)
            img_ready = sharpener.enhance(2.0)
            
            try:
                # Ép Tesseract càn quét mọi khối chữ lộn xộn trên ảnh (--psm 11)
                custom_config = r'--oem 3 --psm 11'
                ocr_text = pytesseract.image_to_string(img_ready, config=custom_config)
                
                chuoi_hinh_anh = ocr_text.upper().replace(" ", "").replace("-", "").replace("\n", "")
                
                tim_thay_trong_hinh = False
                for ma in ds_ma_can_tim:
                    ma_clean = ma.replace(" ", "").replace("-", "")
                    if ma_clean in chuoi_hinh_anh:
                        if ma not in ma_da_thay:
                            ma_da_thay[ma] = []
                        ma_da_thay[ma].append({'ten_file': img_file.name, 'hinh_anh': image_goc})
                        tim_thay_trong_hinh = True
                
                if not tim_thay_trong_hinh:
                    hinh_nghi_ngo.append({'ten_file': img_file.name, 'hinh_anh': image_goc, 'chu_nhan_dien': ocr_text})
                
            except Exception as e:
                st.error(f"Lỗi phân tích ở ảnh {img_file.name}: {e}")

        progress_bar.progress(100, text="✅ Hoàn tất phân tích!")
        time.sleep(0.5)
        progress_bar.empty()

        # BƯỚC 3: BÁO CÁO UI & XUẤT EXCEL
        ma_thieu = ds_ma_can_tim - set(ma_da_thay.keys())
        
        st.markdown("### 📊 BẢNG TỔNG KẾT KIỂM KÊ")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("📋 Tổng Mã Cần Kiểm", f"{len(ds_ma_can_tim)} Lô")
        col_m2.metric("✅ Đã Có Hình", f"{len(ma_da_thay)} Lô")
        col_m3.metric("❌ Thiếu Hình", f"{len(ma_thieu)} Lô", delta=f"{len(ma_thieu)} lô thiếu", delta_color="inverse")
        col_m4.metric("⚠️ Ảnh Nghi Ngờ / Lỗi", f"{len(hinh_nghi_ngo)} Ảnh", delta=f"{len(hinh_nghi_ngo)} lỗi", delta_color="inverse")

        st.markdown("<br>", unsafe_allow_html=True)
        col_c1, col_c2, col_c3 = st.columns(3, gap="medium")
        
        with col_c1:
            st.markdown("<div class='list-box'><h4 style='color:#1e7e34; margin-top:0;'>✅ CÁC MÃ ĐÃ CÓ HÌNH</h4>", unsafe_allow_html=True)
            if len(ma_da_thay) > 0:
                for ma in sorted(ma_da_thay.keys()):
                    with st.expander(f"📦 {ma}"):
                        for img_item in ma_da_thay[ma]:
                            st.image(img_item['hinh_anh'], caption=f"📄 {img_item['ten_file']}", use_container_width=True)
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
            st.markdown("<div class='list-box'><h4 style='color:#e37100; margin-top:0;'>⚠️ HÌNH ẢNH NGHI NGỜ</h4>", unsafe_allow_html=True)
            if len(hinh_nghi_ngo) > 0:
                for item in hinh_nghi_ngo:
                    with st.expander(f"📷 {item['ten_file']}"):
                        st.image(item['hinh_anh'], use_container_width=True)
                        st.caption("🔍 Máy đọc được các chữ sau (nhưng không khớp mã Excel):")
                        st.write(item['chu_nhan_dien'])
            else:
                st.success("Tất cả hình tải lên đều nhận diện thành công.")
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.markdown("### 📝 Biên Bản Kiểm Kê (Xuất Excel)")
        
        bien_ban_data = []
        for ma in sorted(ds_ma_can_tim):
            if ma in ma_da_thay:
                file_names = ", ".join([img['ten_file'] for img in ma_da_thay[ma]])
                bien_ban_data.append({'Mã Lô Cần Kiểm Kê': ma, 'Trạng Thái': '✅ Đã kiểm kê (Có hình)', 'Ghi Chú': f"Nằm trong ảnh: {file_names}"})
            else:
                bien_ban_data.append({'Mã Lô Cần Kiểm Kê': ma, 'Trạng Thái': '❌ THIẾU', 'Ghi Chú': "Không tìm thấy hình ảnh minh chứng"})

        df_bb = pd.DataFrame(bien_ban_data)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_bb.to_excel(writer, index=False, sheet_name='Kiem_Ke_Hinh_Anh')
            
        st.download_button(
            label="📥 Tải Xuống Biên Bản Kiểm Kê Tháng", 
            data=output.getvalue(), 
            file_name=f"Bien_Ban_Kiem_Ke_Hinh_Anh_{chuoi_yymm}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Đã xảy ra lỗi trong quá trình phân tích: {e}")
