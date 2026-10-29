import streamlit as st
import pandas as pd
from PIL import Image
import io
import time
import gc
import base64
from groq import Groq

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN
# -------------------------------------------------------------
st.set_page_config(page_title='Kiểm Kê Kho (Llama 3.2 Chính thức)', page_icon='📸', layout='wide')

st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    .stApp { background-color: #f5f5f7; font-family: -apple-system, sans-serif; color: #1d1d1f; }
    .apple-main-title { font-size: 36px; font-weight: 700; color: #1d1d1f; margin-bottom: 20px; text-align: center; }
    [data-testid="stVerticalBlockBorderWrapper"] { background: #ffffff; border-radius: 16px; padding: 15px; border: 1px solid rgba(0,0,0,0.08); box-shadow: 0 2px 12px rgba(0,0,0,0.03); }
    div[data-testid="stMetric"] { background-color: #ffffff; border-radius: 12px; padding: 15px; border: 1px solid rgba(0,0,0,0.05); box-shadow: 0 2px 8px rgba(0,0,0,0.02); text-align: center; }
    .upload-title { font-size: 18px; font-weight: 700; margin-bottom: 10px; color: #f97316;}
    .list-box { background: #fff; padding: 15px; border-radius: 10px; border: 1px solid #e2e8f0; height: 100%; }
    button[kind="primary"] { background-color: #f97316 !important; color: white !important; border-color: #f97316 !important; }
</style>
""", unsafe_allow_html=True)

col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', type='primary', use_container_width=True):
        try: st.switch_page('app_ksnb.py')
        except: pass

st.markdown("<h1 class='apple-main-title'>📸 AI Kiểm Kê Hình Ảnh (Sức mạnh Llama 3.2 Vision)</h1>", unsafe_allow_html=True)
st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin-bottom: 20px;'>", unsafe_allow_html=True)

# Hàm hỗ trợ ép ngày tháng
def parse_and_check_date(val, t_month, t_year):
    if pd.isna(val) or str(val).strip() == "": return False
    try:
        dt = pd.to_datetime(str(val), errors='coerce')
        if pd.notna(dt) and dt.month == t_month and dt.year == t_year: return True
    except: pass
    return False

# Hàm chuyển ảnh sang Base64 để gửi qua Groq
def encode_image_to_base64(image_bytes):
    img = Image.open(io.BytesIO(image_bytes))
    # Resize nhẹ để tránh vượt quá dung lượng payload của API
    img.thumbnail((800, 800))
    buffered = io.BytesIO()
    img.save(buffered, format="JPEG", quality=85)
    return base64.b64encode(buffered.getvalue()).decode('utf-8')

# -------------------------------------------------------------
# CẤU HÌNH GROQ API
# -------------------------------------------------------------
st.markdown("### 🔑 Cấu hình Groq API (Miễn phí)")
api_key_input = st.text_input("Dán Groq API Key vào đây:", type="password")

# Cập nhật danh sách model chính thức (đã gỡ bỏ bản preview bị khai tử)
selected_model = st.selectbox(
    "🤖 Chọn Model Vision:",
    [
        "llama-3.2-90b-vision-instruct",  # Model 90B chính thức
        "llama-3.2-11b-vision-instruct",  # Model 11B chính thức
        "llama-3.2-90b-text-preview"      # Dự phòng
    ]
)

client = None
if api_key_input:
    try:
        client = Groq(api_key=api_key_input.strip())
        st.success("✅ Đã kết nối thành công với máy chủ Groq!")
    except Exception as e:
        st.error(f"❌ Lỗi khởi tạo Key: {e}")

st.markdown("<hr style='border-top: 1px dashed #ccc; margin: 20px 0;'>", unsafe_allow_html=True)

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
        f_images = st.file_uploader("Tải nhiều hình ảnh", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True, label_visibility="collapsed")

# -------------------------------------------------------------
# LOGIC GỌI LLAMA 3.2 VISION QUA GROQ
# -------------------------------------------------------------
if not f_excel or not f_images:
    st.info("👆 Hãy tải lên File Excel và Hình ảnh để AI bắt đầu quét.")
elif not api_key_input or not client:
    st.warning("⚠️ Vui lòng nhập Groq API Key để bắt đầu.")
else:
    btn_quet = st.button("🚀 BẮT ĐẦU KIỂM KÊ (Dùng Llama 3.2)", use_container_width=True)
    
    if btn_quet:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # 1. LỌC MÃ TỪ EXCEL
            status_text.markdown("**Đang bóc tách mã từ Excel...**")
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
                st.error("🚨 Không tìm thấy mã lô nào trong Excel!")
                st.stop()

            # 2. XỬ LÝ ẢNH & GỌI API (Nhóm 10 ảnh)
            chunk_size = 10 
            image_chunks = [f_images[i:i + chunk_size] for i in range(0, len(f_images), chunk_size)]
            
            ocr_results = {}
            dict_hinh_anh = {}
            total_chunks = len(image_chunks)
            
            for chunk_idx, chunk in enumerate(image_chunks):
                status_text.markdown(f"**🧠 Llama 3.2 đang quét nhóm ảnh {chunk_idx + 1}/{total_chunks}...**")
                progress_bar.progress(int(((chunk_idx+1) / total_chunks) * 85))
                
                # Xây dựng cấu trúc tin nhắn theo chuẩn OpenAI/Groq Vision
                prompt_text = (
                    "Hãy đọc tất cả chữ và số lô trên nhãn mác trong các bức ảnh. "
                    "Tôi sẽ đánh dấu từng ảnh bằng chữ [ẢNH SỐ X]. "
                    "Bạn phải trả lời ĐÚNG cấu trúc:\n\n"
                    "[ẢNH SỐ X]\nNỘI DUNG: <mã đọc được>\n---\n"
                )
                
                content_payload = [{"type": "text", "text": prompt_text}]
                
                for i, img_file in enumerate(chunk):
                    img_bytes = img_file.getvalue()
                    
                    # Nén ảnh UI
                    img_ui = Image.open(io.BytesIO(img_bytes)).copy()
                    img_ui.thumbnail((300, 300))
                    dict_hinh_anh[img_file.name] = img_ui
                    
                    # Nén và Encode ảnh gửi API
                    base64_image = encode_image_to_base64(img_bytes)
                    
                    content_payload.append({"type": "text", "text": f"\n\n[ẢNH SỐ {i+1}] (Tên file: {img_file.name})"})
                    content_payload.append({
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}
                    })
                
                # Gửi yêu cầu lên Groq (Cơ chế tự thử lại)
                max_retries = 3
                for attempt in range(max_retries):
                    try:
                        completion = client.chat.completions.create(
                            model=selected_model,
                            messages=[{"role": "user", "content": content_payload}],
                            temperature=0.1,
                            max_tokens=2048
                        )
                        raw_text = completion.choices[0].message.content
                        
                        # Bóc tách kết quả dựa trên tên file
                        for i, img_file in enumerate(chunk):
                            marker = f"[ẢNH SỐ {i+1}]"
                            if marker in raw_text:
                                block = raw_text.split(marker)[1]
                                if "NỘI DUNG:" in block:
                                    content = block.split("NỘI DUNG:")[1].split("---")[0].strip().upper().replace(" ", "").replace("-", "").replace("\n", "")
                                    ocr_results[img_file.name] = content
                                else:
                                    ocr_results[img_file.name] = ""
                            else:
                                ocr_results[img_file.name] = ""
                        
                        break # Thành công thì thoát vòng lặp retry
                    
                    except Exception as e:
                        if attempt < max_retries - 1:
                            status_text.warning(f"⚠️ Đang thử lại lần {attempt + 2} cho nhóm {chunk_idx + 1}...")
                            time.sleep(3)
                        else:
                            st.error(f"❌ Nhóm {chunk_idx + 1} gặp lỗi: {e}")

                # Nghỉ ngắn giữa các nhóm
                if chunk_idx < total_chunks - 1:
                    time.sleep(2)

            # 3. ĐỐI SOÁT DỮ LIỆU EXCEL
            status_text.markdown("**🔄 Đang đối soát dữ liệu với bảng kế hoạch...**")
            progress_bar.progress(95)
            
            ma_da_thay = {} 
            hinh_nghi_ngo = [] 
            
            for img_file in f_images:
                fname = img_file.name
                chuoi_hinh_anh = ocr_results.get(fname, "")
                
                tim_thay = False
                for ma in ds_ma_can_tim:
                    ma_clean = ma.replace(" ", "").replace("-", "")
                    if ma_clean in chuoi_hinh_anh:
                        if ma not in ma_da_thay:
                            ma_da_thay[ma] = []
                        ma_da_thay[ma].append({'ten_file': fname, 'hinh_anh': dict_hinh_anh[fname]})
                        tim_thay = True
                
                if not tim_thay:
                    hinh_nghi_ngo.append({'ten_file': fname, 'hinh_anh': dict_hinh_anh[fname], 'chu_nhan_dien': chuoi_hinh_anh})

            gc.collect()
            progress_bar.progress(100)
            status_text.success("✅ Hoàn tất toàn bộ quy trình kiểm kê!")
            time.sleep(1)
            progress_bar.empty()
            status_text.empty()

            # 4. XUẤT BÁO CÁO VÀ GIAO DIỆN
            ma_thieu = ds_ma_can_tim - set(ma_da_thay.keys())
            
            st.markdown("### 📊 BẢNG TỔNG KẾT KIỂM KÊ")
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("📋 Tổng Mã Excel", f"{len(ds_ma_can_tim)} Lô")
            col_m2.metric("✅ Có Hình", f"{len(ma_da_thay)} Lô")
            col_m3.metric("❌ Thiếu Hình", f"{len(ma_thieu)} Lô", delta=f"{len(ma_thieu)} lô thiếu", delta_color="inverse")
            col_m4.metric("⚠️ Nghi Ngờ", f"{len(hinh_nghi_ngo)} Ảnh", delta=f"{len(hinh_nghi_ngo)} ảnh", delta_color="inverse")

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
                    st.success("Tuyệt vời! Đã chụp đủ 100%
