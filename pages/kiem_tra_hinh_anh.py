import streamlit as st
import pandas as pd
from PIL import Image
import io
import time
import base64
import requests
import gc

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

col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', type='primary', use_container_width=True):
        try:
            st.switch_page('app_ksnb.py')
        except: pass

st.markdown("<h1 class='apple-main-title'>📸 AI Kiểm Kê Bằng Hình Ảnh (Groq Vision)</h1>", unsafe_allow_html=True)
st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin-bottom: 20px;'>", unsafe_allow_html=True)

# -------------------------------------------------------------
# HÀM GỌI AI GROQ (NÉN MẠNH CHỐNG NGHẼN MẠNG)
# -------------------------------------------------------------
def parse_and_check_date(val, t_month, t_year):
    if pd.isna(val) or str(val).strip() == "": return False
    try:
        dt = pd.to_datetime(str(val), errors='coerce')
        if pd.notna(dt) and dt.month == t_month and dt.year == t_year: return True
    except: pass
    return False

def quet_anh_bang_groq(image_bytes, api_key, model_name, max_retries=3):
    """Gửi ảnh lên siêu máy chủ Groq - Phiên bản chống đứt cáp"""
    # 1. Ép dung lượng ảnh nhỏ lại mức (800x800) để gói tin siêu nhẹ
    img_goc = Image.open(io.BytesIO(image_bytes))
    img_goc.thumbnail((800, 800))
    buffered = io.BytesIO()
    img_goc.save(buffered, format="JPEG", quality=80)
    base64_image = base64.b64encode(buffered.getvalue()).decode('utf-8')
    
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key.strip()}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": model_name, 
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "Bạn là chuyên gia kiểm kê. Hãy đọc toàn bộ chữ, số, mã vạch có trên nhãn mác trong ảnh. Chỉ liệt kê các đoạn ký tự tìm được, không cần giải thích."},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                ]
            }
        ],
        "temperature": 0.1
    }
    
    # 3. Gửi và thử lại với Timeout cao (60s)
    for lan_thu in range(max_retries):
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=60)
            if response.status_code == 200:
                return response.json()['choices'][0]['message']['content']
            elif response.status_code == 429: # Quá tải
                time.sleep(10) # Nghỉ 10s cho máy chủ thở
                continue
            else:
                return f"LỖI TỪ GROQ ({response.status_code}): {response.text}"
        except Exception as e:
            if lan_thu == max_retries - 1: return f"LỖI MẠNG: {str(e)}"
            time.sleep(5)
    return "LỖI: Máy chủ Groq liên tục ngắt kết nối (Quá tải mạng)."

# -------------------------------------------------------------
# KHU VỰC GIAO DIỆN (TỰ ĐỘNG QUÉT MODEL CỦA GROQ)
# -------------------------------------------------------------
st.markdown("### 🔑 Cấu hình Trí Tuệ Nhân Tạo (Groq)")
api_key_input = st.text_input("Dán Groq API Key (bắt đầu bằng gsk_...) vào đây:", type="password")

selected_model = None
if api_key_input:
    with st.spinner("Đang kết nối vào tài khoản Groq để lấy danh sách AI..."):
        try:
            headers = {"Authorization": f"Bearer {api_key_input.strip()}"}
            resp = requests.get("https://api.groq.com/openai/v1/models", headers=headers, timeout=15)
            if resp.status_code == 200:
                models = resp.json().get("data", [])
                # Lọc ra các model chuyên đọc ảnh (Vision)
                danh_sach_model = [m["id"] for m in models if "vision" in m["id"].lower() or "3.2" in m["id"].lower()]
                
                if danh_sach_model:
                    st.success("✅ Kết nối máy chủ thành công!")
                    selected_model = st.selectbox("🤖 Vui lòng chọn Model AI (Không lo bị lỗi 404):", danh_sach_model)
                else:
                    st.warning("⚠️ Tài khoản của bạn hiện tại chưa được Groq cấp phát các model có khả năng đọc ảnh (Vision).")
            else:
                st.error(f"❌ Key không hợp lệ hoặc máy chủ đang bảo trì. Mã lỗi: {resp.status_code}")
        except Exception as e:
            st.error(f"❌ Không thể tải danh sách model do mạng nhà mạng chặn: {e}")

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
# LOGIC XỬ LÝ ẢNH CHỐNG ĐỨT CÁP
# -------------------------------------------------------------
if not f_excel or not f_images:
    st.info("👆 Hãy tải lên File Excel và Hình ảnh để AI bắt đầu quét.")
elif not api_key_input or not selected_model:
    st.warning("⚠️ Vui lòng cấu hình Key và chọn Model bên trên để bắt đầu.")
else:
    btn_quet = st.button("🚀 BẮT ĐẦU QUÉT KIỂM KÊ (DÙNG GROQ)", use_container_width=True)
    
    if btn_quet:
        progress_bar = st.progress(0, text="Đang bóc tách danh sách mã lô từ File Excel...")
        time.sleep(0.5)

        try:
            # 1. LỌC MÃ EXCEL
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

            # 2. QUÉT ẢNH BẰNG GROQ (VỚI CƠ CHẾ CHỜ)
            ma_da_thay = {} 
            hinh_nghi_ngo = [] 
            tong_hinh = len(f_images)
            
            for i, img_file in enumerate(f_images):
                ptram = int(20 + (80 * (i + 1) / tong_hinh))
                progress_bar.progress(ptram, text=f"🤖 Đang soi ảnh {i+1}/{tong_hinh}: {img_file.name}...")
                
                img_bytes = img_file.getvalue()
                
                # Gọi thẳng hàm AI tự tạo
                ocr_text = quet_anh_bang_groq(img_bytes, api_key_input, selected_model)
                
                if "LỖI" in ocr_text:
                    st.error(f"Lỗi ở ảnh {img_file.name}: {ocr_text}")
                    chuoi_hinh_anh = ""
                else:
                    chuoi_hinh_anh = ocr_text.upper().replace(" ", "").replace("-", "").replace("\n", "")
                
                # Bóp ảnh siêu nhỏ lưu RAM chống sập app
                image_hien_thi = Image.open(io.BytesIO(img_bytes))
                image_hien_thi.thumbnail((300, 300))
                
                tim_thay_trong_hinh = False
                for ma in ds_ma_can_tim:
                    ma_clean = ma.replace(" ", "").replace("-", "")
                    if ma_clean in chuoi_hinh_anh:
                        if ma not in ma_da_thay:
                            ma_da_thay[ma] = []
                        ma_da_thay[ma].append({'ten_file': img_file.name, 'hinh_anh': image_hien_thi.copy()})
                        tim_thay_trong_hinh = True
                
                if not tim_thay_trong_hinh and "LỖI" not in ocr_text:
                    hinh_nghi_ngo.append({'ten_file': img_file.name, 'hinh_anh': image_hien_thi.copy(), 'chu_nhan_dien': ocr_text})
                
                # Dọn RAM & Nghỉ nhịp 5 giây để tuân thủ luật chống Spam của máy chủ
                del image_hien_thi, img_bytes
                gc.collect()
                time.sleep(5)

            progress_bar.progress(100, text="✅ Hoàn tất phân tích bằng Groq AI!")
            time.sleep(0.5)
            progress_bar.empty()

            # 3. XUẤT BÁO CÁO
            ma_thieu = ds_ma_can_tim - set(ma_da_thay.keys())
            
            st.markdown("### 📊 BẢNG TỔNG KẾT KIỂM KÊ")
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("📋 Tổng Mã", f"{len(ds_ma_can_tim)} Lô")
            col_m2.metric("✅ Có Hình", f"{len(ma_da_thay)} Lô")
            col_m3.metric("❌ Thiếu Hình", f"{len(ma_thieu)} Lô", delta=f"{len(ma_thieu)} lô thiếu", delta_color="inverse")
            col_m4.metric("⚠️ Nghi Ngờ", f"{len(hinh_nghi_ngo)} Ảnh", delta=f"{len(hinh_nghi_ngo)} lỗi", delta_color="inverse")

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
                            st.caption("AI đọc được:")
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
                    bien_ban_data.append({'Mã Lô': ma, 'Trạng Thái': '✅ Đã kiểm kê', 'Ghi Chú': f"Ảnh: {file_names}"})
                else:
                    bien_ban_data.append({'Mã Lô': ma, 'Trạng Thái': '❌ THIẾU', 'Ghi Chú': "Không tìm thấy hình"})

            df_bb = pd.DataFrame(bien_ban_data)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_bb.to_excel(writer, index=False, sheet_name='Kiem_Ke')
                
            st.download_button(
                label="📥 Tải Xuống Biên Bản Kiểm Kê", 
                data=output.getvalue(), 
                file_name=f"Bien_Ban_Kiem_Ke_{chuoi_yymm}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Lỗi phân tích: {e}")
