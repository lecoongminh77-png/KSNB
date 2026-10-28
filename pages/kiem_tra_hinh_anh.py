import streamlit as st
import pandas as pd
from PIL import Image
import io
import time
import gc
import re
from google import genai
from google.genai import types

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
</style>
""", unsafe_allow_html=True)

col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', type='primary', use_container_width=True):
        try: st.switch_page('app_ksnb.py')
        except: pass

st.markdown("<h1 class='apple-main-title'>📸 AI Kiểm Kê Hình Ảnh (Bản Gộp Siêu Tốc)</h1>", unsafe_allow_html=True)
st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin-bottom: 20px;'>", unsafe_allow_html=True)

def parse_and_check_date(val, t_month, t_year):
    if pd.isna(val) or str(val).strip() == "": return False
    try:
        dt = pd.to_datetime(str(val), errors='coerce')
        if pd.notna(dt) and dt.month == t_month and dt.year == t_year: return True
    except: pass
    return False

# -------------------------------------------------------------
# CẤU HÌNH GOOGLE GEMINI (ĐÚNG CHUẨN SFORUM)
# -------------------------------------------------------------
st.markdown("### 🔑 Cấu hình Trí Tuệ Nhân Tạo (Google Gemini)")
secret_key = ""
try: secret_key = st.secrets.get("GEMINI_API_KEY", "")
except: pass

api_key_input = st.text_input("Dán Google Gemini API Key vào đây:", value=secret_key, type="password")

selected_model = None
client = None

if api_key_input:
    try:
        client = genai.Client(api_key=api_key_input.strip())
        with st.spinner("Đang kết nối lấy danh sách AI..."):
            danh_sach_model = []
            for m in client.models.list():
                if "gemini" in m.name:
                    ten_gọn = m.name.replace("models/", "")
                    if ten_gọn not in danh_sach_model:
                        danh_sach_model.append(ten_gọn)
            if danh_sach_model:
                st.success("✅ Kết nối Google thành công!")
                if "gemini-1.5-flash" in danh_sach_model:
                    danh_sach_model.insert(0, danh_sach_model.pop(danh_sach_model.index("gemini-1.5-flash")))
                selected_model = st.selectbox("🤖 Chọn Model (Nên dùng gemini-1.5-flash cho mượt):", danh_sach_model)
            else:
                st.error("🚨 Không lấy được model từ Google.")
    except Exception as e:
        st.error(f"❌ Lỗi kết nối Key: {e}")

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
# LOGIC GỘP 1 LẦN - TRIỆT TIÊU LỖI 503/429
# -------------------------------------------------------------
if not f_excel or not f_images:
    st.info("👆 Hãy tải lên File Excel và Hình ảnh để AI bắt đầu quét.")
elif not api_key_input or not selected_model:
    st.warning("⚠️ Vui lòng cấu hình Key và chọn Model bên trên để bắt đầu.")
else:
    btn_quet = st.button("🚀 BẮT ĐẦU QUÉT HÀNG LOẠT (KHÔNG SỢ CHẶN)", use_container_width=True)
    
    if btn_quet:
        progress_bar = st.progress(0, text="Đang bóc tách mã từ Excel...")
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

            # 2. ĐÓNG GÓI TẤT CẢ ẢNH VÀO 1 LỆNH (ONE-SHOT)
            progress_bar.progress(30, text=f"Đang nén và đóng gói {len(f_images)} bức ảnh thành 1 hộp...")
            
            contents = [
                "Hãy đọc chữ và số lô trên các nhãn mác trong các bức ảnh dưới đây. "
                "Tôi đã ghi chú 'TÊN FILE:' trước mỗi ảnh. Bạn phải trả lời ĐÚNG cấu trúc sau:\n\n"
                "TÊN FILE: <tên file>\nNỘI DUNG: <các mã đọc được>\n---\n"
            ]
            
            dict_hinh_anh = {}
            
            for img_file in f_images:
                img_bytes = img_file.getvalue()
                img = Image.open(io.BytesIO(img_bytes))
                
                # Ảnh nén gửi AI (nhẹ)
                img_api = img.copy()
                img_api.thumbnail((800, 800))
                
                # Ảnh nén cho giao diện
                img_ui = img.copy()
                img_ui.thumbnail((300, 300))
                dict_hinh_anh[img_file.name] = img_ui
                
                # Nhét vào thùng hàng gửi AI
                contents.append(f"TÊN FILE: {img_file.name}")
                contents.append(img_api)

            # 3. GỌI GOOGLE ĐÚNG 1 LẦN (Né 100% lỗi spam/quá tải)
            progress_bar.progress(60, text="Đang gửi gói hàng khổng lồ lên AI Google phân tích (Chỉ gọi 1 lần duy nhất)...")
            
            response = client.models.generate_content(
                model=selected_model,
                contents=contents,
                config=types.GenerateContentConfig(temperature=0.1)
            )
            
            raw_text = response.text if response.text else ""
            
            progress_bar.progress(90, text="Đang khớp dữ liệu AI trả về với bảng Excel...")
            
            # 4. TÁCH KẾT QUẢ AI TRẢ VỀ
            ocr_results = {}
            blocks = raw_text.split("TÊN FILE:")
            for block in blocks:
                if "NỘI DUNG:" in block:
                    parts = block.split("NỘI DUNG:")
                    fname = parts[0].strip()
                    content = parts[1].split("---")[0].strip().upper().replace(" ", "").replace("-", "").replace("\n", "")
                    ocr_results[fname] = content

            ma_da_thay = {} 
            hinh_nghi_ngo = [] 
            
            for img_file in f_images:
                fname = img_file.name
                # Lấy text AI đọc được cho file này (nếu AI đọc sót thì chuỗi rỗng)
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
            
            del contents
            gc.collect()

            progress_bar.progress(100, text="✅ Hoàn tất!")
            time.sleep(0.5)
            progress_bar.empty()

            # 5. XUẤT BÁO CÁO NHƯ CŨ
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
