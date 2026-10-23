import sys
import os
import requests 

# 1. Ép Python nhận diện thư mục gốc
root_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root_path not in sys.path:
    sys.path.insert(0, root_path)

# 2. Khai báo thư viện
import io
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import streamlit as st
from streamlit_lottie import st_lottie

# 3. Import module nội bộ
from ai_chat import hien_thi_tro_ly_ai

def load_lottieurl(url: str):
    try:
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None

lottie_hoat_anh = load_lottieurl("https://lottie.host/654b9803-9ee2-4c1b-91e3-3f074d4f79d3/b4IPc6jdqR.json")

st.set_page_config(page_title="Kiểm Tra Tồn Kho Cuộn", page_icon="📦", layout="wide")

st.markdown("""
<style>
    [data-testid="stSidebarNav"] { display: none !important; }
    .stApp { background-color: #f5f5f7; font-family: -apple-system, sans-serif; color: #1d1d1f; }
    .apple-header-box { display: flex; justify-content: space-between; align-items: flex-end; padding: 20px 0 10px 0; margin-bottom: 5px; }
    .apple-main-title { font-size: 38px; font-weight: 700; letter-spacing: -0.025em; margin: 0; }
    .apple-sub-text { font-size: 16px; font-weight: 600; text-align: right; margin: 0; line-height: 1.35; }
    .apple-sub-text span { color: #0071e3; font-size: 14px; display: block; margin-top: 2px; font-weight: 500; }
    div.stButton > button { background: #0071e3 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; padding: 10px 24px !important; font-weight: 600 !important; font-size: 15px !important; }
</style>
""", unsafe_allow_html=True)

col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Trang chủ", use_container_width=True):
        st.switch_page("app_ksnb.py")

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)

st.markdown("""
<div class="apple-header-box">
    <div><h1 class="apple-main-title">📦 Kiểm Tra Mẫu Tồn Kho Cuộn </h1></div>
    <div><div class="apple-sub-text">Tự động điền công thức & đối soát trùng lặp 2 file siêu tốc.<span>Lê Công Minh Workspace ›</span></div></div>
</div>
<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>
""", unsafe_allow_html=True)

MA_TRAN_CHUAN = {
    "Dưới 2 tấn":    {"khong_loi": 20, "loi_giay": 40,  "loi_sat": 70},
    "Từ 2 - 5 tấn":  {"khong_loi": 40, "loi_giay": 60,  "loi_sat": 90},
    "Từ 5 - 7 tấn":  {"khong_loi": 50, "loi_giay": 70,  "loi_sat": 100},
    "Từ 7 - 10 tấn": {"khong_loi": 60, "loi_giay": 90,  "loi_sat": 120},
    "Từ 10 - 12 tấn":{"khong_loi": 70, "loi_giay": 100, "loi_sat": 120},
    "Từ 12 - 17 tấn":{"khong_loi": 80, "loi_giay": 110, "loi_sat": 130},
    "Từ 17 - 23 tấn":{"khong_loi": 90, "loi_giay": 120, "loi_sat": 140},
    "Trên 23 tấn":   {"khong_loi": 90, "loi_giay": 120, "loi_sat": 140},
}

def xac_dinh_hang(kl):
    try: kl = float(kl)
    except: return ""
    if kl < 2000: return "Dưới 2 tấn"
    elif kl <= 5000: return "Từ 2 - 5 tấn"
    elif kl <= 7000: return "Từ 5 - 7 tấn"
    elif kl <= 10000: return "Từ 7 - 10 tấn"
    elif kl <= 12000: return "Từ 10 - 12 tấn"
    elif kl <= 17000: return "Từ 12 - 17 tấn"
    elif kl <= 23000: return "Từ 17 - 23 tấn"
    else: return "Trên 23 tấn"

def xac_dinh_ky_chuan(hang, loai_loi):
    loai_str = str(loai_loi).lower().strip() if pd.notna(loai_loi) else ""
    nhom = MA_TRAN_CHUAN.get(hang, {"khong_loi": 0, "loi_giay": 0, "loi_sat": 0})
    if "sắt" in loai_str or "sat" in loai_str: return nhom["loi_sat"]
    elif "giấy" in loai_str or "giay" in loai_str: return nhom["loi_giay"]
    else: return nhom["khong_loi"]

col_f1, col_f2 = st.columns(2)
with col_f1: f_tonkho = st.file_uploader("1. File Báo cáo Tồn Kho (Mẫu 2356)", type=["xlsx", "xls", "xlsb"])
with col_f2: f_doichieu = st.file_uploader("2. File Dữ liệu Đối Chiếu (OPM 105)", type=["xlsx", "xls", "xlsb"])

if f_tonkho is not None:
    if st.button("🚀 Bắt đầu tính toán tự động", use_container_width=True):
        with st.spinner("Đang quét cấu trúc ẩn và điền trực tiếp vào Form..."):
            try:
                # 1. Đọc File Đối Chiếu OPM 105
                set_ma_doi_chieu = set()
                if f_doichieu is not None:
                    try:
                        df_ref = pd.read_excel(f_doichieu)
                        # Quét cột mã bên OPM (Ưu tiên cột F)
                        col_ref_ma = df_ref.columns[5] if len(df_ref.columns) > 5 else df_ref.columns[0]
                        set_ma_doi_chieu = set(df_ref[col_ref_ma].astype(str).str.replace(r"\.0$", "", regex=True).str.strip().str.upper())
                    except:
                        pass

                # 2. Xử lý File Tồn kho bằng thư viện OpenPyXL để BẢO TOÀN 100% CẤU TRÚC GỐC
                wb = openpyxl.load_workbook(f_tonkho)
                ws = wb.active

                # Bộ quét thông minh: Tự tìm vị trí các cột bạn đã tạo sẵn trong File
                col_map = {}
                max_header_row = 1
                
                for r in range(1, 26):
                    for c in range(1, ws.max_column + 5):
                        cell_val = ws.cell(row=r, column=c).value
                        if not cell_val: continue
                        val = str(cell_val).lower().strip()
                        
                        # Cột Dữ Liệu Nguồn
                        if "mã lô" in val or "cuộn" in val: col_map['ma_lo'] = c; max_header_row = max(max_header_row, r)
                        elif "gross" in val: col_map['gross'] = c; max_header_row = max(max_header_row, r)
                        elif val == "net" or "net " in val: col_map['net'] = c; max_header_row = max(max_header_row, r)
                        elif "loại lõi" in val or val == "lõi": col_map['loi'] = c
                        
                        # Cột Công Thức (Hệ thống sẽ điền thẳng vào cột L nếu bạn đặt tên nó là Số ký chuẩn)
                        elif "trùng" in val: col_map['trung'] = c
                        elif val == "hạng": col_map['hang'] = c
                        elif "chênh lệch" in val or "lệch" in val: col_map['lech'] = c
                        elif "ký chuẩn" in val or "kg chuẩn" in val or "chuẩn" in val: col_map['chuan'] = c
                        elif val == "loại": col_map['loai'] = c
                        elif "dây chuyền" in val: col_map['day_chuyen'] = c
                        elif val == "số kg" or val == "số lượng kg": col_map['so_kg'] = c

                # 3. Tạo cột mới Ở CUỐI BẢNG nếu bạn lỡ đưa file thô (Chống dư cột ở A-G)
                next_col = ws.max_column + 1
                for key, name in [
                    ('trung', 'Trùng'), ('hang', 'Hạng'), ('so_kg', 'Số kg'), 
                    ('lech', 'Ký chênh lệch'), ('chuan', 'Số ký chuẩn'), 
                    ('loai', 'Loại'), ('day_chuyen', 'Dây chuyền')
                ]:
                    if key not in col_map:
                        ws.cell(row=max_header_row, column=next_col).value = name
                        col_map[key] = next_col
                        next_col += 1

                # 4. Bắt đầu tính toán
                start_row = max_header_row + 1
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

                for r in range(start_row, ws.max_row + 1):
                    ma_lo_val = str(ws.cell(row=r, column=col_map.get('ma_lo', 1)).value or "").replace(".0", "").strip()
                    gross_val = ws.cell(row=r, column=col_map.get('gross', 1)).value
                    net_val = ws.cell(row=r, column=col_map.get('net', 1)).value
                    loi_val = str(ws.cell(row=r, column=col_map.get('loi', 1)).value or "").strip()
                    
                    try: gross_num = float(gross_val)
                    except: gross_num = 0
                    try: net_num = float(net_val)
                    except: net_num = 0
                    
                    # Bỏ qua dòng trống
                    if not ma_lo_val and gross_num == 0 and net_num == 0:
                        continue
                        
                    # Áp dụng Công Thức
                    day_chuyen = ma_lo_val[3:6] if len(ma_lo_val) >= 6 else ""
                    hang = xac_dinh_hang(gross_num)
                    ky_chuan = xac_dinh_ky_chuan(hang, loi_val)
                    so_kg = gross_num - net_num
                    lech = so_kg - ky_chuan
                    loai_str = loi_val[0] if len(loi_val) > 0 else ""
                    
                    is_trung = "Không"
                    if set_ma_doi_chieu and ma_lo_val.upper() in set_ma_doi_chieu:
                        is_trung = "Có"
                        
                    # GHI KẾT QUẢ VÀO ĐÚNG TỌA ĐỘ BẠN ĐÃ TẠO SẴN TRONG FILE
                    ws.cell(row=r, column=col_map['trung']).value = is_trung
                    ws.cell(row=r, column=col_map['hang']).value = hang
                    ws.cell(row=r, column=col_map['so_kg']).value = so_kg
                    ws.cell(row=r, column=col_map['lech']).value = lech
                    ws.cell(row=r, column=col_map['chuan']).value = ky_chuan
                    ws.cell(row=r, column=col_map['loai']).value = loai_str
                    ws.cell(row=r, column=col_map['day_chuyen']).value = day_chuyen
                    
                    # Logic Tô vàng nguyên dòng nếu Vượt chuẩn & Không trùng
                    if is_trung == "Không" and lech > 0:
                        for c_idx in range(1, ws.max_column + 1):
                            cell = ws.cell(row=r, column=c_idx)
                            cell.fill = yellow_fill

                # 5. Đóng gói file trả về
                output = io.BytesIO()
                wb.save(output)

                st.success(f"✅ Xử lý thành công! Đã điền thẳng dữ liệu vào cấu trúc gốc của bạn.")
                st.download_button(
                    label="📥 Tải File Excel (Chuẩn Form)",
                    data=output.getvalue(),
                    file_name="Bao_Cao_Ton_Kho_Hoan_Thien.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Lỗi xử lý file: {e}")
else:
    if lottie_hoat_anh:
        st_lottie(lottie_hoat_anh, height=300, key="cho_file")
    st.info("👉 Hãy nạp File Báo cáo Tồn Kho (Mẫu 2356) và File Dữ liệu Đối Chiếu để bắt đầu.")

hien_thi_tro_ly_ai()
