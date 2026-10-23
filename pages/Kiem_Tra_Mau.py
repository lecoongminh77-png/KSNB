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
        if r.status_code != 200: return None
        return r.json()
    except: return None

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
    if kl == 0: return ""
    if kl < 2000: return "Dưới 2 tấn"
    elif kl <= 5000: return "Từ 2 - 5 tấn"
    elif kl <= 7000: return "Từ 5 - 7 tấn"
    elif kl <= 10000: return "Từ 7 - 10 tấn"
    elif kl <= 12000: return "Từ 10 - 12 tấn"
    elif kl <= 17000: return "Từ 12 - 17 tấn"
    elif kl <= 23000: return "Từ 17 - 23 tấn"
    else: return "Trên 23 tấn"

def xac_dinh_ky_chuan(hang, loai_loi):
    if not hang: return 0
    loai_str = str(loai_loi).lower().strip() if loai_loi else ""
    nhom = MA_TRAN_CHUAN.get(hang, {"khong_loi": 0, "loi_giay": 0, "loi_sat": 0})
    if "sắt" in loai_str or "sat" in loai_str: return nhom["loi_sat"]
    elif "giấy" in loai_str or "giay" in loai_str: return nhom["loi_giay"]
    else: return nhom["khong_loi"]

col_f1, col_f2 = st.columns(2)
with col_f1: f_tonkho = st.file_uploader("1. File Báo cáo Tồn Kho (Mẫu 2356)", type=["xlsx", "xls", "xlsb"])
with col_f2: f_doichieu = st.file_uploader("2. File Dữ liệu Đối Chiếu (OPM 105)", type=["xlsx", "xls", "xlsb"])

if f_tonkho is not None:
    if st.button("🚀 Bắt đầu tính toán tự động", use_container_width=True):
        with st.spinner("Đang điền trực tiếp dữ liệu vào Form gốc..."):
            try:
                # 1. Đọc File OPM 105 để lấy Mã đối chiếu
                set_ma_doi_chieu = set()
                if f_doichieu is not None:
                    try:
                        df_ref = pd.read_excel(f_doichieu)
                        # Lấy cột F (index 5) để đối chiếu
                        col_ref_ma = df_ref.columns[5] if len(df_ref.columns) > 5 else df_ref.columns[0]
                        set_ma_doi_chieu = set(df_ref[col_ref_ma].astype(str).str.replace(r"\.0$", "", regex=True).str.strip().str.upper())
                    except:
                        pass

                # 2. Xử lý File Tồn kho trực tiếp (Không dùng chèn cột)
                wb = openpyxl.load_workbook(f_tonkho)
                ws = wb.active
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

                # Dữ liệu bắt đầu chuẩn từ dòng 14
                for r in range(14, ws.max_row + 1):
                    # Lấy dữ liệu nguồn theo ĐÚNG tọa độ của file mẫu
                    ma_lo_cell = ws.cell(row=r, column=27).value   # Cột AA (Mã lô)
                    gross_cell = ws.cell(row=r, column=29).value   # Cột AC (Gross)
                    net_cell = ws.cell(row=r, column=30).value     # Cột AD (Net)
                    loi_cell = ws.cell(row=r, column=32).value     # Cột AF (Loại lõi)
                    loi_ngan_cell = ws.cell(row=r, column=9).value # Cột I (Lõi)

                    # Bỏ qua dòng trống không có dữ liệu
                    if not ma_lo_cell and not gross_cell:
                        continue 

                    # Làm sạch dữ liệu
                    ma_lo_val = str(ma_lo_cell).replace(".0", "").strip() if ma_lo_cell else ""
                    loi_val = str(loi_cell).strip() if loi_cell else ""
                    loi_ngan_val = str(loi_ngan_cell).strip() if loi_ngan_cell else ""

                    try: gross_num = float(gross_cell)
                    except: gross_num = 0.0
                    try: net_num = float(net_cell)
                    except: net_num = 0.0

                    # 3. Tính toán các giá trị
                    day_chuyen = ma_lo_val[3:6] if len(ma_lo_val) >= 6 else ""
                    hang = xac_dinh_hang(gross_num)
                    ky_chuan = xac_dinh_ky_chuan(hang, loi_val)
                    so_kg = gross_num - net_num
                    lech = so_kg - ky_chuan
                    loai_str = loi_ngan_val[0] if len(loi_ngan_val) > 0 else ""

                    is_trung = "Không"
                    if set_ma_doi_chieu and ma_lo_val.upper() in set_ma_doi_chieu:
                        is_trung = "Có"

                    # 4. GHI THẲNG VÀO CÁC CỘT TỪ A ĐẾN G CỦA DÒNG ĐÓ
                    ws.cell(row=r, column=1).value = is_trung      # Cột A: Trùng (Có/Không)
                    ws.cell(row=r, column=2).value = hang          # Cột B: Hạng
                    ws.cell(row=r, column=3).value = so_kg         # Cột C: Số kg
                    ws.cell(row=r, column=4).value = lech          # Cột D: Ký chênh lệch
                    ws.cell(row=r, column=5).value = ky_chuan      # Cột E: Số ký chuẩn
                    ws.cell(row=r, column=6).value = loai_str      # Cột F: Loại
                    ws.cell(row=r, column=7).value = day_chuyen    # Cột G: Dây chuyền

                    # 5. TÔ VÀNG DÒNG VƯỢT CHUẨN MÀ KHÔNG TRÙNG
                    if is_trung == "Không" and lech > 0:
                        for c_idx in range(1, ws.max_column + 1):
                            ws.cell(row=r, column=c_idx).fill = yellow_fill

                # Lưu vào bộ nhớ và tải xuống
                output = io.BytesIO()
                wb.save(output)

                st.success(f"✅ Đã xử lý xong! Lần này các cột được điền chính xác vào Form gốc, không chèn lệch.")
                st.download_button(
                    label="📥 Tải File Excel Đã Tính Xong",
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
