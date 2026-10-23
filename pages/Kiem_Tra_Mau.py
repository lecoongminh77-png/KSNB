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
import streamlit as st
from streamlit_lottie import st_lottie
from openpyxl.styles import PatternFill

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
        with st.spinner("Đang phân tích cấu trúc ẩn của file..."):
            try:
                # Đọc GIỮ NGUYÊN cấu trúc (không lấy tên cột tự động để tránh lỗi gộp ô của ERP)
                df = pd.read_excel(f_tonkho, header=None)
                
                # Đảm bảo df có đủ số cột cần thiết (Chống lỗi nếu file hụt cột)
                while len(df.columns) < 26:
                    df[len(df.columns)] = ""

                # 1. Tạo 7 cột rỗng với độ dài bằng file gốc
                col_trung = [""] * len(df)
                col_hang = [""] * len(df)
                col_so_kg = [""] * len(df)
                col_lech = [""] * len(df)
                col_ky_chuan = [""] * len(df)
                col_loai = [""] * len(df)
                col_day_chuyen = [""] * len(df)

                # 2. Ghi Tiêu đề cho 7 cột chèn vào ngay Dòng 13 (Tức index 12 trong Python)
                header_idx = 12
                if len(df) > header_idx:
                    col_trung[header_idx] = "Trùng"
                    col_hang[header_idx] = "hạng"
                    col_so_kg[header_idx] = "số kg"
                    col_lech[header_idx] = "ký chênh lệch"
                    col_ky_chuan[header_idx] = "số ký chuẩn"
                    col_loai[header_idx] = "loại"
                    col_day_chuyen[header_idx] = "Dây chuyền"

                # 3. Quét File Đối chiếu OPM 105 (Quét chết cột F - Index 5 y như COUNTIF)
                set_ma_doi_chieu = set()
                if f_doichieu is not None:
                    try:
                        df_ref = pd.read_excel(f_doichieu)
                        col_ref_ma = df_ref.columns[5] if len(df_ref.columns) > 5 else df_ref.columns[0]
                        set_ma_doi_chieu = set(df_ref[col_ref_ma].astype(str).str.replace(r"\.0$", "", regex=True).str.strip().str.upper())
                    except Exception:
                        pass

                # 4. Bắt đầu tính số từ Dòng 14 (Index 13)
                for i in range(13, len(df)):
                    # Lấy tọa độ cột CHÍNH XÁC (Sau khi trừ đi độ chênh 7 cột bạn chèn)
                    ma_lo = str(df.iloc[i, 19]).replace(".0", "").strip()    # Gốc T (Vào file bạn là AA)
                    gross_val = df.iloc[i, 21]                               # Gốc V (Vào file bạn là AC)
                    net_val = df.iloc[i, 22]                                 # Gốc W (Vào file bạn là AD)
                    loi_val = str(df.iloc[i, 24]).strip()                    # Gốc Y (Vào file bạn là AF)
                    
                    try: gross_num = float(gross_val)
                    except: gross_num = 0
                    try: net_num = float(net_val)
                    except: net_num = 0
                    
                    # Bỏ qua dòng rỗng không có ID hoặc thông số trọng lượng
                    if not ma_lo and gross_num == 0 and net_num == 0:
                        continue

                    # Bắt đầu xử lý Công thức chuẩn
                    day_chuyen = ma_lo[3:6] if len(ma_lo) >= 6 else ""
                    
                    # Hạng (Dựa vào Gross như công thức B14 = IF(AC14<...) của bạn)
                    hang = xac_dinh_hang(gross_num) 
                    ky_chuan = xac_dinh_ky_chuan(hang, loi_val)
                    so_kg = gross_num - net_num
                    lech = so_kg - ky_chuan
                    
                    # Trùng lặp
                    if set_ma_doi_chieu:
                        is_trung = "Có" if ma_lo.upper() in set_ma_doi_chieu else "Không"
                    else:
                        is_trung = "Không"

                    # Lưu vào danh sách
                    col_trung[i] = is_trung
                    col_hang[i] = hang
                    col_so_kg[i] = so_kg
                    col_lech[i] = lech
                    col_ky_chuan[i] = ky_chuan
                    col_day_chuyen[i] = day_chuyen

                # 5. Chèn 7 cột vào bảng DataFrame
                df.insert(0, "Dây chuyền_col", col_day_chuyen)
                df.insert(0, "loại_col", col_loai)
                df.insert(0, "số ký chuẩn_col", col_ky_chuan)
                df.insert(0, "ký chênh lệch_col", col_lech)
                df.insert(0, "số kg_col", col_so_kg)
                df.insert(0, "hạng_col", col_hang)
                df.insert(0, "Trùng_col", col_trung)

                # ==========================================
                # XUẤT EXCEL & TÔ VÀNG VƯỢT CHUẨN (Chuẩn Layout)
                # ==========================================
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    # header=False giấu đi tên cột của Pandas, giữ lại bộ khung gốc của file!
                    df.to_excel(writer, index=False, header=False, sheet_name="Data_Da_Tinh_Toan")
                    
                    ws = writer.sheets["Data_Da_Tinh_Toan"]
                    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                    
                    # Duyệt tô màu (Bắt đầu quét từ dòng dữ liệu số 14)
                    for i in range(13, len(df)):
                        trung_val = str(df.iloc[i, 0]).strip().lower()
                        lech_val = df.iloc[i, 3] # Cột Ký chênh lệch ở vị trí D
                        
                        try:
                            if trung_val == "không" and float(lech_val) > 0:
                                excel_row = i + 1  # Tính khớp sang tọa độ file Excel
                                for cell in ws[excel_row]:
                                    cell.fill = yellow_fill
                        except:
                            pass

                st.success(f"✅ Đã xử lý {len(df)-13:,} dòng. Đã fix hoàn toàn tọa độ cột dữ liệu!")
                st.download_button(
                    label="📥 Tải File Excel Đã Tính Xong",
                    data=output.getvalue(),
                    file_name="Bao_Cao_Ton_Kho_Da_Kiem_Tra.xlsx",
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
