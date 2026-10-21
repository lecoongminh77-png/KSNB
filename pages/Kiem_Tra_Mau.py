import sys
import os
import requests # Bắt buộc phải có để dùng r = requests.get(url)

# 1. Ép Python nhận diện thư mục gốc (nơi chứa ai_chat.py) ngay lập tức
root_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root_path not in sys.path:
    sys.path.insert(0, root_path)

# 2. Khai báo các thư viện khác
import io
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# 3. Import module nội bộ
from ai_chat import hien_thi_tro_ly_ai

# Hàm tải hình động
def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()


# Gắn link bạn vừa copy
lottie_hoat_anh = load_lottieurl("https://lottie.host/654b9803-9ee2-4c1b-91e3-3f074d4f79d3/b4IPc6jdqR.json")

st.set_page_config(
    page_title="Kiểm Tra Tồn Kho Cuộn",
    page_icon="📦",
    layout="wide"
)

# CSS giao diện chuẩn Apple
st.markdown("""
<style>
    [data-testid="stSidebarNav"] { display: none !important; }
    .stApp {
        background-color: #f5f5f7;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif;
        color: #1d1d1f;
    }
    .apple-header-box {
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
        padding: 20px 0 10px 0;
        margin-bottom: 5px;
    }
    .apple-main-title {
        font-size: 38px;
        font-weight: 700;
        letter-spacing: -0.025em;
        margin: 0;
    }
    .apple-sub-text {
        font-size: 16px;
        font-weight: 600;
        text-align: right;
        margin: 0;
        line-height: 1.35;
    }
    .apple-sub-text span {
        color: #0071e3;
        font-size: 14px;
        display: block;
        margin-top: 2px;
        font-weight: 500;
    }
    div.stButton > button {
        background: #0071e3 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding: 10px 24px !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2) !important;
    }
</style>
""", unsafe_allow_html=True)

# Thanh điều hướng (Chỉ giữ duy nhất nút Trang chủ)
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Trang chủ", use_container_width=True):
        st.switch_page("app_ksnb.py")

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)

# Tiêu đề trang
st.markdown("""
<div class="apple-header-box">
    <div>
        <h1 class="apple-main-title">📦 Kiểm Tra Mẫu Tồn Kho Cuộn </h1>
    </div>
    <div>
        <div class="apple-sub-text">
            Tự động điền công thức & đối soát trùng lặp 2 file siêu tốc.
            <span>Lê Công Minh Workspace ›</span>
        </div>
    </div>
</div>
<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>
""", unsafe_allow_html=True)

# Ma trận chuẩn từ Sheet 2
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
    try:
        kl = float(kl)
    except:
        return ""
    if kl < 2000:
        return "Dưới 2 tấn"
    elif kl <= 5000:
        return "Từ 2 - 5 tấn"
    elif kl <= 7000:
        return "Từ 5 - 7 tấn"
    elif kl <= 10000:
        return "Từ 7 - 10 tấn"
    elif kl <= 12000:
        return "Từ 10 - 12 tấn"
    elif kl <= 17000:
        return "Từ 12 - 17 tấn"
    elif kl <= 23000:
        return "Từ 17 - 23 tấn"
    else:
        return "Trên 23 tấn"

def xac_dinh_ky_chuan(hang, loai_loi):
    loai_str = str(loai_loi).lower().strip() if pd.notna(loai_loi) else ""
    nhom = MA_TRAN_CHUAN.get(hang, {"khong_loi": 0, "loi_giay": 0, "loi_sat": 0})
    if "sắt" in loai_str or "sat" in loai_str:
        return nhom["loi_sat"]
    elif "giấy" in loai_str or "giay" in loai_str:
        return nhom["loi_giay"]
    else:
        return nhom["khong_loi"]

# 2 Ô NẠP FILE
col_f1, col_f2 = st.columns(2)
with col_f1:
    f_tonkho = st.file_uploader("1. File Báo cáo Tồn Kho (Mẫu 2356)", type=["xlsx", "xls", "xlsb"])
with col_f2:
    f_doichieu = st.file_uploader("2. File Dữ liệu Đối Chiếu (OPM 105 / File nguồn)", type=["xlsx", "xls", "xlsb"])

if f_tonkho is not None:
    if st.button("🚀 Bắt đầu tính toán tự động", use_container_width=True):
        with st.spinner("Đang xử lý dữ liệu và đối soát trùng lặp..."):
            try:
                # 1. Đọc file tồn kho
                try:
                    df = pd.read_excel(f_tonkho, skiprows=11)
                except Exception:
                    df = pd.read_excel(f_tonkho)

                # Tìm các cột chính trong file tồn kho
                col_ma_lo = next((c for c in df.columns if "mã lô" in str(c).lower() or "cuộn" in str(c).lower()), None)
                col_gross = next((c for c in df.columns if "gross" in str(c).lower()), None)
                col_net = next((c for c in df.columns if "net" in str(c).lower()), None)
                col_loai_loi = next((c for c in df.columns if "loại lõi" in str(c).lower()), None)

                # Làm sạch dữ liệu số
                df["_gross_num"] = pd.to_numeric(df[col_gross], errors="coerce").fillna(0) if col_gross else 0
                df["_net_num"] = pd.to_numeric(df[col_net], errors="coerce").fillna(0) if col_net else 0

                # Dây chuyền = MID(Mã lô, 4, 3)
                if col_ma_lo:
                    s_ma = df[col_ma_lo].astype(str).str.strip()
                    col_day_chuyen = s_ma.apply(lambda x: x[3:6] if len(x) >= 6 else "")
                else:
                    col_day_chuyen = ""

                # Hạng & Số ký chuẩn
                col_hang = df["_net_num"].apply(xac_dinh_hang)
                if col_loai_loi:
                    col_ky_chuan = [xac_dinh_ky_chuan(h, l) for h, l in zip(col_hang, df[col_loai_loi])]
                else:
                    col_ky_chuan = [xac_dinh_ky_chuan(h, "") for h in col_hang]
                col_ky_chuan = pd.Series(col_ky_chuan, index=df.index)

                # Số kg thực tế & Ký chênh lệch
                col_so_kg = df["_gross_num"] - df["_net_num"]
                col_lech = col_so_kg - col_ky_chuan

                # Xử lý Cột A (Trùng): Đối chiếu sang File thứ 2
                set_ma_doi_chieu = set()
                if f_doichieu is not None:
                    try:
                        df_ref = pd.read_excel(f_doichieu)
                        # Tìm cột chứa mã lô trong file đối chiếu
                        col_ref_ma = next((c for c in df_ref.columns if any(k in str(c).lower() for k in ["mã lô", "mã hàng", "cuộn", "item", "lot"])), df_ref.columns[0])
                        set_ma_doi_chieu = set(
                            df_ref[col_ref_ma].astype(str).str.replace(r"\.0$", "", regex=True).str.strip().str.upper()
                        )
                    except Exception:
                        pass

                if set_ma_doi_chieu and col_ma_lo:
                    clean_ma_lo = df[col_ma_lo].astype(str).str.replace(r"\.0$", "", regex=True).str.strip().str.upper()
                    col_trung = clean_ma_lo.apply(lambda x: "Trùng" if x in set_ma_doi_chieu else "")
                else:
                    col_trung = ""

                # Bỏ cột phụ tạm thời
                df.drop(columns=["_gross_num", "_net_num"], inplace=True, errors="ignore")

                # Chèn 7 cột A-G vào đầu bảng
                df.insert(0, "Dây chuyền", col_day_chuyen)
                df.insert(0, "loại", "")
                df.insert(0, "số ký chuẩn", col_ky_chuan)
                df.insert(0, "ký chênh lệch", col_lech)
                df.insert(0, "số kg", col_so_kg)
                df.insert(0, "hạng", col_hang)
                df.insert(0, "Trùng ", col_trung)

                # Xuất ra file Excel kết quả
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Data_Da_Tinh_Toan")

                st.success(f"✅ Đã xử lý xong {len(df):,} dòng dữ liệu!")
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
    # Đoạn này sẽ hiện hình động nếu người dùng chưa nạp file
    st_lottie(lottie_hoat_anh, height=300, key="cho_file")
    st.info("👉 Hãy nạp File Báo cáo Tồn Kho (Mẫu 2356) và File Dữ liệu Đối Chiếu để bắt đầu.")
hien_thi_tro_ly_ai()
