import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Kiểm Tra Mẫu 2356 - Tồn Kho Cuộn",
    page_icon="📦",
    layout="wide"
)

# CSS giao diện tối giản chuẩn Apple
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

# Thanh điều hướng (Chỉ giữ nút Trang chủ)
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Trang chủ", use_container_width=True):
        st.switch_page("app_ksnb.py")

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)

# Tiêu đề trang
st.markdown("""
<div class="apple-header-box">
    <div>
        <h1 class="apple-main-title">📦 Kiểm Tra Mẫu 2356 (Tồn Kho Cuộn)</h1>
    </div>
    <div>
        <div class="apple-sub-text">
            Công cụ tự động hóa công thức<br>xử lý dữ liệu tồn kho siêu tốc.
            <span>Lê Công Minh Workspace ›</span>
        </div>
    </div>
</div>
<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>
""", unsafe_allow_html=True)

# Bảng tra chuẩn từ Sheet 2
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

# Khu vực Upload và xử lý
uploaded_file = st.file_uploader("📂 Chọn file Báo cáo tồn kho (Hỗ trợ .xlsx, .xls):", type=["xlsx", "xls"])

if uploaded_file is not None:
    if st.button("🚀 Bắt đầu tính toán tự động", use_container_width=True):
        with st.spinner("Đang tính toán các công thức A-G cho toàn bộ dữ liệu..."):
            try:
                # Đọc dữ liệu từ dòng tiêu đề (Dòng 12, index 11)
                try:
                    df = pd.read_excel(uploaded_file, skiprows=11)
                except Exception:
                    df = pd.read_excel(uploaded_file)

                # Tìm các cột dữ liệu chính dựa theo tên hoặc vị trí
                col_ma_lo = next((c for c in df.columns if "mã lô" in str(c).lower()), None)
                col_gross = next((c for c in df.columns if "gross" in str(c).lower()), None)
                col_net = next((c for c in df.columns if "net" in str(c).lower()), None)
                col_loai_loi = next((c for c in df.columns if "loại lõi" in str(c).lower()), None)

                # Làm sạch dữ liệu số
                df["_gross_num"] = pd.to_numeric(df[col_gross], errors="coerce").fillna(0) if col_gross else 0
                df["_net_num"] = pd.to_numeric(df[col_net], errors="coerce").fillna(0) if col_net else 0

                # 1. Cột G: Dây chuyền = MID(Mã lô, 4, 3)
                if col_ma_lo:
                    s_ma = df[col_ma_lo].astype(str).str.strip()
                    col_day_chuyen = s_ma.apply(lambda x: x[3:6] if len(x) >= 6 else "")
                else:
                    col_day_chuyen = ""

                # 2. Cột B: Hạng (tính theo Net hoặc Gross)
                col_hang = df["_net_num"].apply(xac_dinh_hang)

                # 3. Cột E: Số ký chuẩn
                if col_loai_loi:
                    col_ky_chuan = [xac_dinh_ky_chuan(h, l) for h, l in zip(col_hang, df[col_loai_loi])]
                else:
                    col_ky_chuan = [xac_dinh_ky_chuan(h, "") for h in col_hang]
                col_ky_chuan = pd.Series(col_ky_chuan, index=df.index)

                # 4. Cột C: Số kg (Gross - Net)
                col_so_kg = df["_gross_num"] - df["_net_num"]

                # 5. Cột D: Ký chênh lệch (Số kg - Số ký chuẩn)
                col_lech = col_so_kg - col_ky_chuan

                # 6. Cột A: Trùng mã lô
                if col_ma_lo:
                    counts = df[col_ma_lo].value_counts()
                    col_trung = df[col_ma_lo].map(counts).fillna(1).astype(int)
                else:
                    col_trung = 1

                # Loại bỏ cột tính toán phụ
                df.drop(columns=["_gross_num", "_net_num"], inplace=True, errors="ignore")

                # Chèn lại chính xác 7 cột mới vào đầu bảng (A -> G)
                df.insert(0, "Dây chuyền", col_day_chuyen)
                df.insert(0, "loại", "")
                df.insert(0, "số ký chuẩn", col_ky_chuan)
                df.insert(0, "ký chênh lệch", col_lech)
                df.insert(0, "số kg", col_so_kg)
                df.insert(0, "hạng", col_hang)
                df.insert(0, "Trùng ", col_trung)

                # Xuất ra file Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Data_Da_Tinh_Toan")

                st.success(f"✅ Đã xử lý xong toàn bộ {len(df):,} dòng dữ liệu!")
                st.download_button(
                    label="📥 Tải File Excel Đã Tính Toán Xong",
                    data=output.getvalue(),
                    file_name="Bao_Cao_Ton_Kho_Da_Tinh.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Lỗi xử lý file: {e}")
else:
    st.info("👉 Hãy nạp file Báo cáo tồn kho cuộn để tính toán.")
