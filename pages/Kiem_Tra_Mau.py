import io
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import streamlit as st

# --- BẢNG MA TRẬN ---
MA_TRAN_CHUAN = {
    "Dưới 2 tấn":    {"giay": 40,  "sat": 70,  "khac": 20},
    "Từ 2 - 5 tấn":  {"giay": 60,  "sat": 90,  "khac": 40},
    "Từ 5 - 7 tấn":  {"giay": 70,  "sat": 100, "khac": 50},
    "Từ 7 - 10 tấn": {"giay": 90,  "sat": 120, "khac": 60},
    "Từ 10 - 12 tấn":{"giay": 100, "sat": 120, "khac": 70},
    "Từ 12 - 17 tấn":{"giay": 110, "sat": 130, "khac": 80},
    "Từ 17 - 23 tấn":{"giay": 120, "sat": 140, "khac": 90},
}

# --- HÀM TIỆN ÍCH AN TOÀN ---
def safe_str(val):
    """Hàm chống lỗi khi gặp ô trống (None)"""
    return str(val).strip() if val is not None else ""

def safe_float(val):
    """Hàm chống lỗi khi ô số bị gõ nhầm chữ"""
    try: return float(val)
    except: return 0.0

def xac_dinh_hang(kl):
    kl = safe_float(kl)
    if kl == 0: return ""
    if kl < 2000: return "Dưới 2 tấn"
    elif kl <= 5000: return "Từ 2 - 5 tấn"
    elif kl <= 7000: return "Từ 5 - 7 tấn"
    elif kl <= 10000: return "Từ 7 - 10 tấn"
    elif kl <= 12000: return "Từ 10 - 12 tấn"
    elif kl <= 17000: return "Từ 12 - 17 tấn"
    else: return "Từ 17 - 23 tấn"

def xac_dinh_ky_chuan(hang, loai_loi):
    if not hang: return 0
    loai_str = safe_str(loai_loi).lower()
    nhom = MA_TRAN_CHUAN.get(hang, {"giay": 0, "sat": 0, "khac": 0})
    if "sắt" in loai_str or "sat" in loai_str: return nhom["sat"]
    elif "giấy" in loai_str or "giay" in loai_str: return nhom["giay"]
    else: return nhom["khac"]

# --- CẤU HÌNH GIAO DIỆN ---
st.set_page_config(page_title="Kiểm Tra Tồn Kho Cuộn", page_icon="📦", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    /* Ẩn hoàn toàn sidebar */
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    
    .stApp { background-color: #f5f5f7; font-family: -apple-system, sans-serif; color: #1d1d1f; }
    .apple-header-box { display: flex; justify-content: space-between; align-items: flex-end; padding: 20px 0 10px 0; margin-bottom: 5px; }
    .apple-main-title { font-size: 38px; font-weight: 700; letter-spacing: -0.025em; margin: 0; }
    div.stButton > button { background: #0071e3 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; transition: all 0.3s ease; }
    div.stButton > button:hover { background: #005bb5 !important; transform: scale(1.02); }
    div.stDownloadButton > button { background: #34c759 !important; }
    div.stDownloadButton > button:hover { background: #2eb050 !important; }
</style>
""", unsafe_allow_html=True)

# --- HEADER & NÚT TRANG CHỦ ---
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Về Trang Chủ", use_container_width=True):
        st.switch_page("app_ksnb.py")
st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)

st.markdown("""
<div class="apple-header-box">
    <div><h1 class="apple-main-title">📦 KSNB: Kiểm Tra Trọng Lượng Cuộn </h1></div>
</div>
""", unsafe_allow_html=True)

# --- VÙNG NẠP FILE ---
col_f1, col_f2 = st.columns(2)
with col_f1: 
    f_tonkho = st.file_uploader("1. Nạp File Tồn Kho (110 / Mẫu 2356)", type=["xlsx", "xlsb", "xls"])
with col_f2: 
    f_doichieu = st.file_uploader("2. Nạp File Đối Chiếu (OPM 105)", type=["xlsx", "xlsb", "xls"])

# Cảnh báo nếu chưa nạp đủ
if f_tonkho is None:
    st.info("💡 Vui lòng nạp File Tồn Kho (110) để bắt đầu. Có thể nạp thêm File OPM 105 để đối soát mã lô.")

if f_tonkho is not None:
    if st.button("🚀 Bắt đầu Quét & Đối soát", use_container_width=True):
        
        # --- UI THANH TIẾN TRÌNH ---
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # 1. ĐỌC FILE 105 
            set_ma_doi_chieu = set()
            if f_doichieu is not None:
                progress_text.text("Đang đọc dữ liệu đối chiếu (OPM 105)...")
                try:
                    df_ref = pd.read_excel(f_doichieu)
                    col_ref_ma = df_ref.columns[5] if len(df_ref.columns) > 5 else df_ref.columns[0]
                    set_ma_doi_chieu = set(df_ref[col_ref_ma].astype(str).str.replace(r"\.0$", "", regex=True).str.strip().str.upper())
                except Exception as e:
                    st.warning(f"⚠️ Không thể đọc file 105: {e}")

            # 2. XỬ LÝ FILE 110
            progress_text.text("Đang tải biểu mẫu Tồn kho (110)...")
            wb = openpyxl.load_workbook(f_tonkho)
            ws = wb.active
            
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

            max_row_thuc_te = ws.max_row
            tong_so_dong = max_row_thuc_te - 13
            
            if tong_so_dong <= 0:
                st.error("❌ File Tồn kho không đúng chuẩn (không có dữ liệu từ dòng 14 trở đi).")
                st.stop()

            so_dong_loi = 0

            # QUÉT CẠN KIỆT TOÀN BỘ FILE TỪ DÒNG 14 ĐẾN CUỐI CÙNG
            for i, r in enumerate(range(14, max_row_thuc_te + 1)):
                try:
                    # Update thanh tiến trình
                    if i % 10 == 0 or i == tong_so_dong - 1:
                        phan_tram = min((i + 1) / tong_so_dong, 1.0)
                        progress_bar.progress(phan_tram)
                        progress_text.text(f"Đang quét dòng {r} / {max_row_thuc_te}...")

                    # --- LẤY TỌA ĐỘ BẰNG HÀM AN TOÀN ---
                    ma_hang = safe_str(ws.cell(row=r, column=3).value)      # Cột C (Mã hàng)
                    locator = safe_str(ws.cell(row=r, column=18).value)     # Cột R (Locator)
                    rma = safe_str(ws.cell(row=r, column=19).value).upper() # Cột S (RMA)
                    ma_lo = safe_str(ws.cell(row=r, column=27).value).replace(".0", "").upper() # Cột AA
                    
                    gross_num = safe_float(ws.cell(row=r, column=31).value) # Cột AE (Gross)
                    net_num = safe_float(ws.cell(row=r, column=30).value)   # Cột AD (Net)
                    kg_loi_num = safe_float(ws.cell(row=r, column=29).value) # Cột AC (Kg lõi sẵn)
                    loai_loi_cell = safe_str(ws.cell(row=r, column=32).value) # Cột AF (Loại lõi)

                    # Bỏ qua dòng trống (nhưng vẫn tiếp tục quét các dòng bên dưới)
                    if not ma_hang and gross_num == 0:
                        continue

                    # --- ĐIỀU KIỆN LỌC KSNB ---
                    if rma == "CÓ": continue
                    if not (ma_hang.startswith('3') or ma_hang.startswith('4')): continue
                    if len(locator) < 4 or locator[3].upper() != 'B': continue

                    has_explicit_core = True if kg_loi_num > 0 else False
                    is_trung = "Có" if (set_ma_doi_chieu and ma_lo in set_ma_doi_chieu) else "Không"
                    lech_tuyet_doi = 0

                    # --- TÍNH TOÁN ---
                    if has_explicit_core:
                        calc_gross = net_num + kg_loi_num
                        lech_tuyet_doi = abs(calc_gross - gross_num)
                        ghi_chu = "KT công thức lõi"
                    else:
                        hang = xac_dinh_hang(gross_num)
                        ky_chuan = xac_dinh_ky_chuan(hang, loai_loi_cell)
                        loi_thuc_te = gross_num - net_num
                        lech_tuyet_doi = abs(loi_thuc_te - ky_chuan)
                        ghi_chu = "KT theo ma trận"

                    # --- AI DETECT ---
                    is_kinh_ngac = False
                    if net_num > gross_num: is_kinh_ngac = True 
                    elif gross_num > 0 and lech_tuyet_doi > (gross_num * 0.05): is_kinh_ngac = True
                    elif lech_tuyet_doi > 150: is_kinh_ngac = True 

                    if is_kinh_ngac or lech_tuyet_doi > 100:
                        trang_thai = "NGHIÊM TRỌNG"
                        fill_color = red_fill
                        so_dong_loi += 1
                    elif lech_tuyet_doi > 30:
                        trang_thai = "NGHI NGỜ"
                        fill_color = yellow_fill
                        so_dong_loi += 1
                    else:
                        trang_thai = "ĐẠT"
                        fill_color = None

                    # --- GHI KẾT QUẢ VÀO 4 CỘT A, B, C, D ---
                    ws.cell(row=r, column=1).value = trang_thai
                    ws.cell(row=r, column=2).value = is_trung
                    ws.cell(row=r, column=3).value = round(lech_tuyet_doi, 2)
                    ws.cell(row=r, column=4).value = ghi_chu

                    if fill_color:
                        for c_idx in range(1, ws.max_column + 1):
                            ws.cell(row=r, column=c_idx).fill = fill_color

                except Exception as row_error:
                    # Bỏ qua dòng bị lỗi để không sập cả file
                    continue 

            # Dọn dẹp UI khi hoàn thành
            progress_bar.empty()
            progress_text.empty()
            
            output = io.BytesIO()
            wb.save(output)
            wb.close() # Giải phóng RAM

            st.toast("Hoàn thành quét dữ liệu!", icon="🎉")
            
            if so_dong_loi > 0:
                st.warning(f"⚠️ Phát hiện {so_dong_loi} dòng có chênh lệch đáng ngờ/nghiêm trọng. Hãy tải file để kiểm tra (các dòng đã được tô màu).")
            else:
                st.success("✅ Toàn bộ danh sách đạt chuẩn, không phát hiện chênh lệch bất thường!")

            # Nút tải file
            st.download_button(
                label="📥 Tải File Đã Kiểm Tra & Đánh Dấu",
                data=output.getvalue(),
                file_name="KSNB_Bao_Cao_Ton_Kho_Update.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Đã xảy ra lỗi hệ thống: {e}")
