import io
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import streamlit as st

# Bảng ma trận 
MA_TRAN_CHUAN = {
    "Dưới 2 tấn":    {"giay": 40,  "sat": 70,  "khac": 20},
    "Từ 2 - 5 tấn":  {"giay": 60,  "sat": 90,  "khac": 40},
    "Từ 5 - 7 tấn":  {"giay": 70,  "sat": 100, "khac": 50},
    "Từ 7 - 10 tấn": {"giay": 90,  "sat": 120, "khac": 60},
    "Từ 10 - 12 tấn":{"giay": 100, "sat": 120, "khac": 70},
    "Từ 12 - 17 tấn":{"giay": 110, "sat": 130, "khac": 80},
    "Từ 17 - 23 tấn":{"giay": 120, "sat": 140, "khac": 90},
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
    else: return "Từ 17 - 23 tấn"

def xac_dinh_ky_chuan(hang, loai_loi):
    if not hang: return 0
    loai_str = str(loai_loi).lower().strip() if loai_loi else ""
    nhom = MA_TRAN_CHUAN.get(hang, {"giay": 0, "sat": 0, "khac": 0})
    if "sắt" in loai_str or "sat" in loai_str: return nhom["sat"]
    elif "giấy" in loai_str or "giay" in loai_str: return nhom["giay"]
    else: return nhom["khac"]

st.set_page_config(page_title="Kiểm Tra Tồn Kho Cuộn", page_icon="📦", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    /* Ẩn Sidebar */
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    
    /* Giao diện */
    .stApp { background-color: #f5f5f7; font-family: -apple-system, sans-serif; color: #1d1d1f; }
    .apple-header-box { display: flex; justify-content: space-between; align-items: flex-end; padding: 20px 0 10px 0; margin-bottom: 5px; }
    .apple-main-title { font-size: 38px; font-weight: 700; letter-spacing: -0.025em; margin: 0; }
    div.stButton > button { background: #0071e3 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; }
</style>
""", unsafe_allow_html=True)

# --- NÚT TRANG CHỦ ---
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Trang chủ", use_container_width=True):
        st.switch_page("app_ksnb.py")
st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)

st.markdown("""
<div class="apple-header-box">
    <div><h1 class="apple-main-title">📦 KSNB: Kiểm Tra Trọng Lượng Cuộn </h1></div>
</div>
""", unsafe_allow_html=True)

# --- CHIA 2 Ô NẠP FILE ---
col_f1, col_f2 = st.columns(2)
with col_f1: 
    f_tonkho = st.file_uploader("1. Nạp File Tồn Kho (110 / Mẫu 2356)", type=["xlsx", "xlsb", "xls"])
with col_f2: 
    f_doichieu = st.file_uploader("2. Nạp File Đối Chiếu (OPM 105)", type=["xlsx", "xlsb", "xls"])

if f_tonkho is not None:
    if st.button("🚀 Quét & Đối soát", use_container_width=True):
        with st.spinner("Đang xử lý logic 110 & 105..."):
            try:
                # 1. ĐỌC FILE OPM 105 ĐỂ LẤY MÃ ĐỐI CHIẾU
                set_ma_doi_chieu = set()
                if f_doichieu is not None:
                    try:
                        df_ref = pd.read_excel(f_doichieu)
                        # Dựa trên form chuẩn, cột F (index 5) thường là mã lô
                        col_ref_ma = df_ref.columns[5] if len(df_ref.columns) > 5 else df_ref.columns[0]
                        set_ma_doi_chieu = set(df_ref[col_ref_ma].astype(str).str.replace(r"\.0$", "", regex=True).str.strip().str.upper())
                    except:
                        pass

                # 2. XỬ LÝ FILE TỒN KHO
                wb = openpyxl.load_workbook(f_tonkho)
                ws = wb.active
                
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

                for r in range(14, ws.max_row + 1):
                    # --- TỌA ĐỘ CỘT ---
                    ma_hang_cell = ws.cell(row=r, column=3).value      # Cột C (Mã hàng)
                    locator_cell = ws.cell(row=r, column=18).value     # Cột R (Locator)
                    rma_cell = ws.cell(row=r, column=19).value         # Cột S (RMA) - Đổi số 19 nếu cột bị lệch
                    ma_lo_cell = ws.cell(row=r, column=27).value       # Cột AA (Mã lô cuộn) - Dùng đối chiếu 105
                    
                    kg_loi_cell = ws.cell(row=r, column=29).value      # Cột AC (Kg lõi ghi sẵn)
                    net_cell = ws.cell(row=r, column=30).value         # Cột AD (Net)
                    gross_cell = ws.cell(row=r, column=31).value       # Cột AE (Gross)
                    loai_loi_cell = ws.cell(row=r, column=32).value    # Cột AF (Loại lõi chữ)

                    if not ma_hang_cell and not gross_cell: continue 

                    ma_hang = str(ma_hang_cell).strip()
                    locator = str(locator_cell).strip()
                    rma = str(rma_cell).strip().upper()
                    ma_lo = str(ma_lo_cell).replace(".0", "").strip().upper() if ma_lo_cell else ""

                    # BỘ LỌC
                    if rma == "CÓ": continue
                    if not (ma_hang.startswith('3') or ma_hang.startswith('4')): continue
                    if len(locator) < 4 or locator[3].upper() != 'B': continue

                    try: gross_num = float(gross_cell)
                    except: gross_num = 0.0
                    try: net_num = float(net_cell)
                    except: net_num = 0.0
                    try: 
                        kg_loi_num = float(kg_loi_cell)
                        has_explicit_core = True if kg_loi_num > 0 else False
                    except: 
                        kg_loi_num = 0.0
                        has_explicit_core = False

                    # KIỂM TRA TRÙNG OPM 105
                    is_trung = "Có" if (set_ma_doi_chieu and ma_lo in set_ma_doi_chieu) else "Không"

                    # TÍNH CHÊNH LỆCH
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

                    # AI QUÉT BẤT THƯỜNG
                    is_kinh_ngac = False
                    if net_num > gross_num: is_kinh_ngac = True 
                    elif gross_num > 0 and lech_tuyet_doi > (gross_num * 0.05): is_kinh_ngac = True
                    elif lech_tuyet_doi > 150: is_kinh_ngac = True 

                    if is_kinh_ngac or lech_tuyet_doi > 100:
                        trang_thai = "NGHIÊM TRỌNG"
                        fill_color = red_fill
                    elif lech_tuyet_doi > 30:
                        trang_thai = "NGHI NGỜ"
                        fill_color = yellow_fill
                    else:
                        trang_thai = "ĐẠT"
                        fill_color = None

                    # GHI KẾT QUẢ VÀO 4 CỘT ĐẦU (A, B, C, D)
                    ws.cell(row=r, column=1).value = trang_thai
                    ws.cell(row=r, column=2).value = is_trung
                    ws.cell(row=r, column=3).value = round(lech_tuyet_doi, 2)
                    ws.cell(row=r, column=4).value = ghi_chu

                    # TÔ MÀU CẢNH BÁO
                    if fill_color:
                        for c_idx in range(1, ws.max_column + 1):
                            ws.cell(row=r, column=c_idx).fill = fill_color

                output = io.BytesIO()
                wb.save(output)

                st.success("✅ Đã hoàn tất quét dữ liệu. Các rủi ro cao đã được tô đỏ!")
                st.download_button(
                    label="📥 Tải File Đã Kiểm Tra",
                    data=output.getvalue(),
                    file_name="KSNB_Bao_Cao_Ton_Kho.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Lỗi: {e}")
