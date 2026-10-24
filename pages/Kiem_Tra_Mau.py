import sys
import os
import requests 
import io
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import streamlit as st
from streamlit_lottie import st_lottie
from ai_chat import hien_thi_tro_ly_ai

# Bảng ma trận theo đúng hình ảnh cung cấp
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
    else: return "Từ 17 - 23 tấn" # Cập nhật theo bảng: lớn nhất là mục này

def xac_dinh_ky_chuan(hang, loai_loi):
    if not hang: return 0
    loai_str = str(loai_loi).lower().strip() if loai_loi else ""
    nhom = MA_TRAN_CHUAN.get(hang, {"giay": 0, "sat": 0, "khac": 0})
    if "sắt" in loai_str or "sat" in loai_str: return nhom["sat"]
    elif "giấy" in loai_str or "giay" in loai_str: return nhom["giay"]
    else: return nhom["khac"]

st.set_page_config(page_title="Kiểm Tra Tồn Kho Cuộn", page_icon="📦", layout="wide")

st.markdown("""
<style>
    [data-testid="stSidebarNav"] { display: none !important; }
    .stApp { background-color: #f5f5f7; font-family: -apple-system, sans-serif; color: #1d1d1f; }
    .apple-header-box { display: flex; justify-content: space-between; align-items: flex-end; padding: 20px 0 10px 0; margin-bottom: 5px; }
    .apple-main-title { font-size: 38px; font-weight: 700; letter-spacing: -0.025em; margin: 0; }
    div.stButton > button { background: #0071e3 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="apple-header-box">
    <div><h1 class="apple-main-title">📦 KSNB: Kiểm Tra Trọng Lượng Cuộn </h1></div>
</div>
""", unsafe_allow_html=True)

f_tonkho = st.file_uploader("Nạp File Báo cáo Tồn Kho (Mẫu 2356)", type=["xlsx", "xlsb"])

if f_tonkho is not None:
    if st.button("🚀 Quét & Đối soát", use_container_width=True):
        with st.spinner("Đang áp dụng bộ lọc KSNB và quét bất thường..."):
            try:
                wb = openpyxl.load_workbook(f_tonkho)
                ws = wb.active
                
                # Khai báo màu sắc cảnh báo
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

                for r in range(14, ws.max_row + 1):
                    # --- CẤU HÌNH TỌA ĐỘ CỘT ---
                    ma_hang_cell = ws.cell(row=r, column=3).value      # Cột C (Mã hàng)
                    locator_cell = ws.cell(row=r, column=18).value     # Cột R (Locator)
                    rma_cell = ws.cell(row=r, column=19).value         # GIẢ ĐỊNH: Cột S là RMA (Anh đổi số 19 thành số cột thực tế)
                    
                    kg_loi_cell = ws.cell(row=r, column=29).value      # Cột AC (Kg lõi ghi sẵn)
                    net_cell = ws.cell(row=r, column=30).value         # Cột AD (Net)
                    gross_cell = ws.cell(row=r, column=31).value       # Cột AE (Gross)
                    loai_loi_cell = ws.cell(row=r, column=32).value    # Cột AF (Loại lõi chữ)

                    # Bỏ qua dòng trống
                    if not ma_hang_cell and not gross_cell: continue 

                    ma_hang = str(ma_hang_cell).strip()
                    locator = str(locator_cell).strip()
                    rma = str(rma_cell).strip().upper()

                    # 1. ĐIỀU KIỆN LỌC (FILTERING)
                    if rma == "CÓ": continue
                    if not (ma_hang.startswith('3') or ma_hang.startswith('4')): continue
                    if len(locator) < 4 or locator[3].upper() != 'B': continue

                    # Chuyển đổi số liệu
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

                    lech_tuyet_doi = 0
                    trang_thai = "OK"

                    # 2. KIỂM TRA THEO 2 KỊCH BẢN
                    if has_explicit_core:
                        # Kịch bản 1: Có ghi số kg lõi
                        calc_gross = net_num + kg_loi_num
                        lech_tuyet_doi = abs(calc_gross - gross_num)
                        ghi_chu = "KT công thức lõi"
                    else:
                        # Kịch bản 2: Không ghi số kg lõi, tính theo ma trận
                        hang = xac_dinh_hang(gross_num)
                        ky_chuan = xac_dinh_ky_chuan(hang, loai_loi_cell)
                        loi_thuc_te = gross_num - net_num
                        lech_tuyet_doi = abs(loi_thuc_te - ky_chuan)
                        ghi_chu = "KT theo ma trận"

                    # 3. QUÉT AI (ANOMALY DETECTION - Cảnh báo kinh ngạc)
                    is_kinh_ngac = False
                    if net_num > gross_num: 
                        is_kinh_ngac = True # Net lớn hơn Gross là vô lý
                    elif gross_num > 0 and lech_tuyet_doi > (gross_num * 0.05):
                        is_kinh_ngac = True # Lệch quá 5% tổng trọng lượng cuộn
                    elif lech_tuyet_doi > 150:
                        is_kinh_ngac = True # Lệch con số tuyệt đối quá lớn

                    # 4. PHÂN LOẠI & TÔ MÀU CẢNH BÁO
                    if is_kinh_ngac or lech_tuyet_doi > 100:
                        trang_thai = "NGHIÊM TRỌNG"
                        fill_color = red_fill
                    elif lech_tuyet_doi > 30:
                        trang_thai = "NGHI NGỜ"
                        fill_color = yellow_fill
                    else:
                        trang_thai = "ĐẠT"
                        fill_color = None

                    # 5. GHI KẾT QUẢ VÀO FILE
                    ws.cell(row=r, column=1).value = trang_thai
                    ws.cell(row=r, column=2).value = round(lech_tuyet_doi, 2)
                    ws.cell(row=r, column=3).value = ghi_chu

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

hien_thi_tro_ly_ai()
