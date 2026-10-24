import io
import zipfile
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
import streamlit as st

# --- BẢNG MA TRẬN KSNB ---
MA_TRAN_CHUAN = {
    "Dưới 2 tấn":    {"giay": 40,  "sat": 70,  "khac": 20},
    "Từ 2 - 5 tấn":  {"giay": 60,  "sat": 90,  "khac": 40},
    "Từ 5 - 7 tấn":  {"giay": 70,  "sat": 100, "khac": 50},
    "Từ 7 - 10 tấn": {"giay": 90,  "sat": 120, "khac": 60},
    "Từ 10 - 12 tấn":{"giay": 100, "sat": 120, "khac": 70},
    "Từ 12 - 17 tấn":{"giay": 110, "sat": 130, "khac": 80},
    "Từ 17 - 23 tấn":{"giay": 120, "sat": 140, "khac": 90},
}

def safe_str(val):
    return str(val).strip() if pd.notna(val) and val is not None else ""

def safe_float(val):
    if pd.isna(val) or val is None or str(val).strip() == "": return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace(",", "").replace(" ", "")
    try: return float(s)
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
st.set_page_config(page_title="KSNB Tồn Kho", page_icon="📦", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    .stApp { background-color: #f5f5f7; font-family: -apple-system, sans-serif; color: #1d1d1f; }
    .apple-header-box { padding: 20px 0 10px 0; margin-bottom: 5px; border-bottom: 1px solid #e2e8f0; }
    .apple-main-title { font-size: 38px; font-weight: 700; margin: 0; }
    div.stButton > button { background: #0071e3 !important; color: white !important; border-radius: 980px !important; border: none !important; transition: 0.3s; }
    div.stButton > button:hover { background: #005bb5 !important; transform: scale(1.02); }
    div.stDownloadButton > button { background: #34c759 !important; }
</style>
""", unsafe_allow_html=True)

col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Về Trang Chủ", use_container_width=True):
        st.switch_page("app_ksnb.py")

st.markdown("<div class='apple-header-box'><h1 class='apple-main-title'>📦 KSNB: Phân Tích Trọng Lượng Cuộn</h1></div>", unsafe_allow_html=True)

# --- VÙNG NẠP FILE KÉO THẢ ĐA TẬP TIN ---
col_f1, col_f2 = st.columns(2)
with col_f1: 
    f_tonkhos = st.file_uploader("1. Nạp File Tồn Kho 110 (Có thể nạp NHIỀU file 1 lúc)", type=["xlsx", "xlsb", "xls"], accept_multiple_files=True)
with col_f2: 
    f_doichieu = st.file_uploader("2. Nạp File RMA 105 (Tuỳ chọn)", type=["xlsx", "xlsb", "xls"])

if not f_tonkhos:
    st.info("💡 Kéo thả một hoặc nhiều file Tồn Kho 110 vào ô số 1 để bắt đầu KSNB.")

if f_tonkhos:
    if st.button("🚀 Bắt đầu Quét & Dọn Dẹp File", use_container_width=True):
        
        progress_text = st.empty()
        
        try:
            # 1. ĐỌC FILE 105 ĐỂ LẤY MÃ LÔ HÀNG TRẢ RMA
            set_rma = set()
            if f_doichieu:
                progress_text.text("Đang đọc danh sách RMA (105)...")
                try:
                    df_rma = pd.read_excel(f_doichieu, header=None)
                    for idx, row in df_rma.iterrows():
                        val = str(row.iloc[5]).replace(".0", "").strip().upper()
                        if val and val != "NAN" and val != "MÃ LÔ":
                            set_rma.add(val)
                except Exception as e:
                    st.warning(f"⚠️ Không thể đọc file 105: {e}")

            # 2. XỬ LÝ DANH SÁCH FILE 110
            processed_files = {}
            
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            bold_font = Font(bold=True)

            tong_loi = 0
            tong_da_xoa = 0

            for f_tk in f_tonkhos:
                progress_text.text(f"Đang xử lý: {f_tk.name}...")
                wb = openpyxl.load_workbook(f_tk)
                ws = wb.active

                # BƯỚC 1: CHÈN THÊM 9 CỘT TRỐNG VÀO ĐẦU TIÊN (ĐẨY DỮ LIỆU CŨ LÙI SANG PHẢI 9 CỘT)
                ws.insert_cols(1, 9)

                # BƯỚC 2: TẠO DÒNG TIÊU ĐỀ CHO 9 CỘT MỚI Ở DÒNG SỐ 13
                tieu_de = [
                    "TRẠNG THÁI KSNB", "ĐỘ LỆCH (KG)", "TỶ LỆ LỆCH (%)", 
                    "KHOẢNG TRỌNG", "KG CHUẨN (Tham khảo)", "LÕI KHAI BÁO (Hệ thống)", 
                    "CĂN CỨ ĐỐI SOÁT", "CHI TIẾT LỖI", "BỘ LỌC KHO"
                ]
                for i, col_name in enumerate(tieu_de):
                    cell = ws.cell(row=13, column=i+1)
                    cell.value = col_name
                    cell.font = bold_font

                max_r = ws.max_row
                
                # BƯỚC 3: QUÉT TỪ DƯỚI LÊN ĐỂ XOÁ & TÍNH TOÁN
                for r in range(max_r, 13, -1):
                    # Tọa độ cột đã thay đổi (Cộng thêm 9 so với gốc):
                    # Cũ Cột B(2) -> Mới Cột K(11): Mã hàng
                    # Cũ Cột R(18) -> Mới Cột AA(27): Locator
                    # Cũ Cột T(20) -> Mới Cột AC(29): Mã lô
                    # Cũ Cột V(22) -> Mới Cột AE(31): Gross
                    # Cũ Cột W(23) -> Mới Cột AF(32): Net
                    # Cũ Cột X(24) -> Mới Cột AG(33): Kg Lõi khai báo
                    # Cũ Cột Y(25) -> Mới Cột AH(34): Loại lõi

                    ma_hang = safe_str(ws.cell(row=r, column=11).value)
                    ma_lo = safe_str(ws.cell(row=r, column=29).value).replace(".0", "").upper()
                    gross_num = safe_float(ws.cell(row=r, column=31).value)
                    
                    # - Lọc 1: Xoá dòng trống rác
                    if not ma_hang and gross_num == 0:
                        ws.delete_rows(r)
                        continue

                    # - Lọc 2: Xoá cuộn hàng RMA
                    if ma_lo in set_rma and ma_lo != "":
                        ws.delete_rows(r)
                        tong_da_xoa += 1
                        continue

                    # - Lọc 3: Xoá cuộn KHÔNG phải đầu 3 hoặc 4
                    if not (ma_hang.startswith('3') or ma_hang.startswith('4')):
                        ws.delete_rows(r)
                        tong_da_xoa += 1
                        continue

                    # === NẾU ĐẠT ĐIỀU KIỆN, TIẾN HÀNH KIỂM TOÁN TỒN KHO ===
                    locator = safe_str(ws.cell(row=r, column=27).value)
                    net_num = safe_float(ws.cell(row=r, column=32).value)
                    kg_loi_san = safe_float(ws.cell(row=r, column=33).value)
                    loai_loi = safe_str(ws.cell(row=r, column=34).value)

                    nhom_kho = "Kho B" if (len(locator) >= 4 and locator[3].upper() == 'B') else "Kho Khác"
                    
                    has_explicit_core = True if kg_loi_san > 0 else False
                    chi_tiet_loi = []
                    lech = 0
                    can_cu = ""
                    
                    # Phân tích theo Ma trận chuẩn
                    hang = xac_dinh_hang(gross_num)
                    ky_chuan = xac_dinh_ky_chuan(hang, loai_loi)
                    
                    if has_explicit_core:
                        can_cu = "Tính theo lõi khai báo (Hệ thống)"
                        lech = abs((net_num + kg_loi_san) - gross_num)
                        
                        # AI bẫy mâu thuẫn hệ thống (Gross - Net != Lõi khai báo) (Cho phép sai số 2kg)
                        if abs((gross_num - net_num) - kg_loi_san) > 2:
                            chi_tiet_loi.append("Mâu thuẫn: Gross/Net và số Lõi ghi nhận")
                    else:
                        can_cu = "Áp dụng Ma trận KSNB"
                        lech = abs((gross_num - net_num) - ky_chuan)

                    ty_le_lech = (lech / gross_num * 100) if gross_num > 0 else 0
                    is_kinh_ngac = False
                    
                    # Các bẫy KSNB (AI Cảnh báo)
                    if net_num > gross_num:
                        is_kinh_ngac = True
                        chi_tiet_loi.append("Lỗi: Net lớn hơn Gross")
                    elif gross_num > 0 and lech > (gross_num * 0.05):
                        is_kinh_ngac = True
                        chi_tiet_loi.append(f"Lỗi: Chênh > 5% Gross")
                    elif lech > 150:
                        is_kinh_ngac = True
                        chi_tiet_loi.append(f"Lỗi: Chênh lệch > 150kg")

                    # Quyết định Trạng thái và Tô màu
                    if is_kinh_ngac or lech > 100 or len(chi_tiet_loi) > 0:
                        trang_thai = "NGHIÊM TRỌNG"
                        fill_color = red_fill
                        tong_loi += 1
                    elif lech > 30:
                        trang_thai = "NGHI NGỜ"
                        fill_color = yellow_fill
                        tong_loi += 1
                    else:
                        trang_thai = "ĐẠT"
                        fill_color = None
                        
                    ghi_chu = " | ".join(chi_tiet_loi) if chi_tiet_loi else ("Vượt chuẩn nhẹ" if lech > 30 else "")

                    # BƯỚC 4: GHI 9 CỘT DỮ LIỆU KẾT QUẢ VÀO EXCEL
                    ws.cell(row=r, column=1).value = trang_thai
                    ws.cell(row=r, column=2).value = round(lech, 2)
                    ws.cell(row=r, column=3).value = f"{round(ty_le_lech, 2)}%"
                    ws.cell(row=r, column=4).value = hang
                    ws.cell(row=r, column=5).value = ky_chuan if ky_chuan > 0 else "-"
                    ws.cell(row=r, column=6).value = kg_loi_san if has_explicit_core else "-"
                    ws.cell(row=r, column=7).value = can_cu
                    ws.cell(row=r, column=8).value = ghi_chu
                    ws.cell(row=r, column=9).value = nhom_kho

                    # Tô màu cảnh báo
                    if fill_color:
                        for c_idx in range(1, ws.max_column + 1):
                            ws.cell(row=r, column=c_idx).fill = fill_color

                file_stream = io.BytesIO()
                wb.save(file_stream)
                wb.close()
                processed_files[f"KSNB_{f_tk.name}"] = file_stream.getvalue()

            progress_text.empty()
            
            # --- BÁO CÁO NHANH ---
            st.toast("Hoàn thành Dọn dẹp & Quét dữ liệu KSNB!", icon="🎉")
            col1, col2 = st.columns(2)
            col1.metric("🗑️ Cuộn rác & RMA đã dọn dẹp", tong_da_xoa)
            col2.metric("⚠️ Phát hiện cuộn bất thường", tong_loi, delta_color="inverse")
            st.info("💡 **Mẹo xem nhanh:** Tải file về ➔ Chọn Filter ở `Cột I` (Bộ Lọc Kho) ➔ Chọn 'Kho B' ➔ Sắp xếp (Sort) `Cột B` (Độ Lệch) từ cao xuống thấp để khoanh vùng thủ phạm!")

            # --- TRẢ FILE VỀ ---
            if len(processed_files) == 1:
                file_name, file_data = list(processed_files.items())[0]
                st.download_button(
                    label="📥 Tải Báo Cáo KSNB (Excel)",
                    data=file_data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for f_name, f_data in processed_files.items():
                        zip_file.writestr(f_name, f_data)
                
                st.download_button(
                    label=f"📥 Tải Về Zip Chứa {len(processed_files)} Báo Cáo KSNB",
                    data=zip_buffer.getvalue(),
                    file_name="Bao_Cao_KSNB_Ton_Kho.zip",
                    mime="application/zip",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Đã xảy ra lỗi hệ thống: {e}")
