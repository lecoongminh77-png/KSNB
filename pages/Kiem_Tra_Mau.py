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

# --- CẤU HÌNH GIAO DIỆN STREAMLIT ---
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

# --- VÙNG NẠP FILE ---
col_f1, col_f2 = st.columns(2)
with col_f1: 
    f_tonkhos = st.file_uploader("1. Nạp File Tồn Kho 110 (Nạp NHIỀU file cùng lúc)", type=["xlsx", "xlsb", "xls"], accept_multiple_files=True)
with col_f2: 
    f_doichieu = st.file_uploader("2. Nạp File RMA 105 (Tuỳ chọn)", type=["xlsx", "xlsb", "xls"])

if not f_tonkhos:
    st.info("💡 Kéo thả một hoặc nhiều file Tồn Kho 110 vào ô số 1 để bắt đầu KSNB.")

if f_tonkhos:
    if st.button("🚀 Bắt đầu Quét & Lọc Dữ Liệu (Bản Siêu Tốc)", use_container_width=True):
        
        progress_text = st.empty()
        
        try:
            # 1. ĐỌC FILE 105 LẤY MÃ LÔ RMA
            set_rma = set()
            if f_doichieu:
                progress_text.text("Đang trích xuất danh sách RMA từ file 105...")
                try:
                    df_rma = pd.read_excel(f_doichieu, header=None)
                    for idx, row in df_rma.iterrows():
                        val = str(row.iloc[5]).replace(".0", "").strip().upper()
                        if val and val != "NAN" and val != "MÃ LÔ":
                            set_rma.add(val)
                except Exception as e:
                    st.warning(f"⚠️ Không thể đọc file 105: {e}")

            # 2. XỬ LÝ DANH SÁCH FILE 110 (THUẬT TOÁN MỚI)
            processed_files = {}
            live_preview_data = [] 
            
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            no_fill = PatternFill(fill_type=None)
            bold_font = Font(bold=True)

            tong_loi = 0
            tong_da_xoa = 0

            for f_tk in f_tonkhos:
                progress_text.text(f"Đang xử lý tốc độ cao: {f_tk.name}...")
                wb = openpyxl.load_workbook(f_tk)
                ws = wb.active

                # BƯỚC 1: CHÈN THÊM 9 CỘT TRỐNG VÀO ĐẦU
                ws.insert_cols(1, 9)
                tieu_de = [
                    "TRẠNG THÁI KSNB", "ĐỘ LỆCH (KG)", "TỶ LỆ LỆCH (%)", 
                    "KHOẢNG TRỌNG", "KG CHUẨN (Tham khảo)", "LÕI KHAI BÁO (Hệ thống)", 
                    "CĂN CỨ ĐỐI SOÁT", "CHI TIẾT LỖI", "BỘ LỌC KHO"
                ]
                for i, col_name in enumerate(tieu_de):
                    cell = ws.cell(row=13, column=i+1)
                    try:
                        cell.value = col_name
                        cell.font = bold_font
                    except AttributeError:
                        pass # Bỏ qua lỗi MergedCell ở tiêu đề

                max_r = ws.max_row
                max_c = ws.max_column
                
                # BƯỚC 2: QUÉT & TÍNH TOÁN TRÊN RAM
                danh_sach_hop_le = [] # Chứa các dòng đã sạch rác
                
                for r in range(14, max_r + 1):
                    ma_hang = safe_str(ws.cell(row=r, column=11).value)
                    ma_lo = safe_str(ws.cell(row=r, column=29).value).replace(".0", "").upper()
                    gross_num = safe_float(ws.cell(row=r, column=31).value)
                    
                    # --- LỌC RÁC TỐC ĐỘ CAO ---
                    if not ma_hang and gross_num == 0:
                        continue
                    if ma_lo in set_rma and ma_lo != "":
                        tong_da_xoa += 1
                        continue
                    if not (ma_hang.startswith('3') or ma_hang.startswith('4')):
                        tong_da_xoa += 1
                        continue

                    # --- TÍNH TOÁN KSNB CHO DÒNG HỢP LỆ ---
                    locator = safe_str(ws.cell(row=r, column=27).value)
                    net_num = safe_float(ws.cell(row=r, column=32).value)
                    kg_loi_san = safe_float(ws.cell(row=r, column=33).value)
                    loai_loi = safe_str(ws.cell(row=r, column=34).value)

                    nhom_kho = "Kho B" if (len(locator) >= 4 and locator[3].upper() == 'B') else "Kho Khác"
                    has_explicit_core = True if kg_loi_san > 0 else False
                    chi_tiet_loi = []
                    lech = 0
                    can_cu = ""
                    
                    hang = xac_dinh_hang(gross_num)
                    ky_chuan = xac_dinh_ky_chuan(hang, loai_loi)
                    
                    if has_explicit_core:
                        can_cu = "Lõi khai báo (Hệ thống)"
                        lech = abs((net_num + kg_loi_san) - gross_num)
                        if abs((gross_num - net_num) - kg_loi_san) > 2:
                            chi_tiet_loi.append("Mâu thuẫn Gross/Net/Lõi")
                    else:
                        can_cu = "Ma trận KSNB"
                        lech = abs((gross_num - net_num) - ky_chuan)

                    ty_le_lech = (lech / gross_num * 100) if gross_num > 0 else 0
                    is_kinh_ngac = False
                    
                    if net_num > gross_num:
                        is_kinh_ngac = True
                        chi_tiet_loi.append("Lỗi: Net > Gross")
                    elif gross_num > 0 and lech > (gross_num * 0.05):
                        is_kinh_ngac = True
                        chi_tiet_loi.append(f"Lỗi: Chênh > 5%")
                    elif lech > 150:
                        is_kinh_ngac = True
                        chi_tiet_loi.append(f"Lỗi: Lệch > 150kg")

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
                        
                    ghi_chu = " | ".join(chi_tiet_loi) if chi_tiet_loi else ("Vượt chuẩn" if lech > 30 else "")

                    # Bưng toàn bộ dữ liệu của dòng này đưa vào RAM
                    row_data = [
                        trang_thai, round(lech, 2), f"{round(ty_le_lech, 2)}%", hang,
                        ky_chuan if ky_chuan > 0 else "-", kg_loi_san if has_explicit_core else "-",
                        can_cu, ghi_chu, nhom_kho
                    ]
                    # Chép nốt các cột gốc còn lại
                    for c in range(10, max_c + 1):
                        row_data.append(ws.cell(row=r, column=c).value)
                        
                    danh_sach_hop_le.append((row_data, fill_color))

                    if fill_color:
                        live_preview_data.append({
                            "File": f_tk.name, "Mã Lô": ma_lo, "Kho": nhom_kho,
                            "Trạng Thái": trang_thai, "Lệch (Kg)": round(lech, 2), "Lý do": ghi_chu
                        })

                # BƯỚC 3: GHI ĐÈ DỮ LIỆU SẠCH XUỐNG EXCEL KÈM BẪY LỖI GỘP Ô
                current_row = 14
                for row_data, fill_color in danh_sach_hop_le:
                    for c_idx, val in enumerate(row_data, start=1):
                        cell = ws.cell(row=current_row, column=c_idx)
                        try:
                            cell.value = val
                            if fill_color:
                                cell.fill = fill_color
                            else:
                                cell.fill = no_fill # Xóa màu cũ nếu có
                        except AttributeError:
                            # KHẮC PHỤC LỖI: Bỏ qua nhẹ nhàng nếu ghi đè trúng phần bị ẩn của ô đã gộp
                            pass
                    current_row += 1
                    
                # BƯỚC 4: XOÁ GỌN ĐUÔI THỪA 1 LẦN DUY NHẤT
                so_dong_thua = max_r - current_row + 1
                if so_dong_thua > 0:
                    ws.delete_rows(current_row, so_dong_thua)

                # Lưu File
                file_stream = io.BytesIO()
                wb.save(file_stream)
                wb.close()
                processed_files[f"KSNB_{f_tk.name}"] = file_stream.getvalue()

            progress_text.empty()
            st.toast("Hoàn thành Quét dữ liệu Siêu Tốc!", icon="🎉")
            
            # --- BÁO CÁO NHANH TRÊN WEB ---
            col1, col2 = st.columns(2)
            col1.metric("🗑️ Cuộn rác & RMA đã dọn dẹp", tong_da_xoa)
            col2.metric("⚠️ Phát hiện cuộn bất thường", tong_loi, delta_color="inverse")
            
            if live_preview_data:
                st.error("🚨 Danh sách các cuộn bị cảnh báo (Xem trước):")
                df_preview = pd.DataFrame(live_preview_data)
                df_preview = df_preview.sort_values(by="Lệch (Kg)", ascending=False).reset_index(drop=True)
                st.dataframe(df_preview, use_container_width=True)
            else:
                st.success("✅ Tuyệt vời! Tất cả các cuộn đều Đạt tiêu chuẩn.")

            # --- TRẢ FILE VỀ MÁY TÍNH ---
            st.markdown("<hr>", unsafe_allow_html=True)
            if len(processed_files) == 1:
                file_name, file_data = list(processed_files.items())[0]
                st.download_button(
                    label="📥 Tải File Báo Cáo KSNB",
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
                    label=f"📥 Tải Về Zip Chứa {len(processed_files)} Báo Cáo",
                    data=zip_buffer.getvalue(),
                    file_name="Bao_Cao_KSNB_Ton_Kho.zip",
                    mime="application/zip",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Đã xảy ra lỗi hệ thống: {e}")
