import io
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import streamlit as st
import time

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
    if kl < 2000: return "Dưới 2 tấn"
    elif kl <= 5000: return "Từ 2 - 5 tấn"
    elif kl <= 7000: return "Từ 5 - 7 tấn"
    elif kl <= 10000: return "Từ 7 - 10 tấn"
    elif kl <= 12000: return "Từ 10 - 12 tấn"
    elif kl <= 17000: return "Từ 12 - 17 tấn"
    else: return "Từ 17 - 23 tấn"

def xac_dinh_ky_chuan(hang, loai_loi):
    loai_str = safe_str(loai_loi).lower()
    nhom = MA_TRAN_CHUAN.get(hang, {"giay": 0, "sat": 0, "khac": 0})
    if "sắt" in loai_str or "sat" in loai_str: return nhom["sat"]
    elif "giấy" in loai_str or "giay" in loai_str: return nhom["giay"]
    else: return nhom["khac"]

# --- GIAO DIỆN STREAMLIT ---
st.set_page_config(page_title="KSNB Tồn Kho", page_icon="📦", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    /* ẨN HOÀN TOÀN THANH SIDEBAR */
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    
    .stApp { background-color: #f5f5f7; }
    div.stButton > button { background: #0071e3 !important; color: white !important; border-radius: 8px !important; border: none !important; font-weight: bold; }
    div.stButton > button:hover { background: #005bb5 !important; transform: scale(1.02); }
    div.stDownloadButton > button { background: #34c759 !important; border-radius: 8px !important; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- NÚT VỀ TRANG CHỦ ---
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Về Trang Chủ", use_container_width=True):
        st.switch_page("app_ksnb.py") # Nhớ sửa thành file trang chủ của anh nếu cần

st.title("📦 HỆ THỐNG KSNB TỒN KHO (BẢN GỘP FILE)")

col_f1, col_f2 = st.columns(2)
with col_f1: 
    f_tonkhos = st.file_uploader("1. Nạp các File 110 (Kéo thả nhiều file)", type=["xlsx", "xlsb", "xls"], accept_multiple_files=True)
with col_f2: 
    f_doichieu = st.file_uploader("2. Nạp File RMA 105 (Tuỳ chọn)", type=["xlsx", "xlsb", "xls"])

if st.button("🚀 Bắt đầu Tổng hợp & Kiểm toán", use_container_width=True) and f_tonkhos:
    
    # 🌟 KHỞI TẠO THANH TIẾN ĐỘ 🌟
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # 1. ĐỌC FILE 105 RMA
        set_rma = set()
        if f_doichieu:
            status_text.text("Đang đọc dữ liệu đối chiếu RMA...")
            df_rma = pd.read_excel(f_doichieu, header=None)
            for idx, row in df_rma.iterrows():
                val = str(row.iloc[5]).replace(".0", "").strip().upper()
                if val and val != "NAN" and val != "MÃ LÔ": set_rma.add(val)

        # 2. CHUẨN BỊ TẠO 1 FILE BÁO CÁO DUY NHẤT
        wb_tong = openpyxl.Workbook()
        ws_tong = wb_tong.active
        ws_tong.title = "Bao_Cao_KSNB"
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="005bb5", end_color="005bb5", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        # Rút gọn chỉ còn 6 cột quan trọng
        tieu_de_ksnb = [
            "TRẠNG THÁI KSNB", "ĐỘ LỆCH KSNB (KG)", "GROSS - NET (KG)", 
            "KHOẢNG TRỌNG", "KG CHUẨN", "BỘ LỌC KHO"
        ]
        
        headers_goc = []
        tong_da_xoa = 0
        tong_loi = 0
        live_preview_data = []
        total_files = len(f_tonkhos)

        # 3. QUÉT DỮ LIỆU TỪ CÁC FILE NẠP VÀO
        for file_idx, f_tk in enumerate(f_tonkhos):
            status_text.text(f"Đang tải file vào bộ nhớ: {f_tk.name}... (Vui lòng đợi)")
            wb_doc = openpyxl.load_workbook(f_tk, data_only=True)
            ws_doc = wb_doc.active
            max_c = ws_doc.max_column
            max_r = ws_doc.max_row
            total_rows = max_r - 13 if max_r > 13 else 1
            
            # Lấy tiêu đề gốc từ file đầu tiên
            if file_idx == 0:
                for c in range(1, max_c + 1):
                    h12 = safe_str(ws_doc.cell(row=12, column=c).value)
                    h13 = safe_str(ws_doc.cell(row=13, column=c).value)
                    if h12 and h13 and h12 != h13: h_final = f"{h12} {h13}"
                    elif h12: h_final = h12
                    elif h13: h_final = h13
                    else: h_final = f"Cột {c}"
                    headers_goc.append(h_final)
                    
                full_headers = tieu_de_ksnb + headers_goc
                ws_tong.append(full_headers)
                for cell in ws_tong[1]: cell.font, cell.fill = header_font, header_fill

            # Đọc dữ liệu từ dòng 14 xuống
            for r in range(14, max_r + 1):
                
                # CẬP NHẬT THANH TIẾN ĐỘ
                if r % 500 == 0 or r == max_r:
                    row_prog = (r - 13) / total_rows
                    overall_prog = (file_idx + row_prog) / total_files
                    progress_bar.progress(min(overall_prog, 1.0))
                    status_text.text(f"🔄 Đang quét: {f_tk.name} (Dòng {r}/{max_r}) - {round(overall_prog * 100)}%")

                ma_hang = safe_str(ws_doc.cell(row=r, column=2).value)
                ten_hang = safe_str(ws_doc.cell(row=r, column=3).value)
                ma_lo = safe_str(ws_doc.cell(row=r, column=20).value).replace(".0", "").upper()
                gross_num = safe_float(ws_doc.cell(row=r, column=22).value)
                
                # LUẬT DỌN RÁC: Trống Gross hoặc Gross = 0 thì XÓA VĨNH VIỄN
                if gross_num <= 0: continue
                if not ma_hang: continue
                if ma_lo in set_rma and ma_lo != "":
                    tong_da_xoa += 1
                    continue
                if not (ma_hang.startswith('3') or ma_hang.startswith('4')):
                    tong_da_xoa += 1
                    continue

                row_goc = [ws_doc.cell(row=r, column=c).value for c in range(1, max_c + 1)]
                locator = safe_str(ws_doc.cell(row=r, column=18).value)
                net_num = safe_float(ws_doc.cell(row=r, column=23).value)
                kg_loi_san = safe_float(ws_doc.cell(row=r, column=24).value)
                loai_loi = safe_str(ws_doc.cell(row=r, column=25).value)

                # Lọc lấy Ký tự thứ 4 trong Locator làm Bộ lọc Kho
                bo_loc_kho = locator[3].upper() if len(locator) >= 4 else "-"
                
                trang_thai, lech, hang, kg_chuan = "", 0.0, "-", 0
                gross_net = gross_num - net_num
                fill_color = None
                is_error = False

                # BẪY LỖI NGHIỆP VỤ CƠ BẢN
                if gross_num < net_num:
                    trang_thai = "NGHIÊM TRỌNG (Gross < Net)"
                    fill_color = red_fill
                    is_error = True
                    tong_loi += 1
                elif gross_num == net_num:
                    if "thép dày" in ten_hang.lower():
                        trang_thai = "ĐẠT"
                    else:
                        trang_thai = "NGHIÊM TRỌNG (Lỗi Gross = Net)"
                        fill_color = red_fill
                        is_error = True
                        tong_loi += 1
                else:
                    # TÍNH TOÁN KSNB THEO "LUẬT CHỌN SỐ LỚN NHẤT"
                    hang = xac_dinh_hang(gross_num)
                    kg_chuan_ma_tran = xac_dinh_ky_chuan(hang, loai_loi)
                    kg_chuan = max(kg_chuan_ma_tran, kg_loi_san)
                    
                    lech = abs(gross_net - kg_chuan)

                    if lech > 100:
                        trang_thai = "NGHIÊM TRỌNG (Lệch > 100kg)"
                        fill_color = red_fill
                        is_error = True
                        tong_loi += 1
                    elif lech > 30:
                        trang_thai = "NGHI NGỜ (Lệch > 30kg)"
                        fill_color = yellow_fill
                        is_error = True
                        tong_loi += 1
                    else:
                        trang_thai = "ĐẠT"

                # Đóng gói dòng báo cáo 6 cột + Cột Gốc
                row_ksnb = [
                    trang_thai, round(lech, 2), round(gross_net, 2), 
                    hang, kg_chuan if kg_chuan > 0 else "-", bo_loc_kho
                ]
                full_row_data = row_ksnb + row_goc
                
                ws_tong.append(full_row_data)
                if fill_color:
                    for cell in ws_tong[ws_tong.max_row]: cell.fill = fill_color
                    
                if is_error:
                    live_preview_data.append({
                        "Tên Kho": bo_loc_kho, "Mã Lô": ma_lo, 
                        "Trạng Thái": trang_thai, "Lệch (Kg)": round(lech, 2), "Tên Hàng": ten_hang
                    })

        # 4. LƯU FILE EXCEL RA BỘ NHỚ
        status_text.text("Đang đóng gói file Excel Báo Cáo...")
        file_tong = io.BytesIO()
        wb_tong.save(file_tong)
        wb_tong.close()

        # 5. HIỂN THỊ KẾT QUẢ LÊN WEB
        progress_bar.empty()
        status_text.empty()
        st.toast("🎉 Phân tích Gộp File thành công!")
        
        c1, c2 = st.columns(2)
        c1.metric("🗑️ Cuộn rác & Khác mã đã dọn", tong_da_xoa)
        c2.metric("⚠️ Số cuộn phát hiện Lỗi/Hao hụt", tong_loi, delta_color="inverse")

        st.markdown("### 📥 XUẤT BÁO CÁO")
        st.download_button(
            label="📊 TẢI BÁO CÁO KSNB (Gộp tất cả các kho)",
            data=file_tong.getvalue(),
            file_name="KSNB_Bao_Cao_Tong_Hop.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        if live_preview_data:
            st.error(f"🔍 XEM TRƯỚC {tong_loi} CUỘN BỊ LỖI (Sắp xếp theo độ lệch):")
            df_preview = pd.DataFrame(live_preview_data).sort_values(by="Lệch (Kg)", ascending=False).reset_index(drop=True)
            st.dataframe(df_preview, use_container_width=True)
        else:
            st.success("✅ Chúc mừng! Mọi số liệu tồn kho đều Đạt chuẩn KSNB.")

    except Exception as e:
        st.error(f"Đã xảy ra lỗi hệ thống: {e}")
