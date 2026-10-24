import io
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
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
    .stApp { background-color: #f5f5f7; }
    div.stButton > button { background: #0071e3 !important; color: white !important; border-radius: 8px !important; border: none !important; font-weight: bold; }
    div.stButton > button:hover { background: #005bb5 !important; transform: scale(1.02); }
</style>
""", unsafe_allow_html=True)

st.title("📦 HỆ THỐNG KSNB TỒN KHO (BẢN GỘP FILE)")

col_f1, col_f2 = st.columns(2)
with col_f1: 
    f_tonkhos = st.file_uploader("1. Nạp các File 110 (Kéo thả nhiều file)", type=["xlsx", "xlsb", "xls"], accept_multiple_files=True)
with col_f2: 
    f_doichieu = st.file_uploader("2. Nạp File RMA 105 (Tuỳ chọn)", type=["xlsx", "xlsb", "xls"])

if st.button("🚀 Bắt đầu Tổng hợp & Kiểm toán", use_container_width=True) and f_tonkhos:
    progress = st.info("Đang xử lý dữ liệu siêu tốc... Vui lòng đợi!")
    
    try:
        # 1. ĐỌC FILE 105 RMA
        set_rma = set()
        if f_doichieu:
            df_rma = pd.read_excel(f_doichieu, header=None)
            for idx, row in df_rma.iterrows():
                val = str(row.iloc[5]).replace(".0", "").strip().upper()
                if val and val != "NAN" and val != "MÃ LÔ": set_rma.add(val)

        # 2. CHUẨN BỊ TẠO FILE BÁO CÁO MỚI TINH
        wb_tong = openpyxl.Workbook()
        ws_tong = wb_tong.active
        ws_tong.title = "Bao_Cao_Tong"
        
        wb_loi = openpyxl.Workbook()
        ws_loi = wb_loi.active
        ws_loi.title = "Danh_Sach_Loi"

        # Định dạng
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="005bb5", end_color="005bb5", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        gray_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")

        # Tiêu đề 10 Cột KSNB Mới
        tieu_de_ksnb = [
            "TRẠNG THÁI KSNB", "ĐỘ LỆCH (KG)", "TỶ LỆ LỆCH (%)", 
            "KHOẢNG TRỌNG", "KG CHUẨN (Tham khảo)", "LÕI KHAI BÁO (Phần mềm)", 
            "CĂN CỨ ĐỐI SOÁT", "CHI TIẾT LỖI", "BỘ LỌC KHO", "NGUỒN FILE GỐC"
        ]
        
        headers_goc = []
        du_lieu_tong = []
        du_lieu_loi = []
        tong_da_xoa = 0
        tong_loi = 0
        live_preview_data = []

        # 3. QUÉT DỮ LIỆU TỪ CÁC FILE NẠP VÀO
        for file_idx, f_tk in enumerate(f_tonkhos):
            wb_doc = openpyxl.load_workbook(f_tk, data_only=True)
            ws_doc = wb_doc.active
            max_c = ws_doc.max_column
            max_r = ws_doc.max_row
            
            # Lấy tiêu đề gốc từ file đầu tiên (Trích xuất từ dòng 12)
            if file_idx == 0:
                for c in range(1, max_c + 1):
                    h12 = safe_str(ws_doc.cell(row=12, column=c).value)
                    h13 = safe_str(ws_doc.cell(row=13, column=c).value)
                    # Gộp dòng 12 và 13 để ra tiêu đề dễ hiểu (Ví dụ: "Khối lượng Gross")
                    if h12 and h13 and h12 != h13: h_final = f"{h12} {h13}"
                    elif h12: h_final = h12
                    elif h13: h_final = h13
                    else: h_final = f"Cột {c}"
                    headers_goc.append(h_final)
                    
                # Ghi tiêu đề vào 2 file mới
                full_headers = tieu_de_ksnb + headers_goc
                ws_tong.append(full_headers)
                ws_loi.append(full_headers)
                for cell in ws_tong[1]: cell.font, cell.fill = header_font, header_fill
                for cell in ws_loi[1]: cell.font, cell.fill = header_font, header_fill

            # Đọc dữ liệu từ dòng 14 xuống
            for r in range(14, max_r + 1):
                ma_hang = safe_str(ws_doc.cell(row=r, column=2).value)
                ma_lo = safe_str(ws_doc.cell(row=r, column=20).value).replace(".0", "").upper()
                gross_num = safe_float(ws_doc.cell(row=r, column=22).value)
                
                # --- VÒNG LỌC 1: DỌN RÁC ---
                if not ma_hang and gross_num == 0: continue
                if ma_lo in set_rma and ma_lo != "":
                    tong_da_xoa += 1
                    continue
                if not (ma_hang.startswith('3') or ma_hang.startswith('4')):
                    tong_da_xoa += 1
                    continue

                # Bốc dữ liệu gốc nguyên dòng (để dán sang file mới)
                row_goc = [ws_doc.cell(row=r, column=c).value for c in range(1, max_c + 1)]

                # --- VÒNG LỌC 2: KSNB VÀ BẪY LỖI ERP ---
                locator = safe_str(ws_doc.cell(row=r, column=18).value)
                net_num = safe_float(ws_doc.cell(row=r, column=23).value)
                kg_loi_san = safe_float(ws_doc.cell(row=r, column=24).value)
                loai_loi = safe_str(ws_doc.cell(row=r, column=25).value)

                nhom_kho = "Kho B" if (len(locator) >= 4 and locator[3].upper() == 'B') else "Kho Khác"
                
                # Biến KSNB
                trang_thai = ""
                lech = 0.0
                ty_le = 0.0
                hang = "-"
                ky_chuan = 0
                can_cu = "-"
                ghi_chu_loi = []
                fill_color = None
                is_error = False

                # 1. BẪY LỖI DỮ LIỆU ĐẦU VÀO
                if gross_num <= 0:
                    trang_thai = "Bỏ qua (Trống Gross)"
                    fill_color = gray_fill
                elif (gross_num - net_num) > 500:
                    trang_thai = "Bỏ qua (Cuộn đã chia hàng)"
                    fill_color = gray_fill
                else:
                    # 2. TÍNH TOÁN KSNB BÌNH THƯỜNG
                    has_explicit_core = True if kg_loi_san > 0 else False
                    hang = xac_dinh_hang(gross_num)
                    ky_chuan = xac_dinh_ky_chuan(hang, loai_loi)
                    
                    if has_explicit_core:
                        can_cu = "Lõi khai báo (Phần mềm)"
                        lech = abs((net_num + kg_loi_san) - gross_num)
                        if abs((gross_num - net_num) - kg_loi_san) > 2:
                            ghi_chu_loi.append("Mâu thuẫn Gross/Net/Lõi")
                    else:
                        can_cu = "Ma trận KSNB"
                        lech = abs((gross_num - net_num) - ky_chuan)

                    ty_le = (lech / gross_num * 100) if gross_num > 0 else 0
                    
                    # Bẫy cảnh báo
                    if net_num > gross_num:
                        ghi_chu_loi.append("Lỗi: Net > Gross")
                    elif lech > (gross_num * 0.05):
                        ghi_chu_loi.append("Lỗi: Chênh > 5%")
                    elif lech > 150:
                        ghi_chu_loi.append("Lỗi: Lệch > 150kg")

                    # Chốt trạng thái
                    if ghi_chu_loi or lech > 100:
                        trang_thai = "NGHIÊM TRỌNG"
                        fill_color = red_fill
                        is_error = True
                        tong_loi += 1
                    elif lech > 30:
                        trang_thai = "NGHI NGỜ"
                        fill_color = yellow_fill
                        is_error = True
                        tong_loi += 1
                    else:
                        trang_thai = "ĐẠT"

                # Đóng gói dữ liệu chèn vào file mới
                str_ghi_chu = " | ".join(ghi_chu_loi) if ghi_chu_loi else ("Vượt chuẩn" if (lech > 30 and not is_error) else "")
                
                row_ksnb = [
                    trang_thai, round(lech, 2), f"{round(ty_le, 2)}%", 
                    hang, ky_chuan if ky_chuan > 0 else "-", kg_loi_san if kg_loi_san > 0 else "-", 
                    can_cu, str_ghi_chu, nhom_kho, f_tk.name
                ]
                full_row_data = row_ksnb + row_goc
                
                # Đẩy vào bộ nhớ Ghi File
                ws_tong.append(full_row_data)
                if fill_color:
                    for cell in ws_tong[ws_tong.max_row]: cell.fill = fill_color
                    
                if is_error:
                    ws_loi.append(full_row_data)
                    for cell in ws_loi[ws_loi.max_row]: cell.fill = fill_color
                    live_preview_data.append({
                        "File gốc": f_tk.name, "Kho": nhom_kho, "Mã Lô": ma_lo, 
                        "Trạng Thái": trang_thai, "Lệch (Kg)": round(lech, 2), "Chi Tiết": str_ghi_chu
                    })

        # 4. LƯU 2 FILE EXCEL RA BỘ NHỚ
        file_tong = io.BytesIO()
        wb_tong.save(file_tong)
        wb_tong.close()

        file_loi = io.BytesIO()
        wb_loi.save(file_loi)
        wb_loi.close()

        # 5. HIỂN THỊ KẾT QUẢ LÊN WEB
        progress.empty()
        st.toast("🎉 Phân tích Gộp File thành công!")
        
        c1, c2 = st.columns(2)
        c1.metric("🗑️ Cuộn rác & Khác mã đã dọn", tong_da_xoa)
        c2.metric("⚠️ Số cuộn phát hiện Lỗi/Hao hụt", tong_loi, delta_color="inverse")

        # Nút Tải File
        st.markdown("### 📥 XUẤT BÁO CÁO")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                label="📊 TẢI BÁO CÁO TỔNG (Toàn bộ các kho)",
                data=file_tong.getvalue(),
                file_name="KSNB_Bao_Cao_Tong_Hop.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with d2:
            st.download_button(
                label="🚨 TẢI TRÁT ĐÒI HÀNG (Chỉ chứa cuộn lỗi)",
                data=file_loi.getvalue(),
                file_name="KSNB_Trat_Doi_Hang.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        # Bảng Live Preview
        if live_preview_data:
            st.error(f"🔍 XEM TRƯỚC {tong_loi} CUỘN BỊ LỖI (Sắp xếp theo độ lệch):")
            df_preview = pd.DataFrame(live_preview_data).sort_values(by="Lệch (Kg)", ascending=False).reset_index(drop=True)
            st.dataframe(df_preview, use_container_width=True)
        else:
            st.success("✅ Chúc mừng! Mọi số liệu tồn kho đều Đạt chuẩn KSNB.")

    except Exception as e:
        st.error(f"Đã xảy ra lỗi hệ thống: {e}")
