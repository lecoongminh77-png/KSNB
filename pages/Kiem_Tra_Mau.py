import io
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
import streamlit as st

# --- BẢNG MA TRẬN KSNB (Mở trần mốc >= 17 tấn) ---
MA_TRAN_CHUAN = {
    "Dưới 2 tấn":        {"giay": 40,  "sat": 70,  "khac": 20},
    "Từ 2 - 5 tấn":      {"giay": 60,  "sat": 90,  "khac": 40},
    "Từ 5 - 7 tấn":      {"giay": 70,  "sat": 100, "khac": 50},
    "Từ 7 - 10 tấn":     {"giay": 90,  "sat": 120, "khac": 60},
    "Từ 10 - 12 tấn":    {"giay": 100, "sat": 120, "khac": 70},
    "Từ 12 - 17 tấn":    {"giay": 110, "sat": 130, "khac": 80},
    "Từ 17 tấn trở lên": {"giay": 120, "sat": 140, "khac": 90},
}

# Radar quét từ khóa loại trừ ở cột AK (Đánh tên)
TU_KHOA_LOAI_AK = [
    "THUHOI", "THU HOI", "THU HỒI",
    "LỢP XƯỞNG", "LOP XUONG", "LOPXUONG"
]

def safe_str(val):
    return str(val).strip() if pd.notna(val) and val is not None else ""

def safe_float(val):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "").replace(" ", "")
    try:
        return float(s)
    except:
        return 0.0

def get_cell(row_tuple, idx):
    return row_tuple[idx] if idx < len(row_tuple) else None

def xac_dinh_hang(kl):
    if kl < 2000: return "Dưới 2 tấn"
    elif kl <= 5000: return "Từ 2 - 5 tấn"
    elif kl <= 7000: return "Từ 5 - 7 tấn"
    elif kl <= 10000: return "Từ 7 - 10 tấn"
    elif kl <= 12000: return "Từ 10 - 12 tấn"
    elif kl <= 17000: return "Từ 12 - 17 tấn"
    else: return "Từ 17 tấn trở lên"

def tinh_kg_chuan(hang, loai_loi, kg_loi_khai_bao):
    loai_str = safe_str(loai_loi).lower()
    nhom = MA_TRAN_CHUAN[hang]
    bao_bi_chuan = nhom["khac"]
    
    if "sắt" in loai_str or "sat" in loai_str:
        chuan_mt = nhom["sat"]
    elif "giấy" in loai_str or "giay" in loai_str:
        chuan_mt = nhom["giay"]
    else:
        chuan_mt = bao_bi_chuan

    # Nếu có khai báo lõi dày đặc biệt -> lấy max(Chuẩn ma trận, Lõi khai báo + Bao bì chuẩn)
    if kg_loi_khai_bao > 0:
        return max(chuan_mt, kg_loi_khai_bao + bao_bi_chuan)
    return chuan_mt

def doc_du_lieu_file_110(uploaded_file):
    """Đọc nhanh mọi định dạng xlsx, xls, xlsb sang danh sách dòng"""
    fname = uploaded_file.name.lower()
    if fname.endswith(".xlsx"):
        wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        return rows
    else:
        df_tmp = pd.read_excel(uploaded_file, header=None, dtype=object)
        return [tuple(x) for x in df_tmp.to_numpy()]

# --- GIAO DIỆN STREAMLIT ---
st.set_page_config(page_title="KSNB Tồn Kho", page_icon="📦", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    .stApp { background-color: #f5f5f7; }
    div.stButton > button { background: #0071e3 !important; color: white !important; border-radius: 8px !important; border: none !important; font-weight: bold; }
    div.stButton > button:hover { background: #005bb5 !important; transform: scale(1.02); }
    div.stDownloadButton > button { background: #34c759 !important; color: white !important; border-radius: 8px !important; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Về Trang Chủ", use_container_width=True):
        st.switch_page("app_ksnb.py")

st.title("📦 HỆ THỐNG KSNB TỒN KHO (BẢN NÂNG CẤP CHUYÊN SÂU)")

col_f1, col_f2 = st.columns(2)
with col_f1:
    f_tonkhos = st.file_uploader("1. Nạp các File 110 (Kéo thả nhiều file)", type=["xlsx", "xlsb", "xls"], accept_multiple_files=True)
with col_f2:
    f_doichieu = st.file_uploader("2. Nạp File RMA 105 (Tuỳ chọn)", type=["xlsx", "xlsb", "xls"])

if st.button("🚀 Bắt đầu Tổng hợp & Kiểm toán", use_container_width=True) and f_tonkhos:
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        # 1. ĐỌC FILE 105 RMA (NẾU CÓ)
        set_rma = set()
        if f_doichieu:
            status_text.text("⏳ Đang đọc danh sách Mã Lô từ File RMA 105...")
            df_rma = pd.read_excel(f_doichieu, header=None, dtype=object)
            for _, row in df_rma.iterrows():
                if len(row) > 5:
                    val = safe_str(row.iloc[5]).replace(".0", "").upper()
                    if val and val not in ("NAN", "NONE", "MÃ LÔ"):
                        set_rma.add(val)

        # 2. KHỞI TẠO 1 FILE BÁO CÁO DUY NHẤT
        wb_tong = openpyxl.Workbook()
        ws_tong = wb_tong.active
        ws_tong.title = "Bao_Cao_KSNB"

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="005BB5", end_color="005BB5", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_fill = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")

        tieu_de_ksnb = [
            "TRẠNG THÁI KSNB", "ĐỘ LỆCH KSNB (KG)", "GROSS - NET (KG)",
            "KHOẢNG TRỌNG", "KG CHUẨN", "BỘ LỌC KHO"
        ]

        headers_goc = []
        max_c_global = 0
        tong_da_xoa = 0
        tong_giu_lai = 0
        tong_do = 0
        tong_vang = 0
        live_preview_data = []
        total_files = len(f_tonkhos)

        # 3. DUYỆT TỪNG FILE 110
        for file_idx, f_tk in enumerate(f_tonkhos):
            status_text.text(f"📂 Đang nạp file {f_tk.name} vào bộ nhớ...")
            rows_all = doc_du_lieu_file_110(f_tk)
            max_r = len(rows_all)
            if max_r < 14:
                continue

            # Xây dựng dòng tiêu đề từ dòng 12 (idx 11) và dòng 13 (idx 12) của file đầu tiên
            if not headers_goc:
                r12 = rows_all[11]
                r13 = rows_all[12]
                max_c_global = max(len(r12), len(r13))
                for c in range(max_c_global):
                    h12 = safe_str(get_cell(r12, c))
                    h13 = safe_str(get_cell(r13, c))
                    if h12 and h13 and h12 != h13:
                        h_final = f"{h12} {h13}"
                    elif h12:
                        h_final = h12
                    elif h13:
                        h_final = h13
                    else:
                        h_final = f"Cột {c + 1}"
                    headers_goc.append(h_final)

                ws_tong.append(tieu_de_ksnb + headers_goc)
                for cell in ws_tong[1]:
                    cell.font = header_font
                    cell.fill = header_fill

            total_data_rows = max_r - 13

            # Quét dữ liệu từ dòng 14 (idx 13)
            for idx in range(13, max_r):
                if (idx - 13) % 500 == 0 or idx == max_r - 1:
                    row_prog = (idx - 12) / total_data_rows
                    overall_prog = (file_idx + row_prog) / total_files
                    progress_bar.progress(min(overall_prog, 1.0))
                    status_text.text(f"🔄 Đang kiểm toán: {f_tk.name} (Dòng {idx + 1}/{max_r}) - {round(overall_prog * 100)}%")

                row_vals = rows_all[idx]
                ma_hang = safe_str(get_cell(row_vals, 1))      # Cột B: Mã hàng
                gross_num = safe_float(get_cell(row_vals, 21)) # Cột V: Gross

                # LUẬT 1: Bỏ dòng trống và xóa toàn bộ cuộn trống Gross (Gross <= 0)
                if not ma_hang and gross_num <= 0:
                    continue
                if gross_num <= 0:
                    tong_da_xoa += 1
                    continue

                # LUẬT 2: Chỉ giữ mã hàng bắt đầu bằng 3 hoặc 4
                if not (ma_hang.startswith("3") or ma_hang.startswith("4")):
                    tong_da_xoa += 1
                    continue

                # LUẬT 3: Đối chiếu loại trừ Mã Lô nằm trong File RMA 105
                ma_lo = safe_str(get_cell(row_vals, 19)).replace(".0", "").upper() # Cột T: Mã lô
                if ma_lo and ma_lo in set_rma:
                    tong_da_xoa += 1
                    continue

                # LUẬT 4: Radar quét cột AK (Cột 37 -> idx 36: Đánh tên) loại hàng Thu hồi / Lợp xưởng
                danh_ten_ak = safe_str(get_cell(row_vals, 36)).upper()
                if danh_ten_ak and any(kw in danh_ten_ak for kw in TU_KHOA_LOAI_AK):
                    tong_da_xoa += 1
                    continue

                # Trích xuất thông tin kiểm toán
                ten_hang = safe_str(get_cell(row_vals, 2))      # Cột C: Tên hàng
                phan_loai = safe_str(get_cell(row_vals, 4)).upper() # Cột E: Phân loại (1, 2, CXL...)
                locator = safe_str(get_cell(row_vals, 17))      # Cột R: Locator
                net_num = safe_float(get_cell(row_vals, 22))    # Cột W: Net
                kg_loi_san = safe_float(get_cell(row_vals, 23)) # Cột X: Lõi khai báo
                loai_loi = safe_str(get_cell(row_vals, 24))     # Cột Y: Loại lõi

                bo_loc_kho = locator[3].upper() if len(locator) >= 4 else "-"
                gross_net = round(gross_num - net_num, 2)
                hang = xac_dinh_hang(gross_num)
                kg_chuan = tinh_kg_chuan(hang, loai_loi, kg_loi_san)
                
                # Giữ nguyên dấu (+/-): Dương = Dư bì, Âm = Thiếu bì
                lech_co_dau = round(gross_net - kg_chuan, 2)
                abs_lech = abs(lech_co_dau)
                tag_cxl = " - Hàng CXL" if phan_loai == "CXL" else ""

                fill_color = None
                is_error = False

                # --- HỆ THỐNG PHÂN LOẠI TRẠNG THÁI NÂNG CẤP ---
                if gross_num < net_num:
                    trang_thai = f"NGHIÊM TRỌNG (Gross < Net{tag_cxl})"
                    fill_color = red_fill
                    is_error = True
                    tong_do += 1
                elif gross_num == net_num:
                    if "thép dày" in ten_hang.lower() and kg_loi_san == 0:
                        trang_thai = "ĐẠT"
                        lech_co_dau = 0.0
                    else:
                        trang_thai = f"NGHIÊM TRỌNG (Lỗi Gross = Net{tag_cxl})"
                        fill_color = red_fill
                        is_error = True
                        tong_do += 1
                elif kg_loi_san > 0 and gross_net < kg_loi_san:
                    trang_thai = f"NGHIÊM TRỌNG (Bì < Lõi khai báo{tag_cxl})"
                    fill_color = red_fill
                    is_error = True
                    tong_do += 1
                elif gross_net > 250:
                    trang_thai = f"NGHIÊM TRỌNG (Nghi chưa cập nhật Gross{tag_cxl})"
                    fill_color = red_fill
                    is_error = True
                    tong_do += 1
                elif abs_lech > 100:
                    trang_thai = f"NGHIÊM TRỌNG (Lệch > 100kg{tag_cxl})"
                    fill_color = red_fill
                    is_error = True
                    tong_do += 1
                elif lech_co_dau > 30:
                    trang_thai = f"NGHI NGỜ (Dư bì > 30kg{tag_cxl})"
                    fill_color = yellow_fill
                    is_error = True
                    tong_vang += 1
                elif lech_co_dau < -30:
                    if kg_loi_san > 0 and gross_net == kg_loi_san:
                        trang_thai = f"NGHI NGỜ (Chưa cộng bao bì{tag_cxl})"
                    else:
                        trang_thai = f"NGHI NGỜ (Thiếu bì > 30kg{tag_cxl})"
                    fill_color = yellow_fill
                    is_error = True
                    tong_vang += 1
                else:
                    trang_thai = "ĐẠT"

                tong_giu_lai += 1
                row_ksnb = [
                    trang_thai, lech_co_dau, gross_net,
                    hang, kg_chuan, bo_loc_kho
                ]
                row_goc = [get_cell(row_vals, c) for c in range(max_c_global)]
                ws_tong.append(row_ksnb + row_goc)

                if fill_color:
                    for cell in ws_tong[ws_tong.max_row]:
                        cell.fill = fill_color

                if is_error:
                    live_preview_data.append({
                        "Bộ Lọc Kho": bo_loc_kho,
                        "Mã Lô": ma_lo,
                        "Trạng Thái KSNB": trang_thai,
                        "Độ Lệch (Kg)": lech_co_dau,
                        "Gross - Net": gross_net,
                        "Kg Chuẩn": kg_chuan,
                        "Tên Hàng": ten_hang
                    })

        # Bật sẵn AutoFilter & cố định dòng tiêu đề trên Excel xuất ra
        if ws_tong.max_row >= 1:
            ws_tong.auto_filter.ref = ws_tong.dimensions
            ws_tong.freeze_panes = "A2"

        # 4. ĐÓNG GÓI FILE EXCEL DUY NHẤT
        status_text.text("📦 Đang đóng gói Báo cáo Excel Tổng hợp...")
        file_tong = io.BytesIO()
        wb_tong.save(file_tong)
        wb_tong.close()

        # 5. HIỂN THỊ KẾT QUẢ
        progress_bar.empty()
        status_text.empty()
        st.toast("🎉 Kiểm toán hoàn tất!")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("✅ Tổng cuộn đưa vào báo cáo", f"{tong_giu_lai:,}")
        c2.metric("🗑️ Dòng rác / Thu hồi đã dọn", f"{tong_da_xoa:,}")
        c3.metric("🔴 Lỗi Nghiêm Trọng (Đỏ)", f"{tong_do:,}", delta_color="inverse")
        c4.metric("🟡 Nghi Ngờ (Vàng)", f"{tong_vang:,}", delta_color="off")

        st.download_button(
            label="📊 TẢI BÁO CÁO KSNB TỔNG HỢP (1 FILE DUY NHẤT)",
            data=file_tong.getvalue(),
            file_name="KSNB_Bao_Cao_Tong_Hop.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        if live_preview_data:
            st.warning(f"🔍 XEM TRƯỚC {len(live_preview_data)} CUỘN BẤT THƯỜNG (Sắp xếp theo mức độ lệch):")
            df_preview = pd.DataFrame(live_preview_data)
            df_preview["abs_sort"] = df_preview["Độ Lệch (Kg)"].abs()
            df_preview = df_preview.sort_values(by="abs_sort", ascending=False).drop(columns=["abs_sort"]).reset_index(drop=True)
            st.dataframe(df_preview, use_container_width=True)
        else:
            st.success("✅ Tất cả các cuộn tồn kho đều ĐẠT chuẩn KSNB!")

    except Exception as e:
        st.error(f"Đã xảy ra lỗi hệ thống: {e}")
