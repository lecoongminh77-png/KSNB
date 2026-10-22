import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Báo Cáo NXT NVL Chính",
    page_icon="📊",
    layout="wide"
)

st.title("📊 BÁO CÁO TỔNG HỢP NXT NVL CHÍNH")


# --- 1. HÀM XUẤT EXCEL CHUẨN ĐẸP 100% THEO FORM GỐC NHÀ MÁY ---
def export_to_excel(df_tb1, df_tb2, df_tb3, df_tb4):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "NXT Tong Hop"
    ws.views.sheetView[0].showGridLines = True

    # 1. Kích thước cột & chiều cao dòng chuẩn theo file nhà máy
    col_widths = {
        'A': 6.5, 'B': 3.5, 'C': 41.3, 'D': 15.2, 'E': 17.8, 'F': 25.8, 'G': 5.8, 'H': 5.3, 'I': 3.2,
        'J': 1.5, 'K': 1.3, 'L': 8.8, 'M': 17.7, 'N': 17.0, 'O': 19.8, 'P': 20.5, 'Q': 29.5, 'R': 14.3, 'S': 5.5
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    row_heights = {
        1: 27.75, 2: 16.5, 3: 46.5, 4: 21.75, 5: 16.0, 6: 27.75, 7: 17.0, 8: 23.25, 9: 21.75,
        10: 25.5, 11: 25.5, 12: 25.5, 13: 25.5, 14: 21.75, 15: 17.0, 16: 23.25, 17: 18.0,
        18: 22.5, 19: 21.75, 20: 26.25, 21: 26.25, 22: 26.25, 23: 26.25, 24: 18.0
    }
    for r, h in row_heights.items():
        ws.row_dimensions[r].height = h

    # 2. Gộp ô
    merges = [
        "C1:I2", "M1:S2", "A3:I3", "K3:S3", "A4:I4", "K4:S4",
        "B8:C9", "M8:N9", "Q8:Q9",
        "B10:C10", "M10:N10", "B11:C11", "M11:N11",
        "B12:C12", "M12:N12", "B13:C13", "M13:N13", "B14:C14", "M14:N14",
        "B18:C19", "D18:E18", "M18:N19", "P18:P19", "D19:E19",
        "D20:E20", "M20:N20", "D21:E21", "M21:N21",
        "B22:C22", "D22:E22", "M22:N22", "D23:E23", "M23:N23"
    ]
    for m in merges:
        ws.merge_cells(m)

    # 3. Định dạng Font & Màu
    f_factory = Font(name="Times New Roman", size=16, bold=True)
    f_title = Font(name="Times New Roman", size=21, bold=True)
    f_date = Font(name="Times New Roman", size=12, bold=True, color="FF0000")
    f_section = Font(name="Times New Roman", size=13.5, bold=True, color="FF0000")
    f_header = Font(name="Times New Roman", size=13, bold=True)
    f_subhead = Font(name="Times New Roman", size=13, bold=False)
    f_data = Font(name="Times New Roman", size=13, bold=True)
    f_total = Font(name="Times New Roman", size=13.5, bold=True, color="FF0000")

    fill_header = PatternFill(start_color="C5D9F1", end_color="C5D9F1", fill_type="solid")
    fill_total = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    al_center = Alignment(horizontal="center", vertical="center")
    al_left = Alignment(horizontal="left", vertical="center")
    al_right = Alignment(horizontal="right", vertical="center")

    thin_border = Side(border_style="thin", color="000000")
    border_cell = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    medium_border = Side(border_style="medium", color="000000")

    # 4. Tiêu đề sổ sách
    ws["C1"] = "NHÀ MÁY TÔN HOA SEN PHÚ MỸ"; ws["C1"].font = f_factory; ws["C1"].alignment = al_center
    ws["M1"] = "NHÀ MÁY TÔN HOA SEN PHÚ MỸ"; ws["M1"].font = f_factory; ws["M1"].alignment = al_center
    ws["A3"] = "BÁO CÁO KHỐI LƯỢNG NLC, NPL TẠI CTC/NM"; ws["A3"].font = f_title; ws["A3"].alignment = al_center
    ws["K3"] = "BÁO CÁO KHỐI LƯỢNG SỬ DỤNG NLC, NPL TẠI CTC/NM"; ws["K3"].font = f_title; ws["K3"].alignment = al_center
    ws["A4"] = "Đến 23 giờ 59 phút ngày báo cáo"; ws["A4"].font = f_date; ws["A4"].alignment = al_center
    ws["K4"] = "Đến 23 giờ 59 phút ngày báo cáo"; ws["K4"].font = f_date; ws["K4"].alignment = al_center

    ws["A6"] = "1. Tồn kho HRC-PO-CR-CM"; ws["A6"].font = f_section; ws["A6"].alignment = al_left
    ws["K6"] = "3. Tình hình sử dụng NVL chính :"; ws["K6"].font = f_section; ws["K6"].alignment = al_left

    # 5. Header Bảng 1 & 3
    for prefix, start_c in [("", "B"), ("", "M")]:
        p_q = "D" if start_c == "B" else "O"
        p_w = "E" if start_c == "B" else "P"
        p_n = "F" if start_c == "B" else "Q"
        ws[f"{start_c}8"] = "Tên hàng"; ws[f"{start_c}8"].font = f_header; ws[f"{start_c}8"].alignment = al_center
        ws[f"{p_q}8"] = "Số lượng"; ws[f"{p_q}8"].font = f_header; ws[f"{p_q}8"].alignment = al_center
        ws[f"{p_q}9"] = "Cuộn"; ws[f"{p_q}9"].font = f_subhead; ws[f"{p_q}9"].alignment = al_center
        ws[f"{p_w}8"] = "Khối lượng"; ws[f"{p_w}8"].font = f_header; ws[f"{p_w}8"].alignment = al_center
        ws[f"{p_w}9"] = "Tấn"; ws[f"{p_w}9"].font = f_subhead; ws[f"{p_w}9"].alignment = al_center
        ws[f"{p_n}8"] = "Ghi chú"; ws[f"{p_n}8"].font = f_header; ws[f"{p_n}8"].alignment = al_center

    for r in [8, 9]:
        for c in ["B", "C", "D", "E", "F", "M", "N", "O", "P", "Q"]:
            ws[f"{c}{r}"].fill = fill_header
            ws[f"{c}{r}"].border = border_cell

    # 6. Đổ dữ liệu Bảng 1
    for i, r in df_tb1.iterrows():
        idx = 10 + i
        is_tot = (r['Tên hàng'] == 'Tổng')
        ws[f"B{idx}"] = r['Tên hàng']; ws[f"B{idx}"].font = f_total if is_tot else f_data; ws[f"B{idx}"].alignment = al_center if is_tot else al_left
        ws[f"D{idx}"] = r['Số lượng (Cuộn)']; ws[f"D{idx}"].font = f_total if is_tot else f_data; ws[f"D{idx}"].alignment = al_center; ws[f"D{idx}"].number_format = "#,##0"
        ws[f"E{idx}"] = r['Khối lượng (Tấn)']; ws[f"E{idx}"].font = f_total if is_tot else f_data; ws[f"E{idx}"].alignment = al_center; ws[f"E{idx}"].number_format = "#,##0.00"
        for c in ["B", "C", "D", "E", "F"]:
            ws[f"{c}{idx}"].border = border_cell
            if is_tot: ws[f"{c}{idx}"].fill = fill_total

    # 7. Đổ dữ liệu Bảng 3
    for i, r in df_tb3.iterrows():
        idx = 10 + i
        is_tot = (r['Tên hàng'] == 'Tổng')
        ws[f"M{idx}"] = r['Tên hàng']; ws[f"M{idx}"].font = f_total if is_tot else f_data; ws[f"M{idx}"].alignment = al_center if is_tot else al_left
        if pd.notnull(r['Số lượng (Cuộn)']):
            ws[f"O{idx}"] = r['Số lượng (Cuộn)']; ws[f"O{idx}"].number_format = "#,##0"
        ws[f"O{idx}"].font = f_total if is_tot else f_data; ws[f"O{idx}"].alignment = al_center
        ws[f"P{idx}"] = r['Khối lượng (Tấn)']; ws[f"P{idx}"].font = f_total if is_tot else f_data; ws[f"P{idx}"].alignment = al_center; ws[f"P{idx}"].number_format = "#,##0.00"
        for c in ["M", "N", "O", "P", "Q"]:
            ws[f"{c}{idx}"].border = border_cell
            if is_tot: ws[f"{c}{idx}"].fill = fill_total

    # 8. Header Bảng 2 & 4
    ws["A16"] = "2. Kim loại màu tại nhà máy NH"; ws["A16"].font = f_section; ws["A16"].alignment = al_left
    ws["K16"] = "4. Tình hình sử dụng NPL :"; ws["K16"].font = f_section; ws["K16"].alignment = al_left

    ws["B18"] = "Tên hàng"; ws["B18"].font = f_header; ws["B18"].alignment = al_center
    ws["D18"] = "Khối lượng"; ws["D18"].font = f_header; ws["D18"].alignment = al_center
    ws["D19"] = "Tấn"; ws["D19"].font = f_subhead; ws["D19"].alignment = al_center
    ws["F18"] = "Ghi chú"; ws["F18"].font = f_header; ws["F18"].alignment = al_center

    ws["M18"] = "Tên hàng"; ws["M18"].font = f_header; ws["M18"].alignment = al_center
    ws["O18"] = "Khối lượng"; ws["O18"].font = f_header; ws["O18"].alignment = al_center
    ws["O19"] = "Tấn"; ws["O19"].font = f_subhead; ws["O19"].alignment = al_center
    ws["P18"] = "Ghi chú"; ws["P18"].font = f_header; ws["P18"].alignment = al_center

    for r in [18, 19]:
        for c in ["B", "C", "D", "E", "F", "M", "N", "O", "P"]:
            ws[f"{c}{r}"].fill = fill_header
            ws[f"{c}{r}"].border = border_cell

    # 9. Đổ dữ liệu Bảng 2 & 4
    for i, r in df_tb2.iterrows():
        idx = 20 + i
        ws[f"B{idx}"] = r['Tên hàng']; ws[f"B{idx}"].font = f_data; ws[f"B{idx}"].alignment = al_left
        ws[f"D{idx}"] = r['Khối lượng (Tấn)']; ws[f"D{idx}"].font = f_data; ws[f"D{idx}"].alignment = al_right; ws[f"D{idx}"].number_format = "#,##0.00"
        for c in ["B", "C", "D", "E", "F"]:
            ws[f"{c}{idx}"].border = border_cell

    for i, r in df_tb4.iterrows():
        idx = 20 + i
        ws[f"M{idx}"] = r['Tên hàng']; ws[f"M{idx}"].font = f_data; ws[f"M{idx}"].alignment = al_left
        ws[f"O{idx}"] = r['Khối lượng (Tấn)']; ws[f"O{idx}"].font = f_data; ws[f"O{idx}"].alignment = al_right; ws[f"O{idx}"].number_format = "#,##0.00"
        for c in ["M", "N", "O", "P"]:
            ws[f"{c}{idx}"].border = border_cell

    # 10. Đóng khung viền đậm hai bên
    def apply_box(min_c, max_c, min_r, max_r):
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                col_let = get_column_letter(c)
                cell = ws[f"{col_let}{r}"]
                top = medium_border if r == min_r else cell.border.top
                bot = medium_border if r == max_r else cell.border.bottom
                left = medium_border if c == min_c else cell.border.left
                right = medium_border if c == max_c else cell.border.right
                cell.border = Border(top=top, bottom=bot, left=left, right=right)

    apply_box(1, 9, 1, 24)
    apply_box(11, 19, 1, 24)

    wb.save(output)
    return output.getvalue()

# --- 2. HÀM XỬ LÝ INV 110 (TỒN KHO) ---
@st.cache_data
def process_inv110(files):
    df_list = []
    for f in files:
        try:
            df = pd.read_excel(f, skiprows=10)
            if len(df.columns) >= 24:
                sub_df = pd.DataFrame({
                    'PhanLoai': df.iloc[:, 0].astype(str).str.strip(),
                    'Kho': df.iloc[:, 1].astype(str).str.strip(),
                    'KhoiLuong': pd.to_numeric(df.iloc[:, 23], errors='coerce').fillna(0)
                })
                df_list.append(sub_df)
        except Exception as e:
            st.error(f"Lỗi đọc file INV 110: {e}")

    if not df_list:
        return None, None

    data = pd.concat(df_list, ignore_index=True)
    data = data[~data['PhanLoai'].isin(['nan', 'Phân loại', 'None'])]

    # Bảng 1: Thép cuộn
    hrc_thuong = data[(data['PhanLoai'] == 'HRC') & (~data['Kho'].str.contains('035', na=False))]
    hrc_cang = data[(data['PhanLoai'] == 'HRC') & (data['Kho'].str.contains('035', na=False))]
    po = data[data['PhanLoai'] == 'PO']
    cm_cr = data[data['PhanLoai'].isin(['CM', 'CR'])]

    tb1 = pd.DataFrame({
        'Tên hàng': ['Tồn kho HRC', 'Tồn kho HRC ̣̣̣( hàng ngoài cảng)', 'Tồn kho PO', 'Tồn kho CM + CR'],
        'Số lượng (Cuộn)': [len(hrc_thuong), len(hrc_cang), len(po), len(cm_cr)],
        'Khối lượng (Tấn)': [
            hrc_thuong['KhoiLuong'].sum() / 1000,
            hrc_cang['KhoiLuong'].sum() / 1000,
            po['KhoiLuong'].sum() / 1000,
            cm_cr['KhoiLuong'].sum() / 1000
        ],
        'Ghi chú': ['', '', '', '']
    })
    tb1.loc[len(tb1)] = ['Tổng', tb1['Số lượng (Cuộn)'].sum(), tb1['Khối lượng (Tấn)'].sum(), '']

    # Bảng 2: Kim loại màu
    cgg = data[data['PhanLoai'].str.contains('Kẽm không hợp kim|CGG', na=False, case=False)]
    nhom_kem = data[data['PhanLoai'].str.contains('nhôm kẽm', na=False, case=False)]
    magie = data[data['PhanLoai'].str.contains('Magie', na=False, case=False)]
    pha_che = data[data['PhanLoai'].str.contains('pha chế', na=False, case=False)]

    tb2 = pd.DataFrame({
        'Tên hàng': ['Tồn kho CGG', 'Tồn kho Hợp kim nhôm kẽm', 'Tồn kho Hợp kim kẽm Magie', 'Tồn kho nguyên liệu pha chế'],
        'Khối lượng (Tấn)': [
            cgg['KhoiLuong'].sum() / 1000,
            nhom_kem['KhoiLuong'].sum() / 1000,
            magie['KhoiLuong'].sum() / 1000,
            pha_che['KhoiLuong'].sum() / 1000
        ],
        'Ghi chú': ['', '', '', '']
    })

    return tb1, tb2


# --- 3. HÀM XỬ LÝ 105 (GIAO DỊCH SỬ DỤNG) ---
@st.cache_data
def process_105(files):
    df_list = []
    for f in files:
        try:
            df = pd.read_excel(f, skiprows=8)
            if len(df.columns) >= 12:
                sub_df = pd.DataFrame({
                    'ColA': df.iloc[:, 0].astype(str).str.strip(),
                    'ColB': df.iloc[:, 1].astype(str).str.strip(),
                    'LoaiGD': df.iloc[:, 2].astype(str).str.strip(),
                    'TenHang': df.iloc[:, 6].astype(str).str.strip(),
                    'KhoiLuong': pd.to_numeric(df.iloc[:, 11], errors='coerce').fillna(0)
                })
                df_list.append(sub_df)
        except Exception as e:
            st.error(f"Lỗi đọc file 105: {e}")

    if not df_list:
        return None, None

    data = pd.concat(df_list, ignore_index=True)
    has_custom_cols = data['ColA'].isin(['HRC', 'PO', 'CM', 'A']).any()

    if has_custom_cols:
        def get_nvl_stats(item_name):
            issue_cnt = len(data[(data['ColA'] == item_name) & (data['ColB'] == '1') & (data['LoaiGD'] == 'WIP Issue')])
            return_cnt = len(data[(data['ColA'] == item_name) & (data['ColB'] == '1') & (data['LoaiGD'] == 'WIP Return')])
            qty = issue_cnt - return_cnt
            weight = abs(data[(data['ColA'] == item_name) & (data['ColB'] == '1')]['KhoiLuong'].sum()) / 1000
            return qty, weight

        hrc_q, hrc_w = get_nvl_stats('HRC')
        po_q, po_w = get_nvl_stats('PO')
        cm_q, cm_w = get_nvl_stats('CM')
        scrap_w = data[(data['ColB'] == '2') & (data['ColA'] == 'A')]['KhoiLuong'].sum() / 1000

        def get_npl_weight(npl_name):
            w3 = data[(data['ColA'].str.contains(npl_name, na=False, case=False)) & (data['ColB'] == '3')]['KhoiLuong'].sum()
            w4 = data[(data['ColA'].str.contains(npl_name, na=False, case=False)) & (data['ColB'] == '4')]['KhoiLuong'].sum()
            return abs(w3 - w4) / 1000

        nhom_kem_w = get_npl_weight('Hợp kim nhôm kẽm')
        kem_w = get_npl_weight('Kẽm không hợp kim')
        magie_w = get_npl_weight('Magie')
        pha_che_w = get_npl_weight('pha chế')

    else:
        def calc_raw(pattern):
            df_item = data[data['TenHang'].str.contains(pattern, na=False, case=False)]
            issue = df_item[df_item['LoaiGD'] == 'WIP Issue']
            ret = df_item[df_item['LoaiGD'] == 'WIP Return']
            qty = len(issue) - len(ret)
            weight = abs(issue['KhoiLuong'].sum() + ret['KhoiLuong'].sum()) / 1000
            return qty, weight

        hrc_q, hrc_w = calc_raw('HRC|Cán nóng')
        po_q, po_w = calc_raw('PO|Tẩy rỉ')
        cm_q, cm_w = calc_raw('CM|Cán nguội')
        
        scrap = data[data['LoaiGD'].isin(['WIP Byproduct Completion', 'WIP Byproduct Return'])]
        scrap_w = abs(scrap['KhoiLuong'].sum()) / 1000

        def calc_raw_npl(pattern):
            df_npl = data[data['TenHang'].str.contains(pattern, na=False, case=False)]
            return abs(df_npl['KhoiLuong'].sum()) / 1000

        nhom_kem_w = calc_raw_npl('nhôm kẽm')
        kem_w = calc_raw_npl('Kẽm')
        magie_w = calc_raw_npl('Magie')
        pha_che_w = calc_raw_npl('pha chế')

    tb3 = pd.DataFrame({
        'Tên hàng': ['HRC', 'PO', 'CM', 'Hao hụt+ phế liệu'],
        'Số lượng (Cuộn)': [hrc_q, po_q, cm_q, np.nan],
        'Khối lượng (Tấn)': [hrc_w, po_w, cm_w, scrap_w],
        'Ghi chú': ['', '', '', '']
    })
    tb3.loc[len(tb3)] = ['Tổng', tb3['Số lượng (Cuộn)'].sum(), tb3['Khối lượng (Tấn)'].sum(), '']

    tb4 = pd.DataFrame({
        'Tên hàng': ['Hợp kim nhôm kẽm', 'Kẽm không hợp kim', 'Hợp kim kẽm Magie', 'Nguyên liệu pha chế'],
        'Khối lượng (Tấn)': [nhom_kem_w, kem_w, magie_w, pha_che_w],
        'Ghi chú': ['', '', '', '']
    })

    return tb3, tb4


# --- 4. GIAO DIỆN HƯỚNG DẪN & UPLOAD ---
st.markdown("### 📥 Nạp Dữ Liệu Nguồn")

with st.expander("📖 Bấm vào đây để xem Hướng dẫn lấy file từ hệ thống"):
    col_huongdan1, col_huongdan2 = st.columns(2)
    with col_huongdan1:
        st.markdown("""
        **INV 110**
        INV_110_GD2
        
        **HSG_KS_INV_Cung Ứng**
        - 031,035,03G,03H,041,046
        - Chọn mốc thời gian ngày thứ 5 tuần báo cáo

        **HSG_KS_INV_Vật tư**
        - 03B
        - Chọn mốc thời gian ngày thứ 5 tuần báo cáo

        **NMBD_KS_PO_ADMIN**
        - 047
        - Chọn mốc thời gian ngày thứ 5 tuần báo cáo
        - Xóa dữ liệu (Hợp kim kẽm nhôm magie (dạng xỉ))
        """)

    with col_huongdan2:
        st.markdown("""
        **OMP 105**
        HSG_OPM_105_BCKiemSoatThongTinGDTheoLo_GD2
        
        **HSG_KS_INV_Cung Ứng**
        - 031,03G,03H,046
        - Loại nghiệp vụ: Inventory, Job or Schedule, Move order
        - Chọn mốc thời gian từ thứ 6 tuần trước đến thứ 5 tuần này.

        **HSG_KS_INV_Vật tư**
        - 03B
        - Loại nghiệp vụ: Inventory, Job or Schedule, Move order
        - Chọn mốc thời gian từ thứ 6 tuần trước đến thứ 5 tuần này.

        **NMBD_KS_PO_ADMIN**
        - 047
        - Loại nghiệp vụ: Inventory, Job or Schedule, Move order
        - Chọn mốc thời gian từ thứ 6 tuần trước đến thứ 5 tuần này.
        """)

col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    inv110_files = st.file_uploader(
        "1️⃣ Tải lên các file INV 110", 
        type=["xlsx"], 
        accept_multiple_files=True,
        help="Nạp cùng lúc các file INV 110 lấy từ Cung ứng, Vật tư, Admin."
    )

with col_upload2:
    file105_files = st.file_uploader(
        "2️⃣ Tải lên các file 105 (Giao dịch sản xuất)", 
        type=["xlsx"], 
        accept_multiple_files=True,
        help="Nạp cùng lúc các file 105 lấy từ Cung ứng, Vật tư, Admin."
    )


# --- 5. HIỂN THỊ KẾT QUẢ VÀ NÚT TẢI EXCEL ---
if inv110_files and file105_files:
    st.divider()

    def style_table(df, is_weight_only=False):
        if is_weight_only:
            return df.style.format({"Khối lượng (Tấn)": "{:,.2f}"}).apply(
                lambda r: ['background-color: #ffff00; font-weight: bold; color: red' if r['Tên hàng'] == 'Tổng' else '' for _ in r], axis=1
            )
        else:
            return df.style.format({
                "Số lượng (Cuộn)": lambda x: f"{x:,.0f}" if pd.notnull(x) else "",
                "Khối lượng (Tấn)": "{:,.2f}"
            }).apply(
                lambda r: ['background-color: #ffff00; font-weight: bold; color: red' if r['Tên hàng'] == 'Tổng' else '' for _ in r], axis=1
            )

    df_tb1, df_tb2 = process_inv110(inv110_files)
    df_tb3, df_tb4 = process_105(file105_files)

    if df_tb1 is not None and df_tb3 is not None:
        col_left, col_right = st.columns(2)

        with col_left:
            st.markdown("<h4 style='text-align: center; color: #b30000;'>BÁO CÁO KHỐI LƯỢNG NLC, NPL TẠI CTC/NM</h4>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center; font-style: italic; font-size: 13px;'>Đến 23 giờ 59 phút ngày báo cáo</p>", unsafe_allow_html=True)
            
            st.markdown("<b style='color: red;'>1. Tồn kho HRC-PO-CR-CM</b>", unsafe_allow_html=True)
            st.dataframe(style_table(df_tb1), use_container_width=True, hide_index=True)

            st.markdown("<b style='color: red;'>2. Kim loại màu tại nhà máy NH</b>", unsafe_allow_html=True)
            st.dataframe(style_table(df_tb2, is_weight_only=True), use_container_width=True, hide_index=True)

        with col_right:
            st.markdown("<h4 style='text-align: center; color: #b30000;'>BÁO CÁO KHỐI LƯỢNG SỬ DỤNG NLC, NPL TẠI CTC/NM</h4>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center; font-style: italic; font-size: 13px;'>Đến 23 giờ 59 phút ngày báo cáo</p>", unsafe_allow_html=True)
            
            st.markdown("<b style='color: red;'>3. Tình hình sử dụng NVL chính :</b>", unsafe_allow_html=True)
            st.dataframe(style_table(df_tb3), use_container_width=True, hide_index=True)

            st.markdown("<b style='color: red;'>4. Tình hình sử dụng NPL :</b>", unsafe_allow_html=True)
            st.dataframe(style_table(df_tb4, is_weight_only=True), use_container_width=True, hide_index=True)

        # NÚT XUẤT VÀ TẢI VỀ FILE EXCEL ĐẸP
        st.write("")
        excel_data = export_to_excel(df_tb1, df_tb2, df_tb3, df_tb4)
        st.download_button(
            label="📥 Tải Báo Cáo Ra File Excel (.xlsx)",
            data=excel_data,
            file_name="BAO_CAO_NXT_TONG_HOP.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
