import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def export_to_excel(df_tb1, df_tb2, df_tb3, df_tb4):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "NXT Tong Hop"
    ws.views.sheetView[0].showGridLines = True

    # Định dạng Font & Màu sắc chuẩn
    font_factory = Font(name="Times New Roman", size=11, bold=True)
    font_title = Font(name="Times New Roman", size=13, bold=True)
    font_date = Font(name="Times New Roman", size=10, italic=True, color="FF0000")
    font_section = Font(name="Times New Roman", size=10, bold=True, color="CC0000")
    font_header = Font(name="Times New Roman", size=10, bold=True)
    font_data = Font(name="Times New Roman", size=10)
    font_total = Font(name="Times New Roman", size=10, bold=True)

    fill_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_total = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    thin_border = Side(border_style="thin", color="000000")
    border_cell = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # 1. Tiêu đề
    ws.merge_cells("A1:G1"); ws["A1"] = "NHÀ MÁY TÔN HOA SEN PHÚ MỸ"; ws["A1"].font = font_factory; ws["A1"].alignment = align_center
    ws.merge_cells("K1:Q1"); ws["K1"] = "NHÀ MÁY TÔN HOA SEN PHÚ MỸ"; ws["K1"].font = font_factory; ws["K1"].alignment = align_center

    ws.merge_cells("A3:G3"); ws["A3"] = "BÁO CÁO KHỐI LƯỢNG NLC, NPL TẠI CTC/NM"; ws["A3"].font = font_title; ws["A3"].alignment = align_center
    ws.merge_cells("K3:Q3"); ws["K3"] = "BÁO CÁO KHỐI LƯỢNG SỬ DỤNG NLC, NPL TẠI CTC/NM"; ws["K3"].font = font_title; ws["K3"].alignment = align_center

    ws.merge_cells("A4:G4"); ws["A4"] = "Đến 23 giờ 59 phút ngày báo cáo"; ws["A4"].font = font_date; ws["A4"].alignment = align_center
    ws.merge_cells("K4:Q4"); ws["K4"] = "Đến 23 giờ 59 phút ngày báo cáo"; ws["K4"].font = font_date; ws["K4"].alignment = align_center

    # 2. Header Mục 1 & 3
    ws["A6"] = "1. Tồn kho HRC-PO-CR-CM"; ws["A6"].font = font_section
    ws["K6"] = "3. Tình hình sử dụng NVL chính :"; ws["K6"].font = font_section

    # Tiêu đề cột Bảng 1 & 3
    for prefix, start_c in [("", "B"), ("", "M")]:
        p_qty = "D" if start_c == "B" else "O"
        p_wt = "E" if start_c == "B" else "P"
        p_note = "F" if start_c == "B" else "Q"
        
        ws[f"{start_c}8"] = "Tên hàng"; ws[f"{start_c}8"].font = font_header; ws[f"{start_c}8"].fill = fill_header; ws[f"{start_c}8"].alignment = align_center
        ws[f"{p_qty}8"] = "Số lượng"; ws[f"{p_qty}8"].font = font_header; ws[f"{p_qty}8"].fill = fill_header; ws[f"{p_qty}8"].alignment = align_center
        ws[f"{p_qty}9"] = "Cuộn"; ws[f"{p_qty}9"].font = font_header; ws[f"{p_qty}9"].fill = fill_header; ws[f"{p_qty}9"].alignment = align_center
        ws[f"{p_wt}8"] = "Khối lượng"; ws[f"{p_wt}8"].font = font_header; ws[f"{p_wt}8"].fill = fill_header; ws[f"{p_wt}8"].alignment = align_center
        ws[f"{p_wt}9"] = "Tấn"; ws[f"{p_wt}9"].font = font_header; ws[f"{p_wt}9"].fill = fill_header; ws[f"{p_wt}9"].alignment = align_center
        ws[f"{p_note}8"] = "Ghi chú"; ws[f"{p_note}8"].font = font_header; ws[f"{p_note}8"].fill = fill_header; ws[f"{p_note}8"].alignment = align_center

    # Đổ dữ liệu Bảng 1
    for i, r in df_tb1.iterrows():
        row_idx = 10 + i
        is_total = (r['Tên hàng'] == 'Tổng')
        ws[f"B{row_idx}"] = r['Tên hàng']; ws[f"B{row_idx}"].font = font_total if is_total else font_data
        ws[f"D{row_idx}"] = r['Số lượng (Cuộn)']; ws[f"D{row_idx}"].font = font_total if is_total else font_data; ws[f"D{row_idx}"].number_format = "#,##0"
        ws[f"E{row_idx}"] = r['Khối lượng (Tấn)']; ws[f"E{row_idx}"].font = font_total if is_total else font_data; ws[f"E{row_idx}"].number_format = "#,##0.00"
        for c in ["B", "C", "D", "E", "F"]:
            ws[f"{c}{row_idx}"].border = border_cell
            if is_total: ws[f"{c}{row_idx}"].fill = fill_total

    # Đổ dữ liệu Bảng 3
    for i, r in df_tb3.iterrows():
        row_idx = 10 + i
        is_total = (r['Tên hàng'] == 'Tổng')
        ws[f"M{row_idx}"] = r['Tên hàng']; ws[f"M{row_idx}"].font = font_total if is_total else font_data
        if pd.notnull(r['Số lượng (Cuộn)']):
            ws[f"O{row_idx}"] = r['Số lượng (Cuộn)']; ws[f"O{row_idx}"].number_format = "#,##0"
        ws[f"O{row_idx}"].font = font_total if is_total else font_data
        ws[f"P{row_idx}"] = r['Khối lượng (Tấn)']; ws[f"P{row_idx}"].font = font_total if is_total else font_data; ws[f"P{row_idx}"].number_format = "#,##0.00"
        for c in ["M", "N", "O", "P", "Q"]:
            ws[f"{c}{row_idx}"].border = border_cell
            if is_total: ws[f"{c}{row_idx}"].fill = fill_total

    # 3. Header Mục 2 & 4
    ws["A16"] = "2. Kim loại màu tại nhà máy NH"; ws["A16"].font = font_section
    ws["K16"] = "4. Tình hình sử dụng NPL :"; ws["K16"].font = font_section

    # Đổ dữ liệu Bảng 2
    for i, r in df_tb2.iterrows():
        row_idx = 20 + i
        ws[f"B{row_idx}"] = r['Tên hàng']; ws[f"B{row_idx}"].font = font_data
        ws[f"D{row_idx}"] = r['Khối lượng (Tấn)']; ws[f"D{row_idx}"].font = font_data; ws[f"D{row_idx}"].number_format = "#,##0.00"
        for c in ["B", "C", "D", "E", "F"]: ws[f"{c}{row_idx}"].border = border_cell

    # Đổ dữ liệu Bảng 4
    for i, r in df_tb4.iterrows():
        row_idx = 20 + i
        ws[f"M{row_idx}"] = r['Tên hàng']; ws[f"M{row_idx}"].font = font_data
        ws[f"O{row_idx}"] = r['Khối lượng (Tấn)']; ws[f"O{row_idx}"].font = font_data; ws[f"O{row_idx}"].number_format = "#,##0.00"
        for c in ["M", "N", "O", "P", "Q"]: ws[f"{c}{row_idx}"].border = border_cell

    # Căn chỉnh độ rộng cột
    col_widths = {'A': 5, 'B': 20, 'C': 12, 'D': 12, 'E': 15, 'F': 10, 'K': 5, 'M': 20, 'N': 12, 'O': 12, 'P': 15, 'Q': 10}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    wb.save(output)
    return output.getvalue()
import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(
    page_title="Báo Cáo NXT NVL Chính",
    page_icon="📊",
    layout="wide"
)

st.title("📊 BÁO CÁO TỔNG HỢP NXT NVL CHÍNH")

# --- 1. HÀM XỬ LÝ INV 110 (TỒN KHO) ---
@st.cache_data
def process_inv110(files):
    df_list = []
    for f in files:
        try:
            # Sheet INV 110: Header bắt đầu từ dòng 11 (skiprows=10)
            df = pd.read_excel(f, skiprows=10)
            # Cột A (index 0): Phân loại, Cột B (index 1): Kho, Cột X (index 23): Khối lượng
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
    # Loại bỏ dòng tiêu đề rác nếu có
    data = data[~data['PhanLoai'].isin(['nan', 'Phân loại', 'None'])]

    # BẢNG 1: TỒN KHO THÉP CUỘN (HRC, PO, CM, CR)
    # HRC kho thường (khác 035)
    hrc_thuong = data[(data['PhanLoai'] == 'HRC') & (~data['Kho'].str.contains('035', na=False))]
    # HRC ngoài cảng (kho 035)
    hrc_cang = data[(data['PhanLoai'] == 'HRC') & (data['Kho'].str.contains('035', na=False))]
    # PO
    po = data[data['PhanLoai'] == 'PO']
    # CM + CR
    cm_cr = data[data['PhanLoai'].isin(['CM', 'CR'])]

    tb1 = pd.DataFrame({
        'Tên hàng': ['Tồn kho HRC', 'Tồn kho HRC ( hàng ngoài cảng)', 'Tồn kho PO', 'Tồn kho CM + CR'],
        'Số lượng (Cuộn)': [len(hrc_thuong), len(hrc_cang), len(po), len(cm_cr)],
        'Khối lượng (Tấn)': [
            hrc_thuong['KhoiLuong'].sum() / 1000,
            hrc_cang['KhoiLuong'].sum() / 1000,
            po['KhoiLuong'].sum() / 1000,
            cm_cr['KhoiLuong'].sum() / 1000
        ],
        'Ghi chú': ['', '', '', '']
    })
    # Dòng Tổng
    tb1.loc[len(tb1)] = ['Tổng', tb1['Số lượng (Cuộn)'].sum(), tb1['Khối lượng (Tấn)'].sum(), '']

    # BẢNG 2: KIM LOẠI MÀU
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


# --- 2. HÀM XỬ LÝ 105 (GIAO DỊCH SỬ DỤNG) ---
@st.cache_data
def process_105(files):
    df_list = []
    for f in files:
        try:
            # Header ở dòng 9 (skiprows=8)
            df = pd.read_excel(f, skiprows=8)
            # Cột C (index 2): Loại GD, Cột G (index 6): Tên mặt hàng, Cột L (index 11): Khối lượng
            if len(df.columns) >= 12:
                # Kiểm tra xem file đã có 2 cột phân loại A, B ở đầu hay chưa
                col_a = df.iloc[:, 0].astype(str).str.strip()
                col_b = df.iloc[:, 1].astype(str).str.strip()
                
                sub_df = pd.DataFrame({
                    'ColA': col_a,
                    'ColB': col_b,
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

    # Tự động gán phân loại nếu file raw chưa điền Cột A & B
    has_custom_cols = data['ColA'].isin(['HRC', 'PO', 'CM', 'A']).any()

    if has_custom_cols:
        # Trường hợp file đã có sẵn công thức cột A, B như sheet mẫu
        def get_nvl_stats(item_name):
            issue_cnt = len(data[(data['ColA'] == item_name) & (data['ColB'] == '1') & (data['LoaiGD'] == 'WIP Issue')])
            return_cnt = len(data[(data['ColA'] == item_name) & (data['ColB'] == '1') & (data['LoaiGD'] == 'WIP Return')])
            qty = issue_cnt - return_cnt
            weight = abs(data[(data['ColA'] == item_name) & (data['ColB'] == '1')]['KhoiLuong'].sum()) / 1000
            return qty, weight

        hrc_q, hrc_w = get_nvl_stats('HRC')
        po_q, po_w = get_nvl_stats('PO')
        cm_q, cm_w = get_nvl_stats('CM')
        
        # Hao hụt phế liệu: Cột B = '2' và Cột A = 'A'
        scrap_w = data[(data['ColB'] == '2') & (data['ColA'] == 'A')]['KhoiLuong'].sum() / 1000

        # NPL: Cột B = '3' trừ Cột B = '4'
        def get_npl_weight(npl_name):
            w3 = data[(data['ColA'].str.contains(npl_name, na=False, case=False)) & (data['ColB'] == '3')]['KhoiLuong'].sum()
            w4 = data[(data['ColA'].str.contains(npl_name, na=False, case=False)) & (data['ColB'] == '4')]['KhoiLuong'].sum()
            return abs(w3 - w4) / 1000

        nhom_kem_w = get_npl_weight('Hợp kim nhôm kẽm')
        kem_w = get_npl_weight('Kẽm không hợp kim')
        magie_w = get_npl_weight('Magie')
        pha_che_w = get_npl_weight('pha chế')

    else:
        # Trường hợp nạp file raw 105 thuần từ Oracle
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

    # BẢNG 3: SỬ DỤNG NVL CHÍNH
    tb3 = pd.DataFrame({
        'Tên hàng': ['HRC', 'PO', 'CM', 'Hao hụt+ phế liệu'],
        'Số lượng (Cuộn)': [hrc_q, po_q, cm_q, np.nan],
        'Khối lượng (Tấn)': [hrc_w, po_w, cm_w, scrap_w],
        'Ghi chú': ['', '', '', '']
    })
    tb3.loc[len(tb3)] = ['Tổng', tb3['Số lượng (Cuộn)'].sum(), tb3['Khối lượng (Tấn)'].sum(), '']

    # BẢNG 4: SỬ DỤNG NPL
    tb4 = pd.DataFrame({
        'Tên hàng': ['Hợp kim nhôm kẽm', 'Kẽm không hợp kim', 'Hợp kim kẽm Magie', 'Nguyên liệu pha chế'],
        'Khối lượng (Tấn)': [nhom_kem_w, kem_w, magie_w, pha_che_w],
        'Ghi chú': ['', '', '', '']
    })

    return tb3, tb4


# --- 3. GIAO DIỆN HƯỚNG DẪN & NẠP DỮ LIỆU ---
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


# --- 4. HIỂN THỊ KẾT QUẢ ĐỐI SOÁT ---
if inv110_files and file105_files:
    st.divider()

    def style_table(df, is_weight_only=False):
        if is_weight_only:
            return df.style.format({"Khối lượng (Tấn)": "{:,.2f}"}).apply(
                lambda r: ['background-color: #ffff00; font-weight: bold' if r['Tên hàng'] == 'Tổng' else '' for _ in r], axis=1
            )
        else:
            return df.style.format({
                "Số lượng (Cuộn)": lambda x: f"{x:,.0f}" if pd.notnull(x) else "",
                "Khối lượng (Tấn)": "{:,.2f}"
            }).apply(
                lambda r: ['background-color: #ffff00; font-weight: bold' if r['Tên hàng'] == 'Tổng' else '' for _ in r], axis=1
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
    # Tạo nút xuất Excel
    st.write("")
    excel_data = export_to_excel(df_tb1, df_tb2, df_tb3, df_tb4)
    st.download_button(
        label="📥 Tải Báo Cáo Ra File Excel (.xlsx)",
        data=excel_data,
        file_name="BAO_CAO_NXT_TONG_HOP.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
