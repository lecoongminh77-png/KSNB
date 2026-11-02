import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN CHUẨN APPLE & THANH TIỆN ÍCH CỐ ĐỊNH
# -------------------------------------------------------------
st.set_page_config(
    page_title='Đối Chiếu Nghiệm Thu HRC - KSNB',
    page_icon='🛡️',
    layout='wide',
)

# CSS ẩn menu sidebar và tùy chỉnh giao diện Apple
st.markdown(
    """
<style>
    /* Ẩn sidebar */
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }

    /* Nền trang */
    .stApp {
        background-color: #f5f5f7;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif;
        color: #1d1d1f;
    }
    
    /* Header */
    .apple-header-box {
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
        padding: 20px 0 10px 0;
        margin-bottom: 5px;
    }
    .apple-main-title {
        font-size: 40px;
        font-weight: 700;
        color: #1d1d1f;
        letter-spacing: -0.025em;
        margin: 0;
    }
    
    /* Nút bấm Apple Style */
    div.stButton > button:first-child {
        background: #0071e3 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding: 10px 20px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2) !important;
    }
    div.stButton > button:first-child:hover {
        background: #0077ed !important;
        color: #ffffff !important;
    }

    /* Container Streamlit Thẻ Apple */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff;
        border-radius: 16px;
        padding: 15px;
        border: 1px solid rgba(0, 0, 0, 0.08);
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.03);
    }

    /* Thành phần chữ trong thẻ */
    .upload-title {
        font-size: 18px;
        font-weight: 700;
        color: #1d1d1f;
        margin-bottom: 16px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .guide-box {
        background: #fbfbfd;
        border-radius: 12px;
        padding: 16px;
        margin-bottom: 16px;
        border: 1px solid rgba(0,0,0,0.04);
    }
    .guide-header {
        font-weight: 700;
        color: #0071e3;
        font-size: 14px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .guide-text {
        font-size: 13px;
        color: #515154;
        line-height: 1.6;
        margin: 0;
    }
</style>
""",
    unsafe_allow_html=True,
)

# Thanh điều hướng (Nút Trang chủ)
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', use_container_width=True):
        st.switch_page('app_ksnb.py')

st.markdown(
    "<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>",
    unsafe_allow_html=True,
)

# Header Trang
st.markdown(
    """
<div class="apple-header-box">
    <div>
        <h1 class="apple-main-title">📑 Đối Chiếu Nghiệm Thu HRC</h1>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    "<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>",
    unsafe_allow_html=True,
)


# -------------------------------------------------------------
# CÁC HÀM XỬ LÝ DỮ LIỆU TỰ ĐỘNG
# -------------------------------------------------------------
def read_auto_header(file_obj, keywords, sheet_name=0, max_scan=25, match_all=False):
    """Tự động tìm dòng header thực tế dựa trên danh sách từ khóa."""
    file_obj.seek(0)
    df_raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None, nrows=max_scan)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
        if match_all:
            if all(k.upper() in row_str for k in keywords):
                header_idx = idx
                break
        else:
            if any(k.upper() in row_str for k in keywords):
                header_idx = idx
                break
    file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name=sheet_name, skiprows=header_idx)
    return df


def to_float(val, default=0.0):
    """Chuyển đổi giá trị ô Excel sang số thực an toàn."""
    if val is None or val == "":
        return default
    try:
        return float(str(val).replace(',', '').strip())
    except Exception:
        return default


def clean_str_code(val):
    """Làm sạch mã cuộn / số phiếu cân (bỏ đuôi .0 và khoảng trắng)."""
    if val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r'\.0$', '', s)
    return s.strip()


def normalize_sheet_key(name):
    """Chuẩn hóa tên sheet để bắt cặp tự động chính xác giữa BBNT và Packing List."""
    s = str(name).upper()
    s = re.sub(r'\b(BBNT|PACKING LIST|PACKINGLIST|PL|BC|BAO CAO)\b', '', s)
    s = re.sub(r'[^A-Z0-9]', '', s)
    return s


def match_pl_sheet(bbnt_sheet_name, pl_sheet_names):
    """Tìm sheet Packing List khớp với sheet BBNT."""
    if not pl_sheet_names:
        return None
    # 1. Khớp chính xác 100%
    exact = next((p for p in pl_sheet_names if p.strip().upper() == bbnt_sheet_name.strip().upper()), None)
    if exact:
        return exact
    # 2. Khớp sau khi đã lược bỏ từ khóa thừa (BBNT, packing list, PL...)
    target_key = normalize_sheet_key(bbnt_sheet_name)
    if target_key:
        norm_match = next((p for p in pl_sheet_names if normalize_sheet_key(p) == target_key), None)
        if norm_match:
            return norm_match
        partial_match = next(
            (p for p in pl_sheet_names if target_key in normalize_sheet_key(p) or normalize_sheet_key(p) in target_key),
            None,
        )
        if partial_match:
            return partial_match
    # 3. Nếu file Packing List chỉ có đúng 1 sheet thì lấy luôn sheet đó
    if len(pl_sheet_names) == 1:
        return pl_sheet_names[0]
    return None


def ghi_va_kiem_tra_worksheet(
    ws,
    s_name,
    f_pl,
    sheet_pl,
    bv_raw_list,
    dict_tc_kl,
    dict_tc_phanloai,
    opm_035_po,
    opm_035_rtv,
    opm_031,
    opm_03g_po,
    opm_03g_rtv,
):
    """
    - Giữ nguyên 100% số liệu & bố cục BBNT gốc (Cột A -> M).
    - Kiểm tra nội bộ BBNT gốc (định dạng Text làm sai hàm SUM, sai phép trừ K=J-I, sai dòng Tổng cộng) và tô đỏ cảnh báo.
    - Ghi 19 cột đối chiếu rõ nghĩa từ Cột N (14) đến Cột AF (32), tự động Unhide mọi cột.
    - Trả về dữ liệu báo cáo nhanh để hiển thị trực tiếp trên Web.
    """
    # 1. Xác định dòng tiêu đề bảng (mặc định là dòng 8)
    header_row = 8
    for r in range(1, 20):
        row_vals = [str(ws.cell(row=r, column=c).value or '').upper() for c in range(1, 14)]
        if any('MÃ SỐ CUỘN' in v or 'MÃ CUỘN' in v for v in row_vals):
            header_row = r
            break

    header_end_row = header_row
    for m_range in ws.merged_cells.ranges:
        if m_range.min_row == header_row and m_range.min_col <= 13 and m_range.max_row > header_row:
            header_end_row = m_range.max_row
            break

    # 2. Tìm các dòng dữ liệu cuộn thép và dòng Tổng cộng
    data_rows = []
    total_row = None
    for r in range(header_end_row + 1, ws.max_row + 1):
        val_a = str(ws.cell(row=r, column=1).value or '').strip()
        val_d = clean_str_code(ws.cell(row=r, column=4).value).upper()

        if 'TỔNG CỘNG' in val_a.upper():
            total_row = r
            break
        if any(k in val_a.upper() for k in ['KẾT LUẬN', 'GHI CHÚ', 'BAN GIÁM ĐỐC']):
            break

        if val_d != '' and val_d != 'NONE':
            data_rows.append(r)

    # 3. Quét mã cuộn trong đúng Cột MÃ HÀNG HÓA của file Bảo vệ
    cuon_list = [clean_str_code(ws.cell(row=r, column=4).value).upper() for r in data_rows]
    dict_bv_p_sheet = {}
    dict_bv_kl_sheet = {}
    if bv_raw_list:
        for txt_ma_hang, p, kl in bv_raw_list:
            for m in cuon_list:
                if m and m in txt_ma_hang and m not in dict_bv_p_sheet:
                    dict_bv_p_sheet[m] = p
                    dict_bv_kl_sheet[m] = kl

    # 4. Đọc Packing List tương ứng
    dict_pl = {}
    if f_pl is not None and sheet_pl is not None:
        try:
            df_pl = read_auto_header(f_pl, ['MÃ SỐ NSX', 'MÃ SỐ HSPM', 'KL(NET)'], sheet_name=sheet_pl)
            c_pl_ma = next((c for c in df_pl.columns if 'MÃ SỐ NSX' in str(c).upper() or 'MÃ SỐ' in str(c).upper()), None)
            c_pl_kl = next(
                (c for c in df_pl.columns if 'NET' in str(c).upper() and 'CHÊNH LỆCH' not in str(c).upper()),
                None,
            )

            if c_pl_ma and c_pl_kl:
                df_pl_clean = df_pl.dropna(subset=[c_pl_ma, c_pl_kl]).copy()
                pl_key = df_pl_clean[c_pl_ma].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper() + 'A'
                pl_val = pd.to_numeric(df_pl_clean[c_pl_kl], errors='coerce').fillna(0)
                dict_pl = pd.DataFrame({'Key': pl_key, 'Net': pl_val}).groupby('Key')['Net'].sum().to_dict()
        except Exception:
            dict_pl = {}

    # 5. Danh sách 19 cột đối chiếu rõ nghĩa (Đã bỏ cột Mã hóa, bắt đầu từ Cột N = 14 đến Cột AF = 32)
    extra_headers = [
        'Loại phiếu trạm cân\n(Nhập/Xuất)',                  # 14 (N)
        'Khối lượng Net\ntheo Packing List',                 # 15 (O)
        'Chênh lệch KL Net\n(BBNT - Packing List)',          # 16 (P)
        'Khối lượng theo\nBáo cáo Bảo vệ',                   # 17 (Q)
        'Số phiếu cân theo\nBáo cáo Bảo vệ',                 # 18 (R)
        'Khớp Số phiếu cân\n(BBNT & Bảo vệ)',                # 19 (S)
        'Khối lượng theo\nDữ liệu Trạm cân',                 # 20 (T)
        'Khớp Khối lượng\n(Trạm cân & Bảo vệ)',              # 21 (U)
        'Khớp Khối lượng\n(Trạm cân & BBNT)',                # 22 (V)
        'Chênh lệch KL (kg)\n(Trạm cân - Gross BBNT)',       # 23 (W)
        'Tỷ lệ chênh lệch\ndung sai (%)',                    # 24 (X)
        'Khớp Tỷ lệ dung sai\nvới BBNT',                     # 25 (Y)
        'Đánh giá dung sai\n(Quy định ±0.3%)',               # 26 (Z)
        'KL Nhập kho 035\n(ERP OPM105)',                     # 27 (AA)
        'Chênh lệch KL Kho 035\n(PL - Kho 035)',             # 28 (AB)
        'KL Nhập kho 031\n(Intransit)',                      # 29 (AC)
        'Chênh lệch KL Kho 031\n(PL - Kho 031)',             # 30 (AD)
        'KL Nhập kho 03G-3H\n(ERP OPM105)',                  # 31 (AE)
        'Chênh lệch KL Kho 03G-3H\n(PL - Kho 03G-3H)',       # 32 (AF)
    ]

    font_header = Font(name='Times New Roman', size=11, bold=True, color='000000')
    font_data = Font(name='Times New Roman', size=11, bold=False, color='000000')
    font_total = Font(name='Times New Roman', size=11, bold=True, color='000000')
    font_error = Font(name='Times New Roman', size=11, bold=True, color='9C0006')

    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

    border_thin = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000'),
    )

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center')

    start_col = 14  # Cột N

    # Xóa các vùng merge cũ từ cột N trở đi (nếu có) và dọn cột thừa
    merges_to_remove = [m for m in ws.merged_cells.ranges if m.min_col >= start_col]
    for m in merges_to_remove:
        ws.unmerge_cells(str(m))

    # 6. Ghi tiêu đề 19 cột Đối chiếu vào dòng 8-9
    for idx, h_title in enumerate(extra_headers):
        col_idx = start_col + idx
        cell_top = ws.cell(row=header_row, column=col_idx, value=h_title)
        cell_top.font = font_header
        cell_top.fill = yellow_fill
        cell_top.alignment = align_center
        cell_top.border = border_thin

        if header_end_row > header_row:
            for r_sub in range(header_row + 1, header_end_row + 1):
                c_sub = ws.cell(row=r_sub, column=col_idx)
                c_sub.fill = yellow_fill
                c_sub.border = border_thin
            ws.merge_cells(
                start_row=header_row,
                start_column=col_idx,
                end_row=header_end_row,
                end_column=col_idx,
            )

    # Bộ đếm tổng cho các cột khối lượng bên phải (Cột N -> AF)
    sum_cols = {15: 0.0, 16: 0.0, 17: 0.0, 20: 0.0, 23: 0.0, 27: 0.0, 28: 0.0, 29: 0.0, 30: 0.0, 31: 0.0, 32: 0.0}
    # Bộ đếm kiểm tra tổng thực tế của BBNT gốc (Cột H=8, I=9, J=10, K=11)
    actual_bbnt_sums = {8: 0.0, 9: 0.0, 10: 0.0, 11: 0.0}

    bbnt_internal_errors = []
    coil_error_records = []

    # 7. Duyệt từng dòng cuộn thép để kiểm tra BBNT gốc & tính cột đối chiếu
    for r in data_rows:
        stt = ws.cell(row=r, column=1).value
        ma_cuon = clean_str_code(ws.cell(row=r, column=4).value).upper()
        phieu_can_bbnt = clean_str_code(ws.cell(row=r, column=5).value)

        # Kiểm tra định dạng Text trên các cột khối lượng của BBNT gốc (H, I, J, K)
        for c_b, c_label in [
            (8, 'Cột H (KL Net chứng từ)'),
            (9, 'Cột I (KL Gross chứng từ)'),
            (10, 'Cột J (KL cân bàn thực tế)'),
            (11, 'Cột K (Chênh lệch kg)'),
        ]:
            raw_v = ws.cell(row=r, column=c_b).value
            if isinstance(raw_v, str) and raw_v.strip() != '':
                col_let = get_column_letter(c_b)
                # Giữ nguyên giá trị gốc nhưng tô đỏ ô trên BBNT để báo lỗi
                ws.cell(row=r, column=c_b).fill = red_fill
                ws.cell(row=r, column=c_b).font = font_error
                bbnt_internal_errors.append({
                    'Sheet': s_name,
                    'Vị trí ô': f'{col_let}{r} (STT {stt} - Cuộn {ma_cuon})',
                    'Loại lỗi trên BBNT gốc': f'Nhập định dạng Chữ (Text) tại {c_label}',
                    'Giá trị trên BBNT': repr(raw_v),
                    'Giá trị đúng': f'{to_float(raw_v):,.0f} (Kiểu Số)',
                    'Đơn vị cần làm việc': 'Đơn vị lập BBNT (Sửa lại định dạng Số để hàm SUM không bỏ sót)',
                })

        net_ct = to_float(ws.cell(row=r, column=8).value, 0.0)
        gross_ct = to_float(ws.cell(row=r, column=9).value, 0.0)
        can_ban = to_float(ws.cell(row=r, column=10).value, 0.0)
        cl_kg_bbnt = to_float(ws.cell(row=r, column=11).value, 0.0)
        dsai_bbnt = to_float(ws.cell(row=r, column=12).value, 0.0)

        actual_bbnt_sums[8] += net_ct
        actual_bbnt_sums[9] += gross_ct
        actual_bbnt_sums[10] += can_ban
        actual_bbnt_sums[11] += cl_kg_bbnt

        # Kiểm tra phép trừ nội bộ trên BBNT gốc: Cột K = Cột J - Cột I
        expected_cl_kg = can_ban - gross_ct
        if round(cl_kg_bbnt, 2) != round(expected_cl_kg, 2):
            ws.cell(row=r, column=11).fill = red_fill
            ws.cell(row=r, column=11).font = font_error
            bbnt_internal_errors.append({
                'Sheet': s_name,
                'Vị trí ô': f'K{r} (STT {stt} - Cuộn {ma_cuon})',
                'Loại lỗi trên BBNT gốc': 'Tính sai Chênh lệch (kg) = Cân bàn (J) - Gross (I)',
                'Giá trị trên BBNT': f'{cl_kg_bbnt:,.0f}',
                'Giá trị đúng': f'{expected_cl_kg:,.0f}',
                'Đơn vị cần làm việc': 'Đơn vị lập BBNT',
            })

        # Tính toán các cột đối chiếu (Cột N -> AF)
        ma_hoa = ma_cuon + 'A'
        nhap_xuat = dict_tc_phanloai.get(phieu_can_bbnt, 'Nhập hàng')
        kl_net_pl = float(dict_pl.get(ma_hoa, 0.0))
        doi_chieu_pl = net_ct - kl_net_pl

        kl_bv = float(to_float(dict_bv_kl_sheet.get(ma_cuon, 0.0), 0.0))
        phieu_bv = str(dict_bv_p_sheet.get(ma_cuon, ''))
        dc_phieu = (phieu_can_bbnt == phieu_bv)

        kl_tc = float(to_float(dict_tc_kl.get(phieu_bv, 0.0), 0.0))
        dc_tc_bv = (kl_tc == kl_bv)
        dc_tc_bbnt = (kl_tc == can_ban)

        chenh_lech_kg = kl_tc - gross_ct
        chenh_lech_dsai = (chenh_lech_kg / gross_ct) if gross_ct != 0 else 0.0
        dc_dsai = (round(chenh_lech_dsai, 6) == round(dsai_bbnt, 6))
        so_quy_dinh = 'VUOT DUNG SAI' if abs(chenh_lech_dsai) > 0.003 else 'OK'

        kl_035 = float(opm_035_po.get(ma_hoa, 0.0) - abs(opm_035_rtv.get(ma_hoa, 0.0)))
        cl_035 = kl_net_pl - kl_035

        kl_031 = float(opm_031.get(ma_hoa, 0.0))
        cl_031 = 0.0 if kl_031 == 0 else (kl_net_pl - kl_031)

        kl_03g = float(opm_03g_po.get(ma_hoa, 0.0) - abs(opm_03g_rtv.get(ma_hoa, 0.0)))
        cl_03g = kl_net_pl - kl_03g

        row_values = [
            nhap_xuat,        # 14 (N)
            kl_net_pl,        # 15 (O)
            doi_chieu_pl,     # 16 (P)
            kl_bv,            # 17 (Q)
            phieu_bv,         # 18 (R)
            dc_phieu,         # 19 (S)
            kl_tc,            # 20 (T)
            dc_tc_bv,         # 21 (U)
            dc_tc_bbnt,       # 22 (V)
            chenh_lech_kg,    # 23 (W)
            chenh_lech_dsai,  # 24 (X)
            dc_dsai,          # 25 (Y)
            so_quy_dinh,      # 26 (Z)
            kl_035,           # 27 (AA)
            cl_035,           # 28 (AB)
            kl_031,           # 29 (AC)
            cl_031,           # 30 (AD)
            kl_03g,           # 31 (AE)
            cl_03g,           # 32 (AF)
        ]

        for idx, val in enumerate(row_values):
            col_idx = start_col + idx
            cell = ws.cell(row=r, column=col_idx, value=val)
            cell.font = font_data
            cell.border = border_thin

            if col_idx in sum_cols:
                cell.alignment = align_right
                cell.number_format = '#,##0'
                sum_cols[col_idx] += float(val)
            elif col_idx == 24:  # Tỷ lệ chênh lệch dung sai (%)
                cell.alignment = align_right
                cell.number_format = '0.00%'
            else:
                cell.alignment = align_center

            # Kiểm tra điều kiện tô đỏ ô sai phạm
            is_error = False
            if col_idx in [19, 21, 22, 25] and val is False:
                is_error = True
            elif col_idx == 26 and str(val).strip() == 'VUOT DUNG SAI':
                is_error = True
            elif col_idx in [16, 28] and isinstance(val, (int, float)) and val != 0:
                is_error = True

            if is_error:
                cell.fill = red_fill
                cell.font = font_error

        # Tổng hợp lỗi của từng cuộn để báo cáo trên Web ("Sai ở đâu & Đơn vị liên quan")
        err_reasons = []
        err_depts = []

        if doi_chieu_pl != 0:
            err_depts.append('Packing List / BBNT')
            err_reasons.append(f'Lệch KL Net BBNT ({net_ct:,.0f}) và Packing List ({kl_net_pl:,.0f}): {doi_chieu_pl:+,.0f} kg')

        if not dc_phieu:
            err_depts.append('Bảo vệ')
            if phieu_bv == '':
                err_reasons.append(f'Bảo vệ chưa nhập Mã cuộn vào cột MÃ HÀNG HÓA (Phiếu BBNT: {phieu_can_bbnt})')
            else:
                err_reasons.append(f'Lệch Số phiếu cân: BBNT ({phieu_can_bbnt}) ≠ Bảo vệ ({phieu_bv})')

        if not dc_tc_bv:
            if 'Bảo vệ' not in err_depts:
                err_depts.append('Bảo vệ / Trạm cân')
            err_reasons.append(f'Lệch KL Trạm cân ({kl_tc:,.0f}) ≠ KL Bảo vệ ({kl_bv:,.0f})')

        if not dc_tc_bbnt:
            if 'Trạm cân / BBNT' not in err_depts:
                err_depts.append('Trạm cân / BBNT')
            err_reasons.append(f'Lệch KL Trạm cân ({kl_tc:,.0f}) ≠ Cân bàn BBNT ({can_ban:,.0f})')

        if not dc_dsai:
            if 'Đơn vị lập BBNT' not in err_depts:
                err_depts.append('Đơn vị lập BBNT')
            err_reasons.append(f'Lệch % dung sai: Tính lại ({chenh_lech_dsai*100:.2f}%) ≠ BBNT ({dsai_bbnt*100:.2f}%)')

        if so_quy_dinh == 'VUOT DUNG SAI':
            err_depts.append('Vượt dung sai (±0.3%)')
            err_reasons.append(f'Vượt dung sai quy định: {chenh_lech_dsai*100:.2f}% ({chenh_lech_kg:+,.0f} kg)')

        if cl_035 != 0:
            err_depts.append('Kho ERP (OPM105)')
            err_reasons.append(f'Lệch KL Nhập kho 035: PL ({kl_net_pl:,.0f}) vs Kho 035 ({kl_035:,.0f})')

        if err_reasons:
            coil_error_records.append({
                'Sheet': s_name,
                'STT': stt,
                'Dòng Excel': r,
                'Mã số cuộn': ma_cuon,
                'Số phiếu cân BBNT': phieu_can_bbnt,
                'Đơn vị / Khâu cần kiểm tra': ', '.join(dict.fromkeys(err_depts)),
                'Chi tiết lỗi phát hiện': ' | '.join(err_reasons),
            })

    # 8. Kiểm tra dòng Tổng cộng trên BBNT gốc (Giữ nguyên số gốc, chỉ tô đỏ nếu sai) & ghi dòng Tổng cộng bên phải
    if total_row is not None:
        # Kiểm tra 4 cột tổng bên BBNT gốc (H=8, I=9, J=10, K=11)
        for c_b, c_label in [
            (8, 'Cột H (Tổng KL Net chứng từ)'),
            (9, 'Cột I (Tổng KL Gross chứng từ)'),
            (10, 'Cột J (Tổng KL cân bàn thực tế)'),
            (11, 'Cột K (Tổng Chênh lệch kg)'),
        ]:
            orig_total = to_float(ws.cell(row=total_row, column=c_b).value, 0.0)
            true_total = actual_bbnt_sums[c_b]
            if round(orig_total, 2) != round(true_total, 2):
                col_let = get_column_letter(c_b)
                ws.cell(row=total_row, column=c_b).fill = red_fill
                ws.cell(row=total_row, column=c_b).font = font_error
                diff_tot = true_total - orig_total
                bbnt_internal_errors.append({
                    'Sheet': s_name,
                    'Vị trí ô': f'{col_let}{total_row} (Dòng Tổng cộng)',
                    'Loại lỗi trên BBNT gốc': f'Cộng sai dòng Tổng cộng tại {c_label}',
                    'Giá trị trên BBNT': f'{orig_total:,.0f}',
                    'Giá trị đúng': f'{true_total:,.0f} (Lệch {diff_tot:+,.0f} kg)',
                    'Đơn vị cần làm việc': 'Đơn vị lập BBNT (Kiểm tra lại công thức SUM dòng Tổng cộng)',
                })

        # Ghi tổng cho các cột đối chiếu bên phải (Cột N -> AF)
        for idx in range(len(extra_headers)):
            col_idx = start_col + idx
            cell = ws.cell(row=total_row, column=col_idx)
            cell.border = border_thin
            cell.font = font_total
            if col_idx in sum_cols:
                cell.value = sum_cols[col_idx]
                cell.alignment = align_right
                cell.number_format = '#,##0'

    # 9. Bỏ ẩn (Unhide) toàn bộ các cột từ A đến AF và căn chỉnh độ rộng đẹp mắt
    for c_i in range(1, start_col + len(extra_headers) + 2):
        col_let = get_column_letter(c_i)
        ws.column_dimensions[col_let].hidden = False

    col_widths = {
        14: 16, 15: 18, 16: 18, 17: 18, 18: 18, 19: 18,
        20: 18, 21: 18, 22: 18, 23: 18, 24: 16, 25: 18,
        26: 18, 27: 16, 28: 18, 29: 16, 30: 18, 31: 16, 32: 18,
    }
    for col_idx, width in col_widths.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    return {
        'sheet_name': s_name,
        'matched_pl': sheet_pl or 'Không tìm thấy',
        'total_coils': len(data_rows),
        'error_coils_count': len(coil_error_records),
        'ok_coils_count': len(data_rows) - len(coil_error_records),
        'bbnt_internal_errors': bbnt_internal_errors,
        'coil_error_records': coil_error_records,
    }


# -------------------------------------------------------------
# CHỨC NĂNG CHÍNH
# -------------------------------------------------------------
def doi_chieu_hrc_tung_file():
    st.caption('Kiểm soát nội bộ • Giữ nguyên BBNT gốc, tự động đối soát 4 bên và chỉ rõ vị trí sai phạm')
    st.markdown("### 📥 Nạp Dữ Liệu Nguồn")

    col_upload1, col_upload2 = st.columns(2, gap="large")

    with col_upload1:
        with st.container(border=True):
            st.markdown("""
<div class="upload-title">1️⃣ Dữ liệu Nghiệm thu & Trạm cân</div>
<div class="guide-box">
<div class="guide-header">📌 1. Biên bản nghiệm thu (BBNT)</div>
<div class="guide-text">Đường dẫn: Nguyen Thi Thu Thuong > Biên bản nghiệm thu (BBNT)</div>
</div>
""", unsafe_allow_html=True)

            f_bbnt = st.file_uploader('1. File Biên bản nghiệm thu (BBNT)', type=['xlsx', 'xls'])

            selected_sheets = []
            if f_bbnt is not None:
                try:
                    f_bbnt.seek(0)
                    xl_bb = pd.ExcelFile(f_bbnt)
                    sheet_names_bb = xl_bb.sheet_names
                    selected_sheets = st.multiselect(
                        '🎯 Chọn Sheet cần đối chiếu:',
                        options=sheet_names_bb,
                        default=[sheet_names_bb[0]] if sheet_names_bb else [],
                        help='Các sheet không được chọn sẽ tự động bị loại bỏ khỏi file kết quả'
                    )
                except Exception as e:
                    st.error(f"Lỗi đọc danh sách sheet: {e}")
            else:
                st.caption("ℹ️ Nạp file BBNT để hiện danh sách sheet.")

            st.markdown("""
<div class="guide-box" style="margin-top: 15px;">
<div class="guide-header">📌 2. Packing List nhà cung cấp</div>
<div class="guide-text">Đường dẫn: Nguyen Thi Thu Thuong > Packing List nhà cung cấp</div>
</div>
""", unsafe_allow_html=True)
            f_pl = st.file_uploader('2. File Packing List nhà cung cấp', type=['xlsx', 'xls'])

            st.markdown("""
<div class="guide-box">
<div class="guide-header">📌 3. Dữ liệu Trạm cân</div>
<div class="guide-text">
Truy cập: http://172.17.24.9/#/login (PM4444 / 123456)<br>
Quản lý truy cập > Tra cứu trạm cân<br>
Thời gian: 1 tháng rưỡi
</div>
</div>
""", unsafe_allow_html=True)
            f_tc = st.file_uploader('3. File Dữ liệu Trạm cân', type=['xlsx', 'xls'])

    with col_upload2:
        with st.container(border=True):
            st.markdown("""
<div class="upload-title">2️⃣ Dữ liệu Bảo vệ & ERP</div>
<div class="guide-box">
<div class="guide-header">📌 4. Dữ liệu Nhập Bảo vệ</div>
<div class="guide-text">
Quản lý truy cập > Tra cứu xuất nhập hàng > Chọn Nhập hàng<br>
Thời gian: Giống mốc Tra cứu trạm cân
</div>
</div>
<div class="guide-box">
<div class="guide-header">📌 5. Báo cáo ERP OPM 105</div>
<div class="guide-text">
Chọn: HSG_OPM_105_BCKiemSoatThongTinGDTheoLo_GD2<br>
Tích chọn kho: 035, 03G, 03H<br>
Nghiệp vụ: Inventory, Job or Schedule, Purchase order<br>
Thời gian: 1 tháng cần kiểm tra
</div>
</div>
""", unsafe_allow_html=True)
            f_bv = st.file_uploader('4. File Dữ liệu Nhập Bảo vệ', type=['xlsx', 'xls'])
            f_opm = st.file_uploader('5. File Báo cáo ERP OPM 105', type=['xlsx', 'xls'])

    # TIẾN HÀNH XỬ LÝ
    st.write("")
    if f_bbnt is not None:
        if not selected_sheets:
            st.warning("⚠️ Vui lòng chọn ít nhất 1 sheet ở khung bên trái để bắt đầu đối chiếu.")
        else:
            if st.button('🚀 Bắt đầu đối chiếu tự động', use_container_width=True):
                with st.spinner(f'Đang đối soát {len(selected_sheets)} sheet dữ liệu và kiểm tra chéo...'):
                    try:
                        # 1. QUÉT DỮ LIỆU BẢO VỆ (CHỈ QUÉT ĐÚNG CỘT MÃ HÀNG HÓA)
                        bv_raw_list = []
                        if f_bv is not None:
                            try:
                                df_bv = read_auto_header(f_bv, ['MÃ PHIẾU CÂN', 'KHỐI LƯỢNG PHIẾU CÂN'])
                                col_bv_mahang = next((c for c in df_bv.columns if 'MÃ HÀNG HÓA' in str(c).upper()), None)
                                col_bv_pcan = next((c for c in df_bv.columns if 'MÃ PHIẾU CÂN' in str(c).upper() or 'PHIẾU CÂN' in str(c).upper()), None)
                                col_bv_kl = next((c for c in df_bv.columns if 'KHỐI LƯỢNG' in str(c).upper()), None)

                                if col_bv_mahang and col_bv_pcan and col_bv_kl:
                                    for _, r in df_bv.iterrows():
                                        val_mh = r.get(col_bv_mahang, '')
                                        txt_ma_hang = '' if pd.isna(val_mh) else str(val_mh).upper().strip()
                                        val_p = r.get(col_bv_pcan, '')
                                        p = '' if pd.isna(val_p) else clean_str_code(val_p)
                                        kl = to_float(r.get(col_bv_kl, 0), 0.0)
                                        bv_raw_list.append((txt_ma_hang, p, kl))
                            except Exception:
                                bv_raw_list = []

                        # 2. QUÉT DỮ LIỆU TRẠM CÂN
                        dict_tc_kl, dict_tc_phanloai = {}, {}
                        if f_tc is not None:
                            try:
                                df_tc = read_auto_header(f_tc, ['PHIẾU CÂN', 'KHỐI LƯỢNG PHIẾU CÂN'])
                                col_tc_p = next((c for c in df_tc.columns if str(c).strip().upper() == 'PHIẾU CÂN'), df_tc.columns[0])
                                col_tc_kl = next((c for c in df_tc.columns if 'KHỐI LƯỢNG' in str(c).upper()), None)
                                col_tc_pl = next((c for c in df_tc.columns if 'PHÂN LOẠI' in str(c).upper()), None)

                                tc_p = df_tc[col_tc_p].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                                tc_kl = pd.to_numeric(df_tc[col_tc_kl], errors='coerce').fillna(0) if col_tc_kl else 0
                                dict_tc_kl = dict(zip(tc_p, tc_kl))
                                if col_tc_pl:
                                    dict_tc_phanloai = dict(zip(tc_p, df_tc[col_tc_pl].fillna('Nhập hàng').astype(str)))
                            except Exception:
                                dict_tc_kl, dict_tc_phanloai = {}, {}

                        # 3. QUÉT DỮ LIỆU ERP OPM 105
                        opm_035_po, opm_035_rtv, opm_031, opm_03g_po, opm_03g_rtv = {}, {}, {}, {}, {}
                        if f_opm is not None:
                            try:
                                df_opm = read_auto_header(f_opm, ['LOẠI GIAO DỊCH', 'MÃ LÔ'], match_all=True)
                                if len(df_opm) > 0 and '[' in str(df_opm.iloc[0, 1]):
                                    df_opm = df_opm.iloc[1:].copy()

                                c_opm_loai = next((c for c in df_opm.columns if 'LOẠI GIAO DỊCH' in str(c).upper()), None)
                                c_opm_malo = next((c for c in df_opm.columns if 'MÃ LÔ' in str(c).upper()), None)
                                c_opm_kl = next((c for c in df_opm.columns if 'KHỐI LƯỢNG' in str(c).upper() and 'GROSS' not in str(c).upper()), None)
                                c_opm_kho = next((c for c in df_opm.columns if any(k in str(c).upper() for k in ['MĂ KHO', 'MÃ KHO', 'MA KHO', 'KHO']) and 'KHU VỰC' not in str(c).upper()), None)

                                if c_opm_loai and c_opm_malo and c_opm_kl and c_opm_kho:
                                    df_opm_clean = df_opm.dropna(subset=[c_opm_malo, c_opm_loai]).copy()
                                    lo_str = df_opm_clean[c_opm_malo].astype(str).str.strip().str.upper()
                                    df_opm_clean['Key'] = lo_str.apply(lambda x: x if x.endswith('A') else x + 'A')
                                    df_opm_clean['KL'] = pd.to_numeric(df_opm_clean[c_opm_kl], errors='coerce').fillna(0)
                                    df_opm_clean['Kho_Clean'] = df_opm_clean[c_opm_kho].astype(str).str.strip().str.upper()

                                    opm_035_po = df_opm_clean[(df_opm_clean['Kho_Clean'] == '035') & (df_opm_clean[c_opm_loai].str.contains('PO Receipt', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
                                    opm_035_rtv = df_opm_clean[(df_opm_clean['Kho_Clean'] == '035') & (df_opm_clean[c_opm_loai].str.contains('Return to Vendor', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
                                    opm_031 = df_opm_clean[(df_opm_clean['Kho_Clean'].isin(['03G', '03H'])) & (df_opm_clean[c_opm_loai].str.contains('Intransit Receipt', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
                                    opm_03g_po = df_opm_clean[(df_opm_clean['Kho_Clean'] == '03G') & (df_opm_clean[c_opm_loai].str.contains('PO Receipt', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
                                    opm_03g_rtv = df_opm_clean[(df_opm_clean['Kho_Clean'] == '03G') & (df_opm_clean[c_opm_loai].str.contains('Return to Vendor', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
                            except Exception:
                                opm_035_po, opm_035_rtv, opm_031, opm_03g_po, opm_03g_rtv = {}, {}, {}, {}, {}

                        # 4. DANH SÁCH SHEET CỦA PACKING LIST ĐỂ BẮT CẶP THÔNG MINH
                        pl_sheet_names = []
                        if f_pl is not None:
                            try:
                                f_pl.seek(0)
                                pl_sheet_names = pd.ExcelFile(f_pl).sheet_names
                            except Exception:
                                pl_sheet_names = []

                        # 5. MỞ TRỰC TIẾP FILE BBNT GỐC BẰNG OPENPYXL (GIỮ NGUYÊN SỐ LIỆU GỐC & BỐ CỤC)
                        f_bbnt.seek(0)
                        wb = openpyxl.load_workbook(f_bbnt, data_only=True)

                        for s in list(wb.sheetnames):
                            if s not in selected_sheets:
                                del wb[s]

                        all_reports = []
                        for s_name in selected_sheets:
                            if s_name in wb.sheetnames:
                                ws = wb[s_name]
                                matched_pl = match_pl_sheet(s_name, pl_sheet_names)

                                report = ghi_va_kiem_tra_worksheet(
                                    ws,
                                    s_name,
                                    f_pl,
                                    matched_pl,
                                    bv_raw_list,
                                    dict_tc_kl,
                                    dict_tc_phanloai,
                                    opm_035_po,
                                    opm_035_rtv,
                                    opm_031,
                                    opm_03g_po,
                                    opm_03g_rtv,
                                )
                                all_reports.append(report)

                        output = io.BytesIO()
                        wb.save(output)

                        st.success(f'✅ Đã đối soát hoàn tất {len(selected_sheets)} sheet! File kết quả giữ nguyên 100% số liệu BBNT gốc và tô đỏ các ô sai lệch.')
                        st.download_button(
                            f'📥 Tải File Kết Quả Nghiệm Thu ({len(selected_sheets)} Sheet)',
                            data=output.getvalue(),
                            file_name='BBNT_Nghiem_Thu_Da_Doi_Chieu.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            use_container_width=True,
                        )

                        # -------------------------------------------------------------
                        # 6. HIỂN THỊ BẢNG BÁO CÁO NHANH KẾT QUẢ TRÊN MÀN HÌNH WEB
                        # -------------------------------------------------------------
                        st.markdown("### 📊 Báo Cáo Nhanh Kết Quả Kiểm Soát (Theo Từng Sheet)")
                        for rep in all_reports:
                            with st.container(border=True):
                                st.markdown(
                                    f"#### 📄 Sheet BBNT: `{rep['sheet_name']}` 🔗 Ghép với Sheet Packing List: `{rep['matched_pl']}`"
                                )
                                c1, c2, c3, c4 = st.columns(4)
                                c1.metric("📦 Tổng số cuộn", f"{rep['total_coils']} cuộn")
                                c2.metric("✅ Khớp đúng 100%", f"{rep['ok_coils_count']} cuộn")
                                c3.metric("❌ Cuộn có sai lệch", f"{rep['error_coils_count']} cuộn")
                                c4.metric("⚠️ Lỗi biểu mẫu BBNT gốc", f"{len(rep['bbnt_internal_errors'])} lỗi")

                                # Bảng 1: Lỗi phát hiện ngay trên BBNT gốc (Cột A -> M)
                                if rep['bbnt_internal_errors']:
                                    st.warning(
                                        f"🚨 **Phát hiện {len(rep['bbnt_internal_errors'])} lỗi ngay trên phần BBNT gốc (Cột A → M):** "
                                        f"(Code đã giữ nguyên số gốc trên file và tô đỏ ô lỗi để bạn làm việc với đơn vị lập BBNT)"
                                    )
                                    st.dataframe(
                                        pd.DataFrame(rep['bbnt_internal_errors']),
                                        use_container_width=True,
                                        hide_index=True,
                                    )
                                else:
                                    st.info("✔️ Phần biểu mẫu và dòng Tổng cộng trên BBNT gốc (Cột A → M) không có lỗi cộng sai hay định dạng chữ.")

                                # Bảng 2: Danh sách cuộn sai lệch khi đối chiếu 4 bên
                                if rep['coil_error_records']:
                                    st.error(
                                        f"🔍 **Chi tiết {rep['error_coils_count']} cuộn phát hiện sai lệch khi đối chiếu (Cột N → AF):**"
                                    )
                                    st.dataframe(
                                        pd.DataFrame(rep['coil_error_records']),
                                        use_container_width=True,
                                        hide_index=True,
                                    )
                                else:
                                    st.success("🎉 Toàn bộ các cuộn trong sheet này đều khớp 100% giữa BBNT, Packing List, Bảo vệ, Trạm cân và Kho ERP!")

                    except Exception as e:
                        st.error(f'Lỗi xử lý: {e}')
    else:
        st.info('👉 Hãy nạp tối thiểu File BBNT để bắt đầu đối chiếu.')


doi_chieu_hrc_tung_file()



