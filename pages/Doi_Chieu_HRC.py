import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN CHUẨN DASHBOARD KIỂM SOÁT NỘI BỘ
# -------------------------------------------------------------
st.set_page_config(
    page_title='Kiểm soát nội bộ - Đối Chiếu Nghiệm Thu HRC',
    page_icon='🛡️',
    layout='wide',
)

st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    /* Ẩn sidebar mặc định */
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }

    /* Nền trang & Font chữ */
    .stApp {
        background-color: #f8fafc;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        color: #1e293b;
    }

    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 3rem !important;
        max-width: 98% !important;
    }

    /* Header giống mẫu KSNB */
    .ksnb-dept-sub {
        font-size: 11.5px;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-bottom: 2px;
    }
    .ksnb-main-title {
        font-size: 26px;
        font-weight: 700;
        color: #0f172a;
        margin: 0 0 12px 0;
        letter-spacing: -0.015em;
    }

    /* Thanh Pill thông tin phiên kiểm tra */
    .session-pill {
        display: inline-flex;
        align-items: center;
        background-color: #f0fdf4;
        border: 1px solid #bbf7d0;
        color: #166534;
        padding: 6px 14px;
        border-radius: 999px;
        font-size: 13px;
        font-weight: 600;
    }

    /* Dòng thông tin tổng hợp (Meta row) */
    .ksnb-meta-bar {
        font-size: 13px;
        color: #475569;
        padding: 10px 0 12px 0;
        border-bottom: 1px solid #e2e8f0;
        margin-bottom: 14px;
        line-height: 1.6;
    }
    .ksnb-meta-bar b {
        color: #0f172a;
        font-weight: 600;
    }

    /* Hộp ghi chú hướng dẫn nằm ngay trên từng ô nạp file */
    .upload-section-title {
        font-size: 16px;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 12px;
        padding-bottom: 8px;
        border-bottom: 1px solid #f1f5f9;
    }
    .guide-box {
        background: #f8fafc;
        border-radius: 8px;
        padding: 10px 14px;
        margin-top: 10px;
        margin-bottom: 6px;
        border: 1px solid #e2e8f0;
        border-left: 3.5px solid #2563eb;
    }
    .guide-header {
        font-weight: 700;
        color: #1d4ed8;
        font-size: 13px;
        margin-bottom: 3px;
    }
    .guide-text {
        font-size: 12.5px;
        color: #475569;
        line-height: 1.5;
        margin: 0;
    }

    /* Thẻ KPI 3 cột giống mẫu */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 16px;
        margin: 14px 0 20px 0;
    }
    .kpi-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.02);
    }
    .kpi-card.amber { border-left: 4px solid #d97706; }
    .kpi-card.red   { border-left: 4px solid #dc2626; }
    .kpi-card.green { border-left: 4px solid #16a34a; }

    .kpi-title {
        font-size: 11.5px;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 28px;
        font-weight: 700;
        line-height: 1.2;
        margin-bottom: 6px;
    }
    .kpi-value.amber { color: #d97706; }
    .kpi-value.red   { color: #dc2626; }
    .kpi-value.green { color: #16a34a; }
    .kpi-sub {
        font-size: 12.5px;
        color: #475569;
    }

    /* Khung Nhận xét tự động giống mẫu */
    .auto-remark-container {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 18px 20px;
        margin-bottom: 20px;
    }
    .auto-remark-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 14px;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 12px;
    }
    .remark-row {
        padding: 10px 14px;
        border-radius: 6px;
        font-size: 13px;
        line-height: 1.55;
        margin-bottom: 8px;
        display: flex;
        align-items: flex-start;
        gap: 10px;
    }
    .remark-row.info {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        color: #1e293b;
    }
    .remark-row.danger {
        background: #fff5f5;
        border: 1px solid #fed7d7;
        color: #7f1d1d;
    }
    .remark-row.warn {
        background: #fffbeb;
        border: 1px solid #fde68a;
        color: #78350f;
    }
    .remark-row.success {
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        color: #14532d;
    }

    /* Nút bấm chuẩn KSNB */
    div.stButton > button:first-child {
        background: #ffffff !important;
        color: #1e293b !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 6px !important;
        padding: 6px 14px !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important;
    }
    div.stButton > button:first-child:hover {
        background: #f1f5f9 !important;
        border-color: #94a3b8 !important;
    }
    /* Nút Primary (Nạp file / Chạy đối chiếu) */
    div.stButton > button[kind="primary"] {
        background: #1d4ed8 !important;
        color: #ffffff !important;
        border: 1px solid #1e40af !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background: #1e40af !important;
    }

    /* Nút Download nổi bật */
    div.stDownloadButton > button:first-child {
        background: #991b1b !important;
        color: #ffffff !important;
        border: 1px solid #7f1d1d !important;
        border-radius: 6px !important;
        padding: 6px 14px !important;
        font-weight: 600 !important;
        font-size: 13px !important;
    }
    div.stDownloadButton > button:first-child:hover {
        background: #7f1d1d !important;
    }

    /* Tinh chỉnh thanh Tabs ngang */
    button[data-baseweb="tab"] {
        font-size: 13px !important;
        font-weight: 600 !important;
        padding: 8px 14px !important;
    }

    /* Container có viền */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02);
    }
</style>
""",
    unsafe_allow_html=True,
)


# -------------------------------------------------------------
# CÁC HÀM XỬ LÝ DỮ LIỆU TỰ ĐỘNG (GIỮ NGUYÊN LOGIC CHUẨN CỦA BẠN)
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
    exact = next((p for p in pl_sheet_names if p.strip().upper() == bbnt_sheet_name.strip().upper()), None)
    if exact:
        return exact
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
    - Trả về dữ liệu chi tiết từng cuộn và từng nhóm lỗi để xử lý trực tiếp trên Web.
    """
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

    cuon_list = [clean_str_code(ws.cell(row=r, column=4).value).upper() for r in data_rows]
    dict_bv_p_sheet = {}
    dict_bv_kl_sheet = {}
    if bv_raw_list:
        for txt_ma_hang, p, kl in bv_raw_list:
            for m in cuon_list:
                if m and m in txt_ma_hang and m not in dict_bv_p_sheet:
                    dict_bv_p_sheet[m] = p
                    dict_bv_kl_sheet[m] = kl

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

    merges_to_remove = [m for m in ws.merged_cells.ranges if m.min_col >= start_col]
    for m in merges_to_remove:
        ws.unmerge_cells(str(m))

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

    sum_cols = {15: 0.0, 16: 0.0, 17: 0.0, 20: 0.0, 23: 0.0, 27: 0.0, 28: 0.0, 29: 0.0, 30: 0.0, 31: 0.0, 32: 0.0}
    actual_bbnt_sums = {8: 0.0, 9: 0.0, 10: 0.0, 11: 0.0}

    bbnt_internal_errors = []
    coil_error_records = []
    full_web_rows = []

    for r in data_rows:
        stt = ws.cell(row=r, column=1).value
        ngay_nhap = ws.cell(row=r, column=2).value
        kho_hang = ws.cell(row=r, column=3).value
        ma_cuon = clean_str_code(ws.cell(row=r, column=4).value).upper()
        phieu_can_bbnt = clean_str_code(ws.cell(row=r, column=5).value)
        so_xe = ws.cell(row=r, column=6).value
        romoc = ws.cell(row=r, column=7).value
        ghi_chu_bbnt = ws.cell(row=r, column=13).value

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
                ws.cell(row=r, column=c_b).fill = red_fill
                ws.cell(row=r, column=c_b).font = font_error
                bbnt_internal_errors.append({
                    'Đã kiểm': False,
                    'Sheet': s_name,
                    'Vị trí ô': f'{col_let}{r} (STT {stt} - Cuộn {ma_cuon})',
                    'Loại lỗi trên BBNT gốc': f'Nhập định dạng Chữ (Text) tại {c_label}',
                    'Giá trị trên BBNT': repr(raw_v),
                    'Giá trị đúng': f'{to_float(raw_v):,.0f} (Kiểu Số)',
                    'Đơn vị cần làm việc': 'Đơn vị lập BBNT',
                    'Trạng thái xử lý': 'Chưa xử lý',
                    'Lý do / Giải trình trường hợp': '',
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
                'Đã kiểm': False,
                'Sheet': s_name,
                'Vị trí ô': f'K{r} (STT {stt} - Cuộn {ma_cuon})',
                'Loại lỗi trên BBNT gốc': 'Tính sai Chênh lệch (kg) = Cân bàn (J) - Gross (I)',
                'Giá trị trên BBNT': f'{cl_kg_bbnt:,.0f}',
                'Giá trị đúng': f'{expected_cl_kg:,.0f}',
                'Đơn vị cần làm việc': 'Đơn vị lập BBNT',
                'Trạng thái xử lý': 'Chưa xử lý',
                'Lý do / Giải trình trường hợp': '',
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
            elif col_idx == 24:
                cell.alignment = align_right
                cell.number_format = '0.00%'
            else:
                cell.alignment = align_center

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

        # Phân loại nhóm lỗi chi tiết để đưa lên các Tab xử lý trực tiếp trên Web
        err_reasons = []
        err_depts = []
        has_kl_ds_err = False
        has_gate_err = False
        has_erp_err = False

        if doi_chieu_pl != 0:
            has_kl_ds_err = True
            err_depts.append('Packing List / BBNT')
            err_reasons.append(f'Lệch KL Net BBNT ({net_ct:,.0f}) vs PL ({kl_net_pl:,.0f}): {doi_chieu_pl:+,.0f} kg')

        if not dc_phieu:
            has_gate_err = True
            err_depts.append('Bảo vệ')
            if phieu_bv == '':
                err_reasons.append(f'Bảo vệ chưa nhập Mã cuộn vào cột MÃ HÀNG HÓA (Phiếu BBNT: {phieu_can_bbnt})')
            else:
                err_reasons.append(f'Lệch Số phiếu cân: BBNT ({phieu_can_bbnt}) ≠ Bảo vệ ({phieu_bv})')

        if not dc_tc_bv:
            has_gate_err = True
            if 'Bảo vệ / Trạm cân' not in err_depts:
                err_depts.append('Bảo vệ / Trạm cân')
            err_reasons.append(f'Lệch KL Trạm cân ({kl_tc:,.0f}) ≠ KL Bảo vệ ({kl_bv:,.0f})')

        if not dc_tc_bbnt:
            has_gate_err = True
            if 'Trạm cân / BBNT' not in err_depts:
                err_depts.append('Trạm cân / BBNT')
            err_reasons.append(f'Lệch KL Trạm cân ({kl_tc:,.0f}) ≠ Cân bàn BBNT ({can_ban:,.0f})')

        if not dc_dsai:
            has_kl_ds_err = True
            if 'Đơn vị lập BBNT' not in err_depts:
                err_depts.append('Đơn vị lập BBNT')
            err_reasons.append(f'Lệch % dung sai: Tính lại ({chenh_lech_dsai*100:.2f}%) ≠ BBNT ({dsai_bbnt*100:.2f}%)')

        if so_quy_dinh == 'VUOT DUNG SAI':
            has_kl_ds_err = True
            err_depts.append('Vượt dung sai (±0.3%)')
            err_reasons.append(f'Vượt dung sai quy định: {chenh_lech_dsai*100:.2f}% ({chenh_lech_kg:+,.0f} kg)')

        if cl_035 != 0:
            has_erp_err = True
            err_depts.append('Kho ERP (OPM105)')
            err_reasons.append(f'Lệch KL Nhập kho 035: PL ({kl_net_pl:,.0f}) vs Kho 035 ({kl_035:,.0f})')

        status_text = 'OK' if not err_reasons else 'CÓ SAI LỆCH'

        # Lưu dòng đầy đủ 32 cột để tra cứu trực tiếp trên Web
        full_web_rows.append({
            'Sheet': s_name,
            'Dòng Excel': r,
            'STT': stt,
            'Ngày nhập': str(ngay_nhap or ''),
            'Kho': str(kho_hang or ''),
            'Mã số cuộn (D)': ma_cuon,
            'Số phiếu cân BBNT (E)': phieu_can_bbnt,
            'Số xe (F)': str(so_xe or ''),
            'Rơ moóc (G)': str(romoc or ''),
            'KL Net BBNT (H)': net_ct,
            'KL Gross BBNT (I)': gross_ct,
            'KL Cân bàn BBNT (J)': can_ban,
            'CL kg BBNT (K)': cl_kg_bbnt,
            '% Dung sai BBNT (L)': round(dsai_bbnt * 100, 2),
            'Ghi chú BBNT (M)': str(ghi_chu_bbnt or ''),
            'Loại phiếu TC (N)': nhap_xuat,
            'KL Net PL (O)': kl_net_pl,
            'CL Net BBNT - PL (P)': doi_chieu_pl,
            'KL Bảo vệ (Q)': kl_bv,
            'Phiếu cân Bảo vệ (R)': phieu_bv,
            'Khớp Phiếu cân (S)': dc_phieu,
            'KL Trạm cân (T)': kl_tc,
            'Khớp KL TC & BV (U)': dc_tc_bv,
            'Khớp KL TC & BBNT (V)': dc_tc_bbnt,
            'CL KL TC - Gross (W)': chenh_lech_kg,
            '% Dung sai Tính lại (X)': round(chenh_lech_dsai * 100, 2),
            'Khớp % Dung sai (Y)': dc_dsai,
            'Đánh giá ±0.3% (Z)': so_quy_dinh,
            'KL Kho 035 (AA)': kl_035,
            'CL Kho 035 (AB)': cl_035,
            'KL Kho 031 (AC)': kl_031,
            'CL Kho 031 (AD)': cl_031,
            'KL Kho 03G-3H (AE)': kl_03g,
            'CL Kho 03G-3H (AF)': cl_03g,
            'Kết luận nhanh': status_text,
            'Chi tiết lỗi': ' | '.join(err_reasons) if err_reasons else 'Khớp đúng 100%',
        })

        if err_reasons:
            coil_error_records.append({
                'ID_Key': f"{s_name}_{r}_{ma_cuon}",
                'Đã kiểm': False,
                'Sheet': s_name,
                'Dòng Excel': r,
                'STT': stt,
                'Mã số cuộn': ma_cuon,
                'Số phiếu cân BBNT': phieu_can_bbnt,
                'Phiếu Bảo vệ': phieu_bv if phieu_bv else '(Trống)',
                'KL Net BBNT': net_ct,
                'KL Net PL': kl_net_pl,
                'CL Net (BBNT-PL)': doi_chieu_pl,
                'KL Gross BBNT': gross_ct,
                'KL Cân bàn BBNT': can_ban,
                'KL Trạm cân': kl_tc,
                'KL Bảo vệ': kl_bv,
                'CL TC - Gross (kg)': chenh_lech_kg,
                '% Dung sai': round(chenh_lech_dsai * 100, 2),
                'Đánh giá DS': so_quy_dinh,
                'KL Kho 035': kl_035,
                'CL Kho 035': cl_035,
                'KL Kho 031': kl_031,
                'KL Kho 03G-3H': kl_03g,
                'Đơn vị / Khâu cần kiểm tra': ', '.join(dict.fromkeys(err_depts)),
                'Chi tiết lỗi phát hiện': ' | '.join(err_reasons),
                'Trạng thái xử lý': 'Chưa xử lý',
                'Lý do / Giải trình trường hợp': '',
                'has_kl_ds_err': has_kl_ds_err,
                'has_gate_err': has_gate_err,
                'has_erp_err': has_erp_err,
            })

    # Kiểm tra dòng Tổng cộng trên BBNT gốc
    if total_row is not None:
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
                    'Đã kiểm': False,
                    'Sheet': s_name,
                    'Vị trí ô': f'{col_let}{total_row} (Dòng Tổng cộng)',
                    'Loại lỗi trên BBNT gốc': f'Cộng sai dòng Tổng cộng tại {c_label}',
                    'Giá trị trên BBNT': f'{orig_total:,.0f}',
                    'Giá trị đúng': f'{true_total:,.0f} (Lệch {diff_tot:+,.0f} kg)',
                    'Đơn vị cần làm việc': 'Đơn vị lập BBNT (Kiểm tra công thức SUM)',
                    'Trạng thái xử lý': 'Chưa xử lý',
                    'Lý do / Giải trình trường hợp': '',
                })

        for idx in range(len(extra_headers)):
            col_idx = start_col + idx
            cell = ws.cell(row=total_row, column=col_idx)
            cell.border = border_thin
            cell.font = font_total
            if col_idx in sum_cols:
                cell.value = sum_cols[col_idx]
                cell.alignment = align_right
                cell.number_format = '#,##0'

    # Unhide toàn bộ cột
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
        'sum_net_bbnt': actual_bbnt_sums[8],
        'sum_gross_bbnt': actual_bbnt_sums[9],
        'sum_can_ban': actual_bbnt_sums[10],
        'sum_pl': sum_cols[15],
        'sum_tc': sum_cols[20],
        'sum_035': sum_cols[27],
        'bbnt_internal_errors': bbnt_internal_errors,
        'coil_error_records': coil_error_records,
        'full_web_rows': full_web_rows,
    }


def xuat_file_bao_cao_giai_trinh(df_coil_Action, df_form_Action, general_conclusion):
    """Xuất file Excel tổng hợp các trường hợp sai lệch kèm Lý do / Giải trình đã nhập trên Web."""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        if not df_coil_Action.empty:
            cols_export = [c for c in df_coil_Action.columns if c not in ['ID_Key', 'has_kl_ds_err', 'has_gate_err', 'has_erp_err']]
            df_coil_Action[cols_export].to_excel(writer, sheet_name='1_Xu_Ly_Sai_Lech_Cuon', index=False)
        else:
            pd.DataFrame({'Thông báo': ['Không có cuộn sai lệch']}).to_excel(writer, sheet_name='1_Xu_Ly_Sai_Lech_Cuon', index=False)

        if not df_form_Action.empty:
            df_form_Action.to_excel(writer, sheet_name='2_Loi_Bieu_Mau_BBNT', index=False)
        else:
            pd.DataFrame({'Thông báo': ['Biểu mẫu BBNT gốc không có lỗi công thức/định dạng']}).to_excel(writer, sheet_name='2_Loi_Bieu_Mau_BBNT', index=False)

        pd.DataFrame({
            'Nội dung Kết luận & Kiến nghị Kiểm soát nội bộ': [general_conclusion]
        }).to_excel(writer, sheet_name='3_Ket_Luan_Kien_Nghi', index=False)
    return out.getvalue()


# -------------------------------------------------------------
# CHỨC NĂNG CHÍNH & GIAO DIỆN XỬ LÝ TRỰC TIẾP TRÊN WEB
# -------------------------------------------------------------
def doi_chieu_hrc_tung_file():
    # Khởi tạo trạng thái lưu trữ dữ liệu trên Web để thao tác không bị mất khi lọc/nhập lý do
    if 'hrc_processed' not in st.session_state:
        st.session_state.hrc_processed = False
    if 'show_uploader' not in st.session_state:
        st.session_state.show_uploader = True
    if 'general_conclusion' not in st.session_state:
        st.session_state.general_conclusion = (
            "1. Đề nghị Đơn vị lập BBNT kiểm tra và chuẩn hóa lại định dạng số (Number) trên biểu mẫu BBNT, tránh nhập định dạng chữ (Text) làm sai lệch hàm SUM ở dòng Tổng cộng.\n"
            "2. Đề nghị Bộ phận Bảo vệ cập nhật đầy đủ, chính xác Mã số cuộn vào cột Mã hàng hóa và đối chiếu khớp số phiếu cân với Trạm cân.\n"
            "3. Đối với các cuộn vượt dung sai quy định (±0.3%) hoặc lệch khối lượng nhập kho ERP (035/031/03G-3H), các bộ phận liên quan giải trình rõ nguyên nhân tại cột 'Lý do / Giải trình trường hợp'."
        )

    # 1. THANH ĐIỀU HƯỚNG TRÊN CÙNG (GIỐNG ẢNH MẪU)
    top_c1, top_c2 = st.columns([1.4, 8.6])
    with top_c1:
        if st.button('← Quay lại các mảng', use_container_width=True):
            st.switch_page('app_ksnb.py')
    with top_c2:
        st.markdown(
            "<div style='padding-top:6px; font-size:13px; color:#64748b; font-weight:500;'>"
            "Kiểm soát nghiệm thu HRC – Trạm cân – Bảo vệ – Kho ERP"
            "</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<hr style='border:none; border-top:1px solid #e2e8f0; margin:6px 0 14px 0;'>", unsafe_allow_html=True)

    # 2. TIÊU ĐỀ CHÍNH
    st.markdown(
        """
        <div class="ksnb-dept-sub">PHÒNG KIỂM SOÁT NỘI BỘ • KIỂM SOÁT NGUYÊN VẬT LIỆU & NGHIỆM THU HRC</div>
        <h1 class="ksnb-main-title">Đối chiếu nghiệm thu HRC – Trạm cân – Bảo vệ – Kho ERP</h1>
        """,
        unsafe_allow_html=True,
    )

    # 3. THANH ACTION BAR (PILL TRẠNG THÁI + NÚT NẠP FILE + TẢI EXCEL PHỤ LỤC + BÁO CÁO GIẢI TRÌNH)
    if st.session_state.hrc_processed:
        df_err_cur = st.session_state.df_coil_errors
        df_form_cur = st.session_state.df_form_errors
        total_issues = len(df_err_cur) + len(df_form_cur)
        checked_issues = (
            (int(df_err_cur['Đã kiểm'].sum()) if not df_err_cur.empty else 0)
            + (int(df_form_cur['Đã kiểm'].sum()) if not df_form_cur.empty else 0)
        )
        pill_html = (
            f'<div class="session-pill">'
            f'Phiên: {len(st.session_state.all_reports)} Sheet BBNT ({st.session_state.total_coils_all} cuộn) '
            f'· Đã kiểm xử lý {checked_issues}/{total_issues} điểm lưu ý'
            f'</div>'
        )
    else:
        pill_html = (
            '<div class="session-pill" style="background:#f1f5f9; border-color:#cbd5e1; color:#475569;">'
            'Phiên: Chưa nạp dữ liệu · Vui lòng nạp file ở khung bên dưới'
            '</div>'
        )

    act_c1, act_c2, act_c3, act_c4 = st.columns([4.5, 1.6, 2.0, 1.9])
    with act_c1:
        st.markdown(pill_html, unsafe_allow_html=True)
    with act_c2:
        btn_label = "🔼 Ẩn khung nạp file" if st.session_state.show_uploader else "📥 + Nạp / Đổi file"
        if st.button(btn_label, use_container_width=True):
            st.session_state.show_uploader = not st.session_state.show_uploader
            st.rerun()
    with act_c3:
        if st.session_state.hrc_processed:
            st.download_button(
                f"📥 Excel phụ lục ({len(st.session_state.all_reports)} Sheet)",
                data=st.session_state.excel_bytes,
                file_name="BBNT_Nghiem_Thu_Da_Doi_Chieu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.button("📥 Excel phụ lục (Chưa có)", disabled=True, use_container_width=True)
    with act_c4:
        if st.session_state.hrc_processed:
            audit_bytes = xuat_file_bao_cao_giai_trinh(
                st.session_state.df_coil_errors,
                st.session_state.df_form_errors,
                st.session_state.general_conclusion,
            )
            st.download_button(
                "📋 Xuất Báo cáo Giải trình",
                data=audit_bytes,
                file_name="Bao_Cao_Xu_Ly_Giai_Trinh_HRC.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.button("📋 Báo cáo Giải trình", disabled=True, use_container_width=True)

    # 4. KHUNG NẠP 5 FILE NGUỒN (CÓ GHI CHÚ ĐẶT NGAY TRÊN TỪNG Ô NẠP FILE ĐÚNG YÊU CẦU)
    if st.session_state.show_uploader or not st.session_state.hrc_processed:
        with st.container(border=True):
            st.markdown("#### 📥 Nạp Dữ Liệu Nguồn Đối Chiếu (Giữ nguyên hướng dẫn tại từng ô nạp)")
            col_upload1, col_upload2 = st.columns(2, gap="large")

            with col_upload1:
                st.markdown('<div class="upload-section-title">1️⃣ Dữ liệu Nghiệm thu & Trạm cân</div>', unsafe_allow_html=True)

                # Ô 1: BBNT + Ghi chú ngay trên ô nạp
                st.markdown(
                    """
                    <div class="guide-box">
                        <div class="guide-header">📌 1. Biên bản nghiệm thu (BBNT)</div>
                        <div class="guide-text">Đường dẫn: Nguyen Thi Thu Thuong > Biên bản nghiệm thu (BBNT)</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                f_bbnt = st.file_uploader('1. Chọn File Biên bản nghiệm thu (BBNT)', type=['xlsx', 'xls'], key='u_bbnt')

                selected_sheets = []
                if f_bbnt is not None:
                    try:
                        f_bbnt.seek(0)
                        xl_bb = pd.ExcelFile(f_bbnt)
                        sheet_names_bb = xl_bb.sheet_names
                        col_sel_all, _ = st.columns([2, 3])
                        with col_sel_all:
                            select_all = st.checkbox("Chọn tất cả các Sheet", value=True)
                        selected_sheets = st.multiselect(
                            '🎯 Chọn Sheet BBNT cần đối chiếu:',
                            options=sheet_names_bb,
                            default=sheet_names_bb if select_all else ([sheet_names_bb[0]] if sheet_names_bb else []),
                            help='Các sheet không được chọn sẽ tự động bị loại bỏ khỏi file kết quả',
                        )
                    except Exception as e:
                        st.error(f"Lỗi đọc danh sách sheet: {e}")
                else:
                    st.caption("ℹ️ Nạp file BBNT để hiện danh sách sheet cần kiểm tra.")

                # Ô 2: Packing List + Ghi chú ngay trên ô nạp
                st.markdown(
                    """
                    <div class="guide-box">
                        <div class="guide-header">📌 2. Packing List nhà cung cấp</div>
                        <div class="guide-text">Đường dẫn: Nguyen Thi Thu Thuong > Packing List nhà cung cấp</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                f_pl = st.file_uploader('2. Chọn File Packing List nhà cung cấp', type=['xlsx', 'xls'], key='u_pl')

                # Ô 3: Trạm cân + Ghi chú ngay trên ô nạp
                st.markdown(
                    """
                    <div class="guide-box">
                        <div class="guide-header">📌 3. Dữ liệu Trạm cân</div>
                        <div class="guide-text">
                            Truy cập: http://172.17.24.9/#/login (PM4444 / 123456)<br>
                            Quản lý truy cập > Tra cứu trạm cân • Thời gian: 1 tháng rưỡi
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                f_tc = st.file_uploader('3. Chọn File Dữ liệu Trạm cân', type=['xlsx', 'xls'], key='u_tc')

            with col_upload2:
                st.markdown('<div class="upload-section-title">2️⃣ Dữ liệu Bảo vệ & Kho ERP</div>', unsafe_allow_html=True)

                # Ô 4: Bảo vệ + Ghi chú ngay trên ô nạp
                st.markdown(
                    """
                    <div class="guide-box">
                        <div class="guide-header">📌 4. Dữ liệu Nhập Bảo vệ</div>
                        <div class="guide-text">
                            Quản lý truy cập > Tra cứu xuất nhập hàng > Chọn Nhập hàng<br>
                            Thời gian: Giống mốc Tra cứu trạm cân
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                f_bv = st.file_uploader('4. Chọn File Dữ liệu Nhập Bảo vệ', type=['xlsx', 'xls'], key='u_bv')

                # Ô 5: ERP OPM 105 + Ghi chú ngay trên ô nạp
                st.markdown(
                    """
                    <div class="guide-box">
                        <div class="guide-header">📌 5. Báo cáo ERP OPM 105</div>
                        <div class="guide-text">
                            Chọn: HSG_OPM_105_BCKiemSoatThongTinGDTheoLo_GD2<br>
                            Tích chọn kho: 035, 03G, 03H • Nghiệp vụ: Inventory, Job or Schedule, Purchase order<br>
                            Thời gian: 1 tháng cần kiểm tra
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                f_opm = st.file_uploader('5. Chọn File Báo cáo ERP OPM 105', type=['xlsx', 'xls'], key='u_opm')

                st.write("")
                if f_bbnt is not None and selected_sheets:
                    if st.button('🚀 Bắt đầu đối chiếu & Phân tích trực tiếp trên Web', type='primary', use_container_width=True):
                        with st.spinner(f'Đang đối soát {len(selected_sheets)} sheet dữ liệu và dựng bảng làm việc trên Web...'):
                            try:
                                # 1. QUÉT DỮ LIỆU BẢO VỆ
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
                                opm_total_rows = 0
                                if f_opm is not None:
                                    try:
                                        df_opm = read_auto_header(f_opm, ['LOẠI GIAO DỊCH', 'MÃ LÔ'], match_all=True)
                                        if len(df_opm) > 0 and '[' in str(df_opm.iloc[0, 1]):
                                            df_opm = df_opm.iloc[1:].copy()
                                        opm_total_rows = len(df_opm)

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

                                # 4. DANH SÁCH SHEET CỦA PACKING LIST
                                pl_sheet_names = []
                                if f_pl is not None:
                                    try:
                                        f_pl.seek(0)
                                        pl_sheet_names = pd.ExcelFile(f_pl).sheet_names
                                    except Exception:
                                        pl_sheet_names = []

                                # 5. MỞ FILE BBNT GỐC BẰNG OPENPYXL
                                f_bbnt.seek(0)
                                wb = openpyxl.load_workbook(f_bbnt, data_only=True)

                                for s in list(wb.sheetnames):
                                    if s not in selected_sheets:
                                        del wb[s]

                                all_reports = []
                                all_coil_errors = []
                                all_form_errors = []
                                all_full_rows = []

                                for s_name in selected_sheets:
                                    if s_name in wb.sheetnames:
                                        ws = wb[s_name]
                                        matched_pl = match_pl_sheet(s_name, pl_sheet_names)
                                        rep = ghi_va_kiem_tra_worksheet(
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
                                        all_reports.append(rep)
                                        all_coil_errors.extend(rep['coil_error_records'])
                                        all_form_errors.extend(rep['bbnt_internal_errors'])
                                        all_full_rows.extend(rep['full_web_rows'])

                                output = io.BytesIO()
                                wb.save(output)

                                # Lưu vào session_state để làm việc trực tiếp trên Web
                                st.session_state.hrc_processed = True
                                st.session_state.show_uploader = False
                                st.session_state.excel_bytes = output.getvalue()
                                st.session_state.all_reports = all_reports
                                st.session_state.df_coil_errors = pd.DataFrame(all_coil_errors)
                                st.session_state.df_form_errors = pd.DataFrame(all_form_errors)
                                st.session_state.df_full_web = pd.DataFrame(all_full_rows)
                                st.session_state.total_coils_all = sum(r['total_coils'] for r in all_reports)
                                st.session_state.meta_counts = {
                                    'tc_count': len(dict_tc_kl),
                                    'bv_count': len(bv_raw_list),
                                    'opm_count': opm_total_rows,
                                    'pl_sheets': len(pl_sheet_names),
                                }
                                st.rerun()
                            except Exception as e:
                                st.error(f'Lỗi xử lý dữ liệu: {e}')

    # -------------------------------------------------------------
    # 5. KHU VỰC XỬ LÝ TRỰC TIẾP TRÊN WEB (9 TABS GIỐNG MẪU KSNB)
    # -------------------------------------------------------------
    if st.session_state.hrc_processed:
        all_reports = st.session_state.all_reports
        df_coil_err = st.session_state.df_coil_errors
        df_form_err = st.session_state.df_form_errors
        df_full = st.session_state.df_full_web
        meta = st.session_state.meta_counts

        total_coils = max(st.session_state.total_coils_all, 1)
        total_err_coils = len(df_coil_err)
        total_form_errs = len(df_form_err)

        # Đếm số lượng cho từng Tab
        cnt_kl_ds = int(df_coil_err['has_kl_ds_err'].sum()) if not df_coil_err.empty else 0
        cnt_gate = int(df_coil_err['has_gate_err'].sum()) if not df_coil_err.empty else 0
        cnt_erp = int(df_coil_err['has_erp_err'].sum()) if not df_coil_err.empty else 0
        cnt_luu_y = total_err_coils + total_form_errs

        # Dòng Meta tóm tắt giống hệt ảnh mẫu
        sheet_list_str = ", ".join([r['sheet_name'] for r in all_reports])
        st.markdown(
            f"""
            <div class="ksnb-meta-bar">
                Số Sheet BBNT đã kiểm: <b>{len(all_reports)} sheet ({sheet_list_str})</b>  • 
                Tổng cuộn HRC: <b>{st.session_state.total_coils_all:,} cuộn</b>  • 
                Packing List: <b>{meta['pl_sheets']} sheet</b>  • 
                Trạm cân: <b>{meta['tc_count']:,} phiếu</b>  • 
                Bảo vệ: <b>{meta['bv_count']:,} dòng</b>  • 
                ERP OPM 105: <b>{meta['opm_count']:,} giao dịch (Kho 035, 031, 03G, 03H)</b>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Tính 3 chỉ số KPI lớn
        ok_kl_ds = st.session_state.total_coils_all - cnt_kl_ds
        pct_kl_ds = (ok_kl_ds / total_coils) * 100

        ok_gate = st.session_state.total_coils_all - cnt_gate
        pct_gate = (ok_gate / total_coils) * 100

        ok_erp = st.session_state.total_coils_all - cnt_erp
        pct_erp = (ok_erp / total_coils) * 100

        color_kpi1 = "green" if cnt_kl_ds == 0 else ("amber" if pct_kl_ds >= 95 else "red")
        color_kpi2 = "green" if cnt_gate == 0 else ("amber" if pct_gate >= 95 else "red")
        color_kpi3 = "green" if cnt_erp == 0 else ("amber" if pct_erp >= 95 else "red")

        # Khởi tạo 9 Tabs điều hướng có gắn số lượng lỗi
        tabs = st.tabs([
            "1. Tổng quan",
            f"2. Điểm cần lưu ý ({cnt_luu_y})",
            f"3. Lệch KL & Dung sai ({cnt_kl_ds})",
            f"4. Chứng từ – Trạm cân – Bảo vệ ({cnt_gate})",
            f"5. Đối chiếu Kho ERP ({cnt_erp})",
            f"6. Lỗi biểu mẫu BBNT gốc ({total_form_errs})",
            f"7. Bảng đối chiếu đầy đủ ({st.session_state.total_coils_all})",
            f"8. Tổng hợp theo Sheet ({len(all_reports)})",
            "9. Kết luận & Kiến nghị",
        ])

        # =========================================================
        # TAB 1: TỔNG QUAN (KPI CARDS + NHẬN XÉT TỰ ĐỘNG GIỐNG MẪU)
        # =========================================================
        with tabs[0]:
            st.markdown("#### Tổng quan đối chiếu nghiệm thu HRC")
            st.caption(
                "Kiểm tra tự động 4 lớp: (1) Biểu mẫu & công thức nội bộ BBNT gốc (Cột A → M); "
                "(2) Khối lượng Net BBNT so với Packing List & đánh giá dung sai ±0.3%; "
                "(3) Đối chiếu 3 bên BBNT – Phiếu cân Trạm cân – Báo cáo Bảo vệ; "
                "(4) Đối chiếu khối lượng nhập kho ERP OPM 105 (Kho 035, 031, 03G-3H)."
            )

            st.markdown(
                f"""
                <div class="kpi-grid">
                    <div class="kpi-card {color_kpi1}">
                        <div class="kpi-title">KL NGHIỆM THU – PACKING LIST & DUNG SAI (±0.3%)</div>
                        <div class="kpi-value {color_kpi1}">{pct_kl_ds:.1f}%</div>
                        <div class="kpi-sub">{ok_kl_ds:,}/{st.session_state.total_coils_all:,} cuộn khớp KL & trong dung sai ±0.3% · <b>{cnt_kl_ds} cuộn lệch/vượt DS</b></div>
                    </div>
                    <div class="kpi-card {color_kpi2}">
                        <div class="kpi-title">CHỨNG TỪ BBNT – PHIẾU TRẠM CÂN – BẢO VỆ</div>
                        <div class="kpi-value {color_kpi2}">{pct_gate:.1f}%</div>
                        <div class="kpi-sub">{ok_gate:,}/{st.session_state.total_coils_all:,} cuộn khớp số phiếu & KL cân · <b>{cnt_gate} cuộn lệch phiếu/KL</b></div>
                    </div>
                    <div class="kpi-card {color_kpi3}">
                        <div class="kpi-title">LOGIC NHẬP KHO ERP (KHO 035 / 031 / 03G-3H)</div>
                        <div class="kpi-value {color_kpi3}">{pct_erp:.1f}%</div>
                        <div class="kpi-sub">{ok_erp:,}/{st.session_state.total_coils_all:,} cuộn khớp KL Packing List & ERP · <b>{cnt_erp} cuộn lệch kho</b></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            # XÂY DỰNG KHUNG NHẬN XÉT TỰ ĐỘNG GIỐNG ẢNH MẪU
            sum_net_all = sum(r['sum_net_bbnt'] for r in all_reports) / 1000.0
            sum_tc_all = sum(r['sum_tc'] for r in all_reports) / 1000.0

            remarks_html = f"""
            <div class="auto-remark-container">
                <div class="auto-remark-header">
                    <span>Nhận xét tự động</span>
                    <span style="font-size:12px; font-weight:500; color:#64748b;">Tổng hợp từ {len(all_reports)} Sheet BBNT</span>
                </div>
                <div class="remark-row info">
                    <span>↗</span>
                    <div>
                        <b>Tổng quy mô kiểm tra: {st.session_state.total_coils_all:,} cuộn HRC</b> thuộc {len(all_reports)} sheet BBNT
                        (Tổng KL Net BBNT: <b>{sum_net_all:,.1f} tấn</b> · Tổng KL qua Trạm cân: <b>{sum_tc_all:,.1f} tấn</b>),
                        tỷ lệ khớp toàn diện đạt <b>{((st.session_state.total_coils_all - total_err_coils)/total_coils)*100:.1f}%</b>.
                    </div>
                </div>
            """

            if total_form_errs > 0:
                remarks_html += f"""
                <div class="remark-row warn">
                    <span>⚠️</span>
                    <div>
                        <b>Lỗi nội bộ ngay trên biểu mẫu BBNT gốc (Cột A → M): {total_form_errs} vị trí</b> – Phát hiện ô nhập định dạng Chữ (Text) hoặc sai công thức trừ/hàm SUM dòng Tổng cộng.
                        Code đã giữ nguyên số gốc và tô đỏ trên file Excel (xem và xử lý tại tab <b>6. Lỗi biểu mẫu BBNT gốc</b>).
                    </div>
                </div>
                """

            if not df_coil_err.empty:
                # Thống kê chi tiết từng loại lỗi để in ra giống ảnh
                vuot_ds_df = df_coil_err[df_coil_err['Đánh giá DS'] == 'VUOT DUNG SAI']
                if len(vuot_ds_df) > 0:
                    tong_lech_kg = vuot_ds_df['CL TC - Gross (kg)'].sum()
                    remarks_html += f"""
                    <div class="remark-row danger">
                        <span>🔴</span>
                        <div>
                            <b>KL Trạm cân lệch KL Gross BBNT vượt dung sai quy định (±0.3%): {len(vuot_ds_df)} cuộn · Tổng lệch {tong_lech_kg:+,.0f} kg</b>
                            – Cần kiểm tra lại phiếu cân thực tế và biên bản nghiệm thu tại tab <b>3. Lệch KL & Dung sai</b>.
                        </div>
                    </div>
                    """

                lech_pl_df = df_coil_err[df_coil_err['CL Net (BBNT-PL)'] != 0]
                if len(lech_pl_df) > 0:
                    remarks_html += f"""
                    <div class="remark-row danger">
                        <span>🔴</span>
                        <div>
                            <b>Lệch KL Net giữa BBNT và Packing List nhà cung cấp: {len(lech_pl_df)} cuộn</b>
                            – Tổng chênh lệch <b>{lech_pl_df['CL Net (BBNT-PL)'].sum():+,.0f} kg</b>.
                        </div>
                    </div>
                    """

                bv_missing_df = df_coil_err[df_coil_err['Phiếu Bảo vệ'] == '(Trống)']
                if len(bv_missing_df) > 0:
                    remarks_html += f"""
                    <div class="remark-row danger">
                        <span>🔴</span>
                        <div>
                            <b>Bảo vệ chưa cập nhật Mã số cuộn vào cột MÃ HÀNG HÓA hoặc lệch số phiếu cân: {len(bv_missing_df)} cuộn</b>
                            – Dẫn đến không truy vết tự động được phiếu cân từ báo cáo Bảo vệ (xem tab <b>4. Chứng từ – Trạm cân – Bảo vệ</b>).
                        </div>
                    </div>
                    """

                if cnt_erp > 0:
                    remarks_html += f"""
                    <div class="remark-row warn">
                        <span>🟠</span>
                        <div>
                            <b>Lệch khối lượng nhập kho ERP (Kho 035) so với Packing List: {cnt_erp} cuộn</b>
                            – Kiểm tra giao dịch PO Receipt / Return to Vendor trên báo cáo OPM 105 tại tab <b>5. Đối chiếu Kho ERP</b>.
                        </div>
                    </div>
                    """

                remarks_html += f"""
                <div class="remark-row info">
                    <span>📌</span>
                    <div>
                        <b>Tổng cộng {cnt_luu_y} điểm cần lưu ý</b> ({total_err_coils} cuộn có sai lệch đối chiếu, {total_form_errs} lỗi công thức/định dạng BBNT)
                        – Chuyển sang tab <b>2. Điểm cần lưu ý</b> để phân loại trạng thái và nhập lý do giải trình trực tiếp trên Web.
                    </div>
                </div>
                """
            else:
                remarks_html += """
                <div class="remark-row success">
                    <span>✅</span>
                    <div>
                        <b>Hoàn tất đối soát: 100% các cuộn HRC đều khớp chính xác</b> giữa BBNT, Packing List, Báo cáo Bảo vệ, Trạm cân và Kho ERP!
                    </div>
                </div>
                """

            remarks_html += "</div>"
            st.markdown(remarks_html, unsafe_allow_html=True)

        # =========================================================
        # TAB 2: ĐIỂM CẦN LƯU Ý (THAO TÁC XỬ LÝ & NHẬP LÝ DO TRÊN WEB)
        # =========================================================
        with tabs[1]:
            st.markdown("#### 🛠️ Bảng Xử Lý & Giải Trình Các Điểm Cần Lưu Ý Trực Tiếp Trên Web")
            st.caption(
                "Bạn có thể lọc theo Sheet hoặc Bộ phận liên quan, tích vào ô **Đã kiểm**, chọn **Trạng thái xử lý** và gõ trực tiếp **Lý do / Giải trình trường hợp** vào bảng bên dưới."
            )

            if df_coil_err.empty and df_form_err.empty:
                st.success("🎉 Không có bất kỳ điểm lưu ý hay sai lệch nào trong phiên kiểm tra này!")
            else:
                if not df_coil_err.empty:
                    f_c1, f_c2, f_c3 = st.columns([2.5, 3, 4.5])
                    with f_c1:
                        sheet_filter = st.multiselect(
                            "Lọc theo Sheet BBNT:",
                            options=df_coil_err['Sheet'].unique().tolist(),
                            default=df_coil_err['Sheet'].unique().tolist(),
                            key="flt_sheet_t2",
                        )
                    with f_c2:
                        status_filter = st.multiselect(
                            "Lọc theo Trạng thái xử lý:",
                            options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'],
                            default=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'],
                            key="flt_status_t2",
                        )
                    with f_c3:
                        kw_search = st.text_input("🔍 Tìm nhanh Mã cuộn / Số phiếu cân / Nội dung lỗi:", key="kw_t2")

                    mask = df_coil_err['Sheet'].isin(sheet_filter) & df_coil_err['Trạng thái xử lý'].isin(status_filter)
                    if kw_search.strip():
                        kw = kw_search.strip().upper()
                        mask = mask & (
                            df_coil_err['Mã số cuộn'].astype(str).str.upper().str.contains(kw)
                            | df_coil_err['Số phiếu cân BBNT'].astype(str).str.upper().str.contains(kw)
                            | df_coil_err['Chi tiết lỗi phát hiện'].astype(str).str.upper().str.contains(kw)
                            | df_coil_err['Đơn vị / Khâu cần kiểm tra'].astype(str).str.upper().str.contains(kw)
                        )

                    cols_t2 = [
                        'Đã kiểm', 'Sheet', 'Dòng Excel', 'STT', 'Mã số cuộn', 'Số phiếu cân BBNT',
                        'Đơn vị / Khâu cần kiểm tra', 'Chi tiết lỗi phát hiện',
                        'Trạng thái xử lý', 'Lý do / Giải trình trường hợp',
                    ]
                    edited_t2 = st.data_editor(
                        df_coil_err.loc[mask, cols_t2],
                        use_container_width=True,
                        hide_index=True,
                        num_rows="fixed",
                        key="editor_tab2_coils",
                        column_config={
                            'Đã kiểm': st.column_config.CheckboxColumn('Đã kiểm', width='small'),
                            'Sheet': st.column_config.TextColumn('Sheet', disabled=True, width='small'),
                            'Dòng Excel': st.column_config.NumberColumn('Dòng', disabled=True, width='small'),
                            'STT': st.column_config.NumberColumn('STT', disabled=True, width='small'),
                            'Mã số cuộn': st.column_config.TextColumn('Mã số cuộn', disabled=True),
                            'Số phiếu cân BBNT': st.column_config.TextColumn('Phiếu cân BBNT', disabled=True),
                            'Đơn vị / Khâu cần kiểm tra': st.column_config.TextColumn('Khâu cần kiểm tra', disabled=True),
                            'Chi tiết lỗi phát hiện': st.column_config.TextColumn('Chi tiết lỗi phát hiện', disabled=True, width='large'),
                            'Trạng thái xử lý': st.column_config.SelectboxColumn(
                                'Trạng thái xử lý',
                                options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'],
                                required=True,
                                width='medium',
                            ),
                            'Lý do / Giải trình trường hợp': st.column_config.TextColumn(
                                '✍️ Lý do / Giải trình trường hợp (Nhập trực tiếp)',
                                width='large',
                            ),
                        },
                    )
                    # Đồng bộ ngược lại vào session_state
                    st.session_state.df_coil_errors.loc[mask, ['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']] = (
                        edited_t2[['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']]
                    )

        # =========================================================
        # TAB 3: LỆCH KHỐI LƯỢNG & DUNG SAI (±0.3%)
        # =========================================================
        with tabs[2]:
            st.markdown("#### ⚖️ Kiểm Soát Chênh Lệch Khối Lượng & Dung Sai (±0.3%)")
            st.caption("Danh sách các cuộn bị lệch KL Net giữa BBNT với Packing List, hoặc có tỷ lệ chênh lệch giữa Trạm cân và Gross BBNT vượt ±0.3%.")
            if df_coil_err.empty or cnt_kl_ds == 0:
                st.success("✅ Toàn bộ các cuộn đều khớp khối lượng Packing List và nằm trong dung sai ±0.3%.")
            else:
                mask_klds = df_coil_err['has_kl_ds_err'] == True
                cols_t3 = [
                    'Đã kiểm', 'Sheet', 'STT', 'Mã số cuộn', 'KL Net BBNT', 'KL Net PL', 'CL Net (BBNT-PL)',
                    'KL Gross BBNT', 'KL Trạm cân', 'CL TC - Gross (kg)', '% Dung sai', 'Đánh giá DS',
                    'Trạng thái xử lý', 'Lý do / Giải trình trường hợp',
                ]
                edited_t3 = st.data_editor(
                    df_coil_err.loc[mask_klds, cols_t3],
                    use_container_width=True,
                    hide_index=True,
                    key="editor_tab3_klds",
                    column_config={
                        'Đã kiểm': st.column_config.CheckboxColumn('Đã kiểm', width='small'),
                        'KL Net BBNT': st.column_config.NumberColumn('KL Net BBNT', format="%d", disabled=True),
                        'KL Net PL': st.column_config.NumberColumn('KL Net PL', format="%d", disabled=True),
                        'CL Net (BBNT-PL)': st.column_config.NumberColumn('Lệch Net (kg)', format="%+d", disabled=True),
                        'KL Gross BBNT': st.column_config.NumberColumn('Gross BBNT', format="%d", disabled=True),
                        'KL Trạm cân': st.column_config.NumberColumn('KL Trạm cân', format="%d", disabled=True),
                        'CL TC - Gross (kg)': st.column_config.NumberColumn('TC - Gross (kg)', format="%+d", disabled=True),
                        '% Dung sai': st.column_config.NumberColumn('% Dung sai', format="%.2f%%", disabled=True),
                        'Trạng thái xử lý': st.column_config.SelectboxColumn(
                            'Trạng thái xử lý',
                            options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'],
                            required=True,
                        ),
                        'Lý do / Giải trình trường hợp': st.column_config.TextColumn('✍️ Lý do / Giải trình trường hợp', width='large'),
                    },
                )
                st.session_state.df_coil_errors.loc[mask_klds, ['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']] = (
                    edited_t3[['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']]
                )

        # =========================================================
        # TAB 4: CHỨNG TỪ BBNT – TRẠM CÂN – BẢO VỆ
        # =========================================================
        with tabs[3]:
            st.markdown("#### 🚛 Đối Chiếu Chéo 3 Bên: BBNT – Trạm Cân – Báo Cáo Bảo Vệ")
            st.caption("Phát hiện các trường hợp Bảo vệ bỏ trống mã cuộn trong cột Mã hàng hóa, ghi sai số phiếu cân hoặc lệch khối lượng giữa 3 bên.")
            if df_coil_err.empty or cnt_gate == 0:
                st.success("✅ 100% số phiếu cân và khối lượng cân khớp hoàn toàn giữa BBNT, Trạm cân và Bảo vệ.")
            else:
                mask_gate = df_coil_err['has_gate_err'] == True
                cols_t4 = [
                    'Đã kiểm', 'Sheet', 'STT', 'Mã số cuộn', 'Số phiếu cân BBNT', 'Phiếu Bảo vệ',
                    'KL Cân bàn BBNT', 'KL Trạm cân', 'KL Bảo vệ', 'Chi tiết lỗi phát hiện',
                    'Trạng thái xử lý', 'Lý do / Giải trình trường hợp',
                ]
                edited_t4 = st.data_editor(
                    df_coil_err.loc[mask_gate, cols_t4],
                    use_container_width=True,
                    hide_index=True,
                    key="editor_tab4_gate",
                    column_config={
                        'Đã kiểm': st.column_config.CheckboxColumn('Đã kiểm', width='small'),
                        'KL Cân bàn BBNT': st.column_config.NumberColumn('Cân bàn BBNT', format="%d", disabled=True),
                        'KL Trạm cân': st.column_config.NumberColumn('KL Trạm cân', format="%d", disabled=True),
                        'KL Bảo vệ': st.column_config.NumberColumn('KL Bảo vệ', format="%d", disabled=True),
                        'Trạng thái xử lý': st.column_config.SelectboxColumn(
                            'Trạng thái xử lý',
                            options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'],
                            required=True,
                        ),
                        'Lý do / Giải trình trường hợp': st.column_config.TextColumn('✍️ Lý do / Giải trình trường hợp', width='large'),
                    },
                )
                st.session_state.df_coil_errors.loc[mask_gate, ['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']] = (
                    edited_t4[['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']]
                )

        # =========================================================
        # TAB 5: ĐỐI CHIẾU KHO ERP (OPM 105: KHO 035, 031, 03G-3H)
        # =========================================================
        with tabs[4]:
            st.markdown("#### 🏭 Đối Chiếu Nhập Kho ERP OPM 105 (Kho 035, 031 Intransit, 03G-3H)")
            st.caption("Kiểm tra khối lượng PO Receipt trừ Return to Vendor tại Kho 035, 031, 03G-3H so với KL Net trên Packing List.")
            if df_coil_err.empty or cnt_erp == 0:
                st.success("✅ Toàn bộ các cuộn đều khớp khối lượng nhập kho 035 với Packing List.")
            else:
                mask_erp = df_coil_err['has_erp_err'] == True
                cols_t5 = [
                    'Đã kiểm', 'Sheet', 'STT', 'Mã số cuộn', 'KL Net PL', 'KL Kho 035', 'CL Kho 035',
                    'KL Kho 031', 'KL Kho 03G-3H', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp',
                ]
                edited_t5 = st.data_editor(
                    df_coil_err.loc[mask_erp, cols_t5],
                    use_container_width=True,
                    hide_index=True,
                    key="editor_tab5_erp",
                    column_config={
                        'Đã kiểm': st.column_config.CheckboxColumn('Đã kiểm', width='small'),
                        'KL Net PL': st.column_config.NumberColumn('KL Net PL', format="%d", disabled=True),
                        'KL Kho 035': st.column_config.NumberColumn('KL Kho 035', format="%d", disabled=True),
                        'CL Kho 035': st.column_config.NumberColumn('Lệch Kho 035', format="%+d", disabled=True),
                        'KL Kho 031': st.column_config.NumberColumn('KL Kho 031', format="%d", disabled=True),
                        'KL Kho 03G-3H': st.column_config.NumberColumn('KL Kho 03G-3H', format="%d", disabled=True),
                        'Trạng thái xử lý': st.column_config.SelectboxColumn(
                            'Trạng thái xử lý',
                            options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'],
                            required=True,
                        ),
                        'Lý do / Giải trình trường hợp': st.column_config.TextColumn('✍️ Lý do / Giải trình trường hợp', width='large'),
                    },
                )
                st.session_state.df_coil_errors.loc[mask_erp, ['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']] = (
                    edited_t5[['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']]
                )

        # =========================================================
        # TAB 6: LỖI BIỂU MẪU BBNT GỐC (CỘT A -> M)
        # =========================================================
        with tabs[5]:
            st.markdown("#### ⚠️ Lỗi Phát Hiện Ngay Trên Biểu Mẫu BBNT Gốc (Cột A → M)")
            st.caption(
                "Các lỗi do người lập BBNT nhập số dưới dạng Chữ (Text) làm hàm SUM bỏ sót, hoặc tính sai cột K = J - I. "
                "Trên file Excel xuất ra, các ô này được giữ nguyên giá trị gốc và tô nền đỏ để bạn làm bằng chứng làm việc."
            )
            if df_form_err.empty:
                st.success("✅ Biểu mẫu BBNT gốc hoàn toàn chuẩn xác, không có lỗi định dạng Text hay cộng sai dòng Tổng cộng.")
            else:
                edited_form = st.data_editor(
                    df_form_err,
                    use_container_width=True,
                    hide_index=True,
                    key="editor_tab6_form",
                    column_config={
                        'Đã kiểm': st.column_config.CheckboxColumn('Đã kiểm', width='small'),
                        'Sheet': st.column_config.TextColumn('Sheet', disabled=True),
                        'Vị trí ô': st.column_config.TextColumn('Vị trí ô Excel', disabled=True),
                        'Loại lỗi trên BBNT gốc': st.column_config.TextColumn('Loại lỗi phát hiện', disabled=True),
                        'Giá trị trên BBNT': st.column_config.TextColumn('Giá trị gốc trên BBNT', disabled=True),
                        'Giá trị đúng': st.column_config.TextColumn('Giá trị tính đúng', disabled=True),
                        'Đơn vị cần làm việc': st.column_config.TextColumn('Đơn vị cần làm việc', disabled=True),
                        'Trạng thái xử lý': st.column_config.SelectboxColumn(
                            'Trạng thái xử lý',
                            options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'],
                            required=True,
                        ),
                        'Lý do / Giải trình trường hợp': st.column_config.TextColumn('✍️ Lý do / Giải trình trường hợp', width='large'),
                    },
                )
                st.session_state.df_form_errors = edited_form

        # =========================================================
        # TAB 7: BẢNG ĐỐI CHIẾU ĐẦY ĐỦ 32 CỘT TRÊN WEB (CÓ TÔ ĐỎ Ô LỖI)
        # =========================================================
        with tabs[6]:
            st.markdown("#### 📋 Bảng Dữ Liệu Đối Chiếu Đầy Đủ Trên Web (Tương đương Cột A → AF trên Excel)")
            c_f1, c_f2, c_f3 = st.columns([2.5, 2.5, 5])
            with c_f1:
                sel_sheet_full = st.selectbox(
                    "Chọn Sheet BBNT:",
                    options=["Tất cả Sheet"] + [r['sheet_name'] for r in all_reports],
                    key="sel_full_sheet",
                )
            with c_f2:
                sel_status_full = st.selectbox(
                    "Trạng thái cuộn:",
                    options=["Tất cả cuộn", "Chỉ hiện cuộn CÓ SAI LỆCH", "Chỉ hiện cuộn OK"],
                    key="sel_full_status",
                )
            with c_f3:
                search_full = st.text_input("🔍 Tìm theo Mã số cuộn hoặc Số phiếu cân:", key="search_full_input")

            df_view = df_full.copy()
            if sel_sheet_full != "Tất cả Sheet":
                df_view = df_view[df_view['Sheet'] == sel_sheet_full]
            if sel_status_full == "Chỉ hiện cuộn CÓ SAI LỆCH":
                df_view = df_view[df_view['Kết luận nhanh'] == 'CÓ SAI LỆCH']
            elif sel_status_full == "Chỉ hiện cuộn OK":
                df_view = df_view[df_view['Kết luận nhanh'] == 'OK']
            if search_full.strip():
                q = search_full.strip().upper()
                df_view = df_view[
                    df_view['Mã số cuộn (D)'].astype(str).str.upper().str.contains(q)
                    | df_view['Số phiếu cân BBNT (E)'].astype(str).str.upper().str.contains(q)
                ]

            def highlight_errors(row):
                styles = [''] * len(row)
                if row['Kết luận nhanh'] == 'CÓ SAI LỆCH':
                    styles[row.index.get_loc('Kết luận nhanh')] = 'background-color: #fecaca; color: #991b1b; font-weight: bold;'
                    styles[row.index.get_loc('Mã số cuộn (D)')] = 'background-color: #fef2f2; color: #991b1b; font-weight: bold;'
                if row['CL Net BBNT - PL (P)'] != 0:
                    styles[row.index.get_loc('CL Net BBNT - PL (P)')] = 'background-color: #fecaca; color: #991b1b; font-weight: bold;'
                if row['Khớp Phiếu cân (S)'] is False:
                    styles[row.index.get_loc('Khớp Phiếu cân (S)')] = 'background-color: #fecaca; color: #991b1b; font-weight: bold;'
                if row['Khớp KL TC & BV (U)'] is False:
                    styles[row.index.get_loc('Khớp KL TC & BV (U)')] = 'background-color: #fecaca; color: #991b1b; font-weight: bold;'
                if row['Khớp KL TC & BBNT (V)'] is False:
                    styles[row.index.get_loc('Khớp KL TC & BBNT (V)')] = 'background-color: #fecaca; color: #991b1b; font-weight: bold;'
                if row['Khớp % Dung sai (Y)'] is False:
                    styles[row.index.get_loc('Khớp % Dung sai (Y)')] = 'background-color: #fecaca; color: #991b1b; font-weight: bold;'
                if row['Đánh giá ±0.3% (Z)'] == 'VUOT DUNG SAI':
                    styles[row.index.get_loc('Đánh giá ±0.3% (Z)')] = 'background-color: #fecaca; color: #991b1b; font-weight: bold;'
                if row['CL Kho 035 (AB)'] != 0:
                    styles[row.index.get_loc('CL Kho 035 (AB)')] = 'background-color: #fecaca; color: #991b1b; font-weight: bold;'
                return styles

            st.dataframe(
                df_view.style.apply(highlight_errors, axis=1),
                use_container_width=True,
                hide_index=True,
                height=520,
            )

        # =========================================================
        # TAB 8: TỔNG HỢP THEO TỪNG SHEET BBNT
        # =========================================================
        with tabs[7]:
            st.markdown("#### 📊 Bảng Tổng Hợp Kết Quả Theo Từng Sheet BBNT")
            sheet_summary_rows = []
            for r in all_reports:
                sheet_summary_rows.append({
                    'Sheet BBNT': r['sheet_name'],
                    'Sheet Packing List ghép': r['matched_pl'],
                    'Tổng số cuộn': r['total_coils'],
                    'Số cuộn khớp 100%': r['ok_coils_count'],
                    'Số cuộn có sai lệch': r['error_coils_count'],
                    'Lỗi biểu mẫu BBNT gốc': len(r['bbnt_internal_errors']),
                    'Tổng KL Net BBNT (kg)': r['sum_net_bbnt'],
                    'Tổng KL Net PL (kg)': r['sum_pl'],
                    'Tổng KL Trạm cân (kg)': r['sum_tc'],
                    'Tổng KL Nhập kho 035 (kg)': r['sum_035'],
                })
            st.dataframe(
                pd.DataFrame(sheet_summary_rows),
                use_container_width=True,
                hide_index=True,
            )

        # =========================================================
        # TAB 9: KẾT LUẬN & KIẾN NGHỊ KIỂM SOÁT
        # =========================================================
        with tabs[8]:
            st.markdown("#### 📝 Soạn Thảo Kết Luận & Kiến Nghị Kiểm Soát Nội Bộ")
            st.caption("Nội dung kết luận tại đây sẽ được xuất kèm vào Sheet 3 của file **Báo cáo Giải trình**.")
            st.session_state.general_conclusion = st.text_area(
                "Nội dung Kết luận & Kiến nghị:",
                value=st.session_state.general_conclusion,
                height=180,
            )

            st.markdown("##### 📌 Tóm tắt các trường hợp đã ghi nhận Lý do / Giải trình:")
            if not df_coil_err.empty:
                explained_df = st.session_state.df_coil_errors[
                    (st.session_state.df_coil_errors['Lý do / Giải trình trường hợp'].astype(str).str.strip() != '')
                    | (st.session_state.df_coil_errors['Trạng thái xử lý'] != 'Chưa xử lý')
                ]
                if not explained_df.empty:
                    st.dataframe(
                        explained_df[[
                            'Đã kiểm', 'Sheet', 'Mã số cuộn', 'Số phiếu cân BBNT',
                            'Đơn vị / Khâu cần kiểm tra', 'Chi tiết lỗi phát hiện',
                            'Trạng thái xử lý', 'Lý do / Giải trình trường hợp',
                        ]],
                        use_container_width=True,
                        hide_index=True,
                    )
                else:
                    st.info("Chưa có cuộn nào được nhập Lý do / Giải trình. Bạn có thể nhập trực tiếp tại Tab 2, 3, 4 hoặc 5.")


doi_chieu_hrc_tung_file()
