import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN PHONG CÁCH APPLE MINIMALIST PASTEL
# -------------------------------------------------------------
st.set_page_config(
    page_title='Kiểm soát nội bộ - Đối Chiếu Nghiệm Thu HRC',
    page_icon='🛡️',
    layout='wide',
)

st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    /* Ẩn sidebar mặc định */
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }

    /* Nền trang sáng sủa, tinh tế */
    .stApp {
        background-color: #f4f6f9;
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #1e293b;
    }

    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 3rem !important;
        max-width: 96% !important;
    }

    /* Tiêu đề phong cách Apple */
    .apple-dept {
        font-size: 11.5px;
        font-weight: 700;
        color: #8b5cf6;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 4px;
    }
    .apple-title {
        font-size: 28px;
        font-weight: 700;
        color: #0f172a;
        margin: 0 0 14px 0;
        letter-spacing: -0.02em;
    }

    /* Thẻ Pill trạng thái bo tròn mượt mà */
    .apple-pill {
        display: inline-flex;
        align-items: center;
        background: #ede9fe;
        border: 1px solid #ddd6fe;
        color: #6d28d9;
        padding: 6px 16px;
        border-radius: 999px;
        font-size: 13px;
        font-weight: 600;
    }

    /* Thanh Meta bar */
    .apple-meta {
        background: #ffffff;
        border-radius: 14px;
        padding: 12px 20px;
        font-size: 13px;
        color: #475569;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.01);
        margin-bottom: 18px;
        line-height: 1.6;
    }
    .apple-meta b {
        color: #0f172a;
    }

    /* Hộp hướng dẫn nạp file bo tròn */
    .upload-box-title {
        font-size: 16px;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 12px;
    }
    .guide-card {
        background: #faf5ff;
        border-radius: 12px;
        padding: 12px 16px;
        margin-top: 10px;
        margin-bottom: 8px;
        border: 1px solid #f3e8ff;
        border-left: 4px solid #8b5cf6;
    }
    .guide-card-head {
        font-weight: 700;
        color: #7c3aed;
        font-size: 13px;
        margin-bottom: 4px;
    }
    .guide-card-txt {
        font-size: 12.5px;
        color: #581c87;
        line-height: 1.5;
        margin: 0;
    }

    /* Thẻ KPI 3 cột bo tròn kiểu Apple */
    .kpi-grid-apple {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 18px;
        margin: 16px 0 22px 0;
    }
    .kpi-apple-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 20px 24px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.02);
        transition: transform 0.2s ease;
    }
    .kpi-apple-card.amber { border-top: 4px solid #f59e0b; }
    .kpi-apple-card.red   { border-top: 4px solid #ef4444; }
    .kpi-apple-card.green { border-top: 4px solid #10b981; }

    .kpi-apple-title {
        font-size: 11.5px;
        font-weight: 700;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 8px;
    }
    .kpi-apple-val {
        font-size: 30px;
        font-weight: 700;
        line-height: 1.1;
        margin-bottom: 8px;
    }
    .kpi-apple-val.amber { color: #d97706; }
    .kpi-apple-val.red   { color: #dc2626; }
    .kpi-apple-val.green { color: #059669; }
    .kpi-apple-sub {
        font-size: 13px;
        color: #475569;
    }

    /* Khung Nhận xét tự động */
    .remark-apple-container {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 20px 24px;
        margin-bottom: 22px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.02);
    }
    .remark-apple-head {
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 15px;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 14px;
    }
    .remark-apple-row {
        padding: 12px 16px;
        border-radius: 10px;
        font-size: 13.5px;
        line-height: 1.6;
        margin-bottom: 10px;
        display: flex;
        align-items: flex-start;
        gap: 12px;
    }
    .remark-apple-row.info { background: #f8fafc; border: 1px solid #e2e8f0; color: #1e293b; }
    .remark-apple-row.danger { background: #fef2f2; border: 1px solid #fecaca; color: #7f1d1d; }
    .remark-apple-row.warn { background: #fffbeb; border: 1px solid #fde68a; color: #78350f; }
    .remark-apple-row.success { background: #f0fdf4; border: 1px solid #bbf7d0; color: #14532d; }

    /* Nút bấm Apple Pastel Style */
    div.stButton > button:first-child {
        background: #ffffff !important;
        color: #334155 !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
        padding: 8px 16px !important;
        font-weight: 600 !important;
        font-size: 13.5px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.02) !important;
    }
    div.stButton > button:first-child:hover {
        background: #f8fafc !important;
        border-color: #94a3b8 !important;
        color: #0f172a !important;
    }
    /* Nút Primary (Tím pastel Apple) */
    div.stButton > button[kind="primary"] {
        background: #7c3aed !important;
        color: #ffffff !important;
        border: 1px solid #6d28d9 !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background: #6d28d9 !important;
    }

    /* Nút Download nổi bật */
    div.stDownloadButton > button:first-child {
        background: #dc2626 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 8px 16px !important;
        font-weight: 600 !important;
        font-size: 13.5px !important;
    }
    div.stDownloadButton > button:first-child:hover {
        background: #b91c1c !important;
    }

    /* Container bo tròn */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff;
        border-radius: 16px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 16px rgba(0,0,0,0.02);
    }
</style>
""",
    unsafe_allow_html=True,
)


# -------------------------------------------------------------
# CÁC HÀM XỬ LÝ DỮ LIỆU TỰ ĐỘNG (GIỮ NGUYÊN LOGIC CHUẨN)
# -------------------------------------------------------------
def read_auto_header(file_obj, keywords, sheet_name=0, max_scan=25, match_all=False):
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
    if val is None or val == "":
        return default
    try:
        return float(str(val).replace(',', '').strip())
    except Exception:
        return default


def clean_str_code(val):
    if val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r'\.0$', '', s)
    return s.strip()


def normalize_sheet_key(name):
    s = str(name).upper()
    s = re.sub(r'\b(BBNT|PACKING LIST|PACKINGLIST|PL|BC|BAO CAO)\b', '', s)
    s = re.sub(r'[^A-Z0-9]', '', s)
    return s


def match_pl_sheet(bbnt_sheet_name, pl_sheet_names):
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
        'Loại phiếu trạm cân\n(Nhập/Xuất)',
        'Khối lượng Net\ntheo Packing List',
        'Chênh lệch KL Net\n(BBNT - Packing List)',
        'Khối lượng theo\nBáo cáo Bảo vệ',
        'Số phiếu cân theo\nBáo cáo Bảo vệ',
        'Khớp Số phiếu cân\n(BBNT & Bảo vệ)',
        'Khối lượng theo\nDữ liệu Trạm cân',
        'Khớp Khối lượng\n(Trạm cân & Bảo vệ)',
        'Khớp Khối lượng\n(Trạm cân & BBNT)',
        'Chênh lệch KL (kg)\n(Trạm cân - Gross BBNT)',
        'Tỷ lệ chênh lệch\ndung sai (%)',
        'Khớp Tỷ lệ dung sai\nvới BBNT',
        'Đánh giá dung sai\n(Quy định ±0.3%)',
        'KL Nhập kho 035\n(ERP OPM105)',
        'Chênh lệch KL Kho 035\n(PL - Kho 035)',
        'KL Nhập kho 031\n(Intransit)',
        'Chênh lệch KL Kho 031\n(PL - Kho 031)',
        'KL Nhập kho 03G-3H\n(ERP OPM105)',
        'Chênh lệch KL Kho 03G-3H\n(PL - Kho 03G-3H)',
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

    start_col = 14

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
            nhap_xuat, kl_net_pl, doi_chieu_pl, kl_bv, phieu_bv, dc_phieu,
            kl_tc, dc_tc_bv, dc_tc_bbnt, chenh_lech_kg, chenh_lech_dsai, dc_dsai,
            so_quy_dinh, kl_035, cl_035, kl_031, cl_031, kl_03g, cl_03g,
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
                    'Đơn vị cần làm việc': 'Đơn vị lập BBNT',
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

    for c_i in range(1, start_col + len(extra_headers) + 2):
        col_let = get_column_letter(c_i)
        ws.column_dimensions[col_let].hidden = False

    # Đánh giá phân loại sheet dựa trên số lượng lỗi
    total_sheet_errs = len(coil_error_records) + len(bbnt_internal_errors)
    sheet_evaluation = "✅ An toàn (Đạt chuẩn)" if total_sheet_errs == 0 else ("⚠️ Cần lưu ý" if total_sheet_errs <= 3 else "🔴 Cần xử lý gấp")

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
        'sheet_evaluation': sheet_evaluation,
    }


def xuat_file_bao_cao_giai_trinh(df_coil_Action, df_form_Action, general_conclusion):
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
            pd.DataFrame({'Thông báo': ['Biểu mẫu BBNT gốc không có lỗi']}).to_excel(writer, sheet_name='2_Loi_Bieu_Mau_BBNT', index=False)

        pd.DataFrame({
            'Nội dung Kết luận & Kiến nghị Kiểm soát nội bộ': [general_conclusion]
        }).to_excel(writer, sheet_name='3_Ket_Luan_Kien_Nghi', index=False)
    return out.getvalue()


# -------------------------------------------------------------
# CHỨC NĂNG CHÍNH & GIAO DIỆN APPLE MINIMALIST PASTEL
# -------------------------------------------------------------
def doi_chieu_hrc_tung_file():
    if 'hrc_processed' not in st.session_state:
        st.session_state.hrc_processed = False
    if 'show_uploader' not in st.session_state:
        st.session_state.show_uploader = True
    if 'general_conclusion' not in st.session_state:
        st.session_state.general_conclusion = (
            "1. Đề nghị Đơn vị lập BBNT kiểm tra và chuẩn hóa lại định dạng số (Number) trên biểu mẫu BBNT.\n"
            "2. Đề nghị Bộ phận Bảo vệ cập nhật đầy đủ, chính xác Mã số cuộn vào cột Mã hàng hóa và đối chiếu khớp số phiếu cân.\n"
            "3. Các bộ phận giải trình rõ nguyên nhân đối với các cuộn vượt dung sai hoặc lệch kho ERP."
        )

    top_c1, top_c2 = st.columns([1.4, 8.6])
    with top_c1:
        if st.button('← Quay lại các mảng', use_container_width=True):
            st.switch_page('app_ksnb.py')
    with top_c2:
        st.markdown(
            "<div style='padding-top:6px; font-size:13px; color:#64748b; font-weight:500;'>"
            "Hệ thống Kiểm soát nội bộ Nguyên vật liệu & Nghiệm thu HRC"
            "</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<hr style='border:none; border-top:1px solid #e2e8f0; margin:6px 0 14px 0;'>", unsafe_allow_html=True)

    st.markdown(
        """
        <div class="apple-dept">PHÒNG KIỂM SOÁT NỘI BỘ • KSNB</div>
        <h1 class="apple-title">Đối chiếu nghiệm thu HRC – Trạm cân – Bảo vệ – Kho ERP</h1>
        """,
        unsafe_allow_html=True,
    )

    if st.session_state.hrc_processed:
        df_err_cur = st.session_state.df_coil_errors
        df_form_cur = st.session_state.df_form_errors
        total_issues = len(df_err_cur) + len(df_form_cur)
        checked_issues = (
            (int(df_err_cur['Đã kiểm'].sum()) if not df_err_cur.empty else 0)
            + (int(df_form_cur['Đã kiểm'].sum()) if not df_form_cur.empty else 0)
        )
        pill_html = (
            f'<div class="apple-pill">'
            f'🛡️ Phiên: {len(st.session_state.all_reports)} Sheet BBNT ({st.session_state.total_coils_all} cuộn) '
            f'· Đã xử lý {checked_issues}/{total_issues} điểm lưu ý'
            f'</div>'
        )
    else:
        pill_html = (
            '<div class="apple-pill" style="background:#f1f5f9; border-color:#cbd5e1; color:#475569;">'
            '🛡️ Phiên: Chưa nạp dữ liệu · Vui lòng tải file ở bên dưới'
            '</div>'
        )

    act_c1, act_c2, act_c3, act_c4 = st.columns([4.2, 1.8, 2.0, 2.0])
    with act_c1:
        st.markdown(pill_html, unsafe_allow_html=True)
    with act_c2:
        btn_label = "🔼 Thu gọn nạp file" if st.session_state.show_uploader else "📥 + Nạp / Đổi file"
        if st.button(btn_label, use_container_width=True):
            st.session_state.show_uploader = not st.session_state.show_uploader
            st.rerun()
    with act_c3:
        if st.session_state.hrc_processed:
            st.download_button(
                f"📥 Tải Excel phụ lục",
                data=st.session_state.excel_bytes,
                file_name="BBNT_Nghiem_Thu_Da_Doi_Chieu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.button("📥 Tải Excel phụ lục", disabled=True, use_container_width=True)
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

    # KHUNG NẠP FILE (CÓ GHI CHÚ ĐẶT NGAY TRÊN TỪNG Ô NẠP)
    if st.session_state.show_uploader or not st.session_state.hrc_processed:
        with st.container(border=True):
            st.markdown('<div class="upload-box-title">📥 Nạp Dữ Liệu Nguồn Đối Chiếu</div>', unsafe_allow_html=True)
            col_upload1, col_upload2 = st.columns(2, gap="large")

            with col_upload1:
                st.markdown(
                    """
                    <div class="guide-card">
                        <div class="guide-card-head">📌 1. Biên bản nghiệm thu (BBNT)</div>
                        <div class="guide-card-txt">Đường dẫn: Nguyen Thi Thu Thuong > Biên bản nghiệm thu (BBNT)</div>
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
                        )
                    except Exception as e:
                        st.error(f"Lỗi đọc danh sách sheet: {e}")
                else:
                    st.caption("ℹ️ Nạp file BBNT để hiển thị danh sách sheet.")

                st.markdown(
                    """
                    <div class="guide-card">
                        <div class="guide-card-head">📌 2. Packing List nhà cung cấp</div>
                        <div class="guide-card-txt">Đường dẫn: Nguyen Thi Thu Thuong > Packing List nhà cung cấp</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                f_pl = st.file_uploader('2. Chọn File Packing List nhà cung cấp', type=['xlsx', 'xls'], key='u_pl')

                st.markdown(
                    """
                    <div class="guide-card">
                        <div class="guide-card-head">📌 3. Dữ liệu Trạm cân</div>
                        <div class="guide-card-txt">
                            Truy cập: http://172.17.24.9/#/login (PM4444 / 123456)<br>
                            Quản lý truy cập > Tra cứu trạm cân • Thời gian: 1 tháng rưỡi
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                f_tc = st.file_uploader('3. Chọn File Dữ liệu Trạm cân', type=['xlsx', 'xls'], key='u_tc')

            with col_upload2:
                st.markdown(
                    """
                    <div class="guide-card">
                        <div class="guide-card-head">📌 4. Dữ liệu Nhập Bảo vệ</div>
                        <div class="guide-card-txt">
                            Quản lý truy cập > Tra cứu xuất nhập hàng > Chọn Nhập hàng<br>
                            Thời gian: Giống mốc Tra cứu trạm cân
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                f_bv = st.file_uploader('4. Chọn File Dữ liệu Nhập Bảo vệ', type=['xlsx', 'xls'], key='u_bv')

                st.markdown(
                    """
                    <div class="guide-card">
                        <div class="guide-card-head">📌 5. Báo cáo ERP OPM 105</div>
                        <div class="guide-card-txt">
                            Chọn: HSG_OPM_105_BCKiemSoatThongTinGDTheoLo_GD2<br>
                            Tích chọn kho: 035, 03G, 03H • Nghiệp vụ: Inventory, Job, Purchase order
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                f_opm = st.file_uploader('5. Chọn File Báo cáo ERP OPM 105', type=['xlsx', 'xls'], key='u_opm')

                st.write("")
                if f_bbnt is not None and selected_sheets:
                    if st.button('🚀 Bắt đầu đối chiếu & Phân tích tự động', type='primary', use_container_width=True):
                        with st.spinner(f'Đang đối soát {len(selected_sheets)} sheet dữ liệu...'):
                            try:
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

                                pl_sheet_names = []
                                if f_pl is not None:
                                    try:
                                        f_pl.seek(0)
                                        pl_sheet_names = pd.ExcelFile(f_pl).sheet_names
                                    except Exception:
                                        pl_sheet_names = []

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
                                            ws, s_name, f_pl, matched_pl, bv_raw_list,
                                            dict_tc_kl, dict_tc_phanloai, opm_035_po, opm_035_rtv,
                                            opm_031, opm_03g_po, opm_03g_rtv,
                                        )
                                        all_reports.append(rep)
                                        all_coil_errors.extend(rep['coil_error_records'])
                                        all_form_errors.extend(rep['bbnt_internal_errors'])
                                        all_full_rows.extend(rep['full_web_rows'])

                                output = io.BytesIO()
                                wb.save(output)

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
    # KHU VỰC HIỂN THỊ KẾT QUẢ & 9 TABS ĐIỀU HƯỚNG
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

        cnt_kl_ds = int(df_coil_err['has_kl_ds_err'].sum()) if not df_coil_err.empty else 0
        cnt_gate = int(df_coil_err['has_gate_err'].sum()) if not df_coil_err.empty else 0
        cnt_erp = int(df_coil_err['has_erp_err'].sum()) if not df_coil_err.empty else 0
        cnt_luu_y = total_err_coils + total_form_errs

        sheet_list_str = ", ".join([r['sheet_name'] for r in all_reports])
        st.markdown(
            f"""
            <div class="apple-meta">
                📊 Số Sheet BBNT kiểm tra: <b>{len(all_reports)} sheet ({sheet_list_str})</b>  • 
                Tổng cuộn: <b>{st.session_state.total_coils_all:,} cuộn</b>  • 
                Packing List: <b>{meta['pl_sheets']} sheet</b>  • 
                Trạm cân: <b>{meta['tc_count']:,} phiếu</b>  • 
                Bảo vệ: <b>{meta['bv_count']:,} dòng</b>  • 
                ERP OPM 105: <b>{meta['opm_count']:,} giao dịch</b>
            </div>
            """,
            unsafe_allow_html=True,
        )

        ok_kl_ds = st.session_state.total_coils_all - cnt_kl_ds
        pct_kl_ds = (ok_kl_ds / total_coils) * 100

        ok_gate = st.session_state.total_coils_all - cnt_gate
        pct_gate = (ok_gate / total_coils) * 100

        ok_erp = st.session_state.total_coils_all - cnt_erp
        pct_erp = (ok_erp / total_coils) * 100

        color_kpi1 = "green" if cnt_kl_ds == 0 else ("amber" if pct_kl_ds >= 95 else "red")
        color_kpi2 = "green" if cnt_gate == 0 else ("amber" if pct_gate >= 95 else "red")
        color_kpi3 = "green" if cnt_erp == 0 else ("amber" if pct_erp >= 95 else "red")

        tabs = st.tabs([
            "1. Tổng quan",
            f"2. Điểm cần lưu ý ({cnt_luu_y})",
            f"3. Lệch KL & Dung sai ({cnt_kl_ds})",
            f"4. Trạm cân – Bảo vệ ({cnt_gate})",
            f"5. Đối chiếu Kho ERP ({cnt_erp})",
            f"6. Lỗi biểu mẫu BBNT ({total_form_errs})",
            f"7. Bảng đầy đủ ({st.session_state.total_coils_all})",
            f"8. Tổng hợp & Đánh giá Sheet ({len(all_reports)})",
            "9. Kết luận & Kiến nghị",
        ])

        # TAB 1: TỔNG QUAN + BIỂU ĐỒ TRỰC QUAN
        with tabs[0]:
            st.markdown("#### 🌟 Tổng quan đối chiếu nghiệm thu HRC")
            st.markdown(
                f"""
                <div class="kpi-grid-apple">
                    <div class="kpi-apple-card {color_kpi1}">
                        <div class="kpi-apple-title">KL & DUNG SAI ±0.3%</div>
                        <div class="kpi-apple-val {color_kpi1}">{pct_kl_ds:.1f}%</div>
                        <div class="kpi-apple-sub">{ok_kl_ds:,}/{st.session_state.total_coils_all:,} cuộn khớp · <b>{cnt_kl_ds} cuộn lệch</b></div>
                    </div>
                    <div class="kpi-apple-card {color_kpi2}">
                        <div class="kpi-apple-title">CHỨNG TỪ BBNT – TRẠM CÂN – BẢO VỆ</div>
                        <div class="kpi-apple-val {color_kpi2}">{pct_gate:.1f}%</div>
                        <div class="kpi-apple-sub">{ok_gate:,}/{st.session_state.total_coils_all:,} cuộn khớp · <b>{cnt_gate} cuộn lệch phiếu</b></div>
                    </div>
                    <div class="kpi-apple-card {color_kpi3}">
                        <div class="kpi-apple-title">NHẬP KHO ERP (035 / 031 / 03G)</div>
                        <div class="kpi-apple-val {color_kpi3}">{pct_erp:.1f}%</div>
                        <div class="kpi-apple-sub">{ok_erp:,}/{st.session_state.total_coils_all:,} cuộn khớp · <b>{cnt_erp} cuộn lệch kho</b></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            # BỔ SUNG BIỂU ĐỒ TRỰC QUAN (CHART)
            st.markdown("##### 📈 Biểu đồ phân bổ cuộn HRC (Khớp đúng vs Có sai lệch) theo từng Sheet")
            chart_data = []
            for r in all_reports:
                chart_data.append({
                    'Sheet': r['sheet_name'],
                    'Khớp đúng 100%': r['ok_coils_count'],
                    'Có sai lệch': r['error_coils_count'],
                })
            df_chart = pd.DataFrame(chart_data).set_index('Sheet')
            st.bar_chart(df_chart, color=["#10b981", "#ef4444"])

            # NHẬN XÉT TỰ ĐỘNG
            sum_net_all = sum(r['sum_net_bbnt'] for r in all_reports) / 1000.0
            sum_tc_all = sum(r['sum_tc'] for r in all_reports) / 1000.0

            remarks_html = f"""
            <div class="remark-apple-container">
                <div class="remark-apple-head">
                    <span>💡 Nhận xét tự động từ hệ thống KSNB</span>
                    <span style="font-size:12px; font-weight:600; color:#7c3aed;">{len(all_reports)} Sheet BBNT</span>
                </div>
                <div class="remark-apple-row info">
                    <span>📌</span>
                    <div>
                        Tổng quy mô: <b>{st.session_state.total_coils_all:,} cuộn HRC</b> (Tổng Net: <b>{sum_net_all:,.1f} tấn</b> · Trạm cân: <b>{sum_tc_all:,.1f} tấn</b>).
                        Tỷ lệ tuân thủ toàn hệ thống đạt <b>{((st.session_state.total_coils_all - total_err_coils)/total_coils)*100:.1f}%</b>.
                    </div>
                </div>
            """
            if total_form_errs > 0:
                remarks_html += f"""
                <div class="remark-apple-row warn">
                    <span>⚠️</span>
                    <div>
                        <b>Lỗi biểu mẫu BBNT gốc (Cột A → M): {total_form_errs} vị trí</b> – Sai định dạng Text hoặc sai công thức tổng. Đã tô đỏ trên file Excel (xem tab <b>6. Lỗi biểu mẫu BBNT</b>).
                    </div>
                </div>
                """
            if total_err_coils > 0:
                remarks_html += f"""
                <div class="remark-apple-row danger">
                    <span>🔴</span>
                    <div>
                        <b>Phát hiện tổng số {total_err_coils} cuộn có sai lệch đối chiếu 4 bên</b>. Đề nghị các bộ phận kiểm tra và nhập giải trình trực tiếp tại tab <b>2. Điểm cần lưu ý</b>.
                    </div>
                </div>
                """
            else:
                remarks_html += """
                <div class="remark-apple-row success">
                    <span>✅</span>
                    <div><b>Tuyệt vời! Toàn bộ cuộn HRC trong phiên đều khớp 100%.</b></div>
                </div>
                """
            remarks_html += "</div>"
            st.markdown(remarks_html, unsafe_allow_html=True)

        # TAB 2: ĐIỂM CẦN LƯU Ý (XỬ LÝ TRỰC TIẾP TRÊN WEB)
        with tabs[1]:
            st.markdown("#### 🛠️ Xử Lý & Nhập Giải Trình Các Điểm Cần Lưu Ý")
            if df_coil_err.empty and df_form_err.empty:
                st.success("🎉 Không có điểm lưu ý nào!")
            else:
                if not df_coil_err.empty:
                    f_c1, f_c2, f_c3 = st.columns([2.5, 3, 4.5])
                    with f_c1:
                        sheet_filter = st.multiselect("Lọc Sheet:", options=df_coil_err['Sheet'].unique().tolist(), default=df_coil_err['Sheet'].unique().tolist(), key="flt_s2")
                    with f_c2:
                        status_filter = st.multiselect("Lọc Trạng thái:", options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'], default=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'], key="flt_st2")
                    with f_c3:
                        kw_search = st.text_input("🔍 Tìm nhanh Mã cuộn / Lỗi:", key="kw_t2")

                    mask = df_coil_err['Sheet'].isin(sheet_filter) & df_coil_err['Trạng thái xử lý'].isin(status_filter)
                    if kw_search.strip():
                        kw = kw_search.strip().upper()
                        mask = mask & (df_coil_err['Mã số cuộn'].astype(str).str.upper().str.contains(kw) | df_coil_err['Chi tiết lỗi phát hiện'].astype(str).str.upper().str.contains(kw))

                    cols_t2 = ['Đã kiểm', 'Sheet', 'STT', 'Mã số cuộn', 'Số phiếu cân BBNT', 'Đơn vị / Khâu cần kiểm tra', 'Chi tiết lỗi phát hiện', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']
                    edited_t2 = st.data_editor(
                        df_coil_err.loc[mask, cols_t2],
                        use_container_width=True, hide_index=True, key="ed_t2",
                        column_config={
                            'Đã kiểm': st.column_config.CheckboxColumn('Đã kiểm', width='small'),
                            'Trạng thái xử lý': st.column_config.SelectboxColumn('Trạng thái', options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'], required=True),
                            'Lý do / Giải trình trường hợp': st.column_config.TextColumn('✍️ Nhập Lý do / Giải trình trực tiếp', width='large'),
                        },
                    )
                    st.session_state.df_coil_errors.loc[mask, ['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']] = edited_t2[['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']]

        # TAB 3: LỆCH KL & DUNG SAI
        with tabs[2]:
            st.markdown("#### ⚖️ Kiểm Soát Chênh Lệch KL & Dung Sai (±0.3%)")
            if df_coil_err.empty or cnt_kl_ds == 0:
                st.success("✅ Đạt chuẩn dung sai ±0.3%.")
            else:
                mask_klds = df_coil_err['has_kl_ds_err'] == True
                cols_t3 = ['Đã kiểm', 'Sheet', 'STT', 'Mã số cuộn', 'KL Net BBNT', 'KL Net PL', 'CL Net (BBNT-PL)', 'KL Trạm cân', '% Dung sai', 'Đánh giá DS', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']
                edited_t3 = st.data_editor(df_coil_err.loc[mask_klds, cols_t3], use_container_width=True, hide_index=True, key="ed_t3",
                    column_config={'Trạng thái xử lý': st.column_config.SelectboxColumn('Trạng thái', options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'], required=True), 'Lý do / Giải trình trường hợp': st.column_config.TextColumn('✍️ Giải trình', width='large')})
                st.session_state.df_coil_errors.loc[mask_klds, ['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']] = edited_t3[['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']]

        # TAB 4: TRẠM CÂN - BẢO VỆ
        with tabs[3]:
            st.markdown("#### 🚛 Đối Chiếu Chéo 3 Bên: BBNT – Trạm Cân – Bảo Vệ")
            if df_coil_err.empty or cnt_gate == 0:
                st.success("✅ Khớp 100% phiếu cân và bảo vệ.")
            else:
                mask_gate = df_coil_err['has_gate_err'] == True
                cols_t4 = ['Đã kiểm', 'Sheet', 'STT', 'Mã số cuộn', 'Số phiếu cân BBNT', 'Phiếu Bảo vệ', 'KL Trạm cân', 'Chi tiết lỗi phát hiện', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']
                edited_t4 = st.data_editor(df_coil_err.loc[mask_gate, cols_t4], use_container_width=True, hide_index=True, key="ed_t4",
                    column_config={'Trạng thái xử lý': st.column_config.SelectboxColumn('Trạng thái', options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'], required=True), 'Lý do / Giải trình trường hợp': st.column_config.TextColumn('✍️ Giải trình', width='large')})
                st.session_state.df_coil_errors.loc[mask_gate, ['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']] = edited_t4[['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']]

        # TAB 5: KHO ERP
        with tabs[4]:
            st.markdown("#### 🏭 Đối Chiếu Nhập Kho ERP OPM 105")
            if df_coil_err.empty or cnt_erp == 0:
                st.success("✅ Khớp kho ERP.")
            else:
                mask_erp = df_coil_err['has_erp_err'] == True
                cols_t5 = ['Đã kiểm', 'Sheet', 'STT', 'Mã số cuộn', 'KL Net PL', 'KL Kho 035', 'CL Kho 035', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']
                edited_t5 = st.data_editor(df_coil_err.loc[mask_erp, cols_t5], use_container_width=True, hide_index=True, key="ed_t5",
                    column_config={'Trạng thái xử lý': st.column_config.SelectboxColumn('Trạng thái', options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'], required=True), 'Lý do / Giải trình trường hợp': st.column_config.TextColumn('✍️ Giải trình', width='large')})
                st.session_state.df_coil_errors.loc[mask_erp, ['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']] = edited_t5[['Đã kiểm', 'Trạng thái xử lý', 'Lý do / Giải trình trường hợp']]

        # TAB 6: LỖI BIỂU MẪU BBNT GỐC
        with tabs[5]:
            st.markdown("#### ⚠️ Lỗi Phát Hiện Trên Biểu Mẫu BBNT Gốc (Cột A → M)")
            if df_form_err.empty:
                st.success("✅ Biểu mẫu chuẩn xác.")
            else:
                edited_form = st.data_editor(df_form_err, use_container_width=True, hide_index=True, key="ed_t6",
                    column_config={'Trạng thái xử lý': st.column_config.SelectboxColumn('Trạng thái', options=['Chưa xử lý', 'Đang xác minh', 'Đã giải trình (Hợp lệ)', 'Yêu cầu điều chỉnh'], required=True), 'Lý do / Giải trình trường hợp': st.column_config.TextColumn('✍️ Giải trình', width='large')})
                st.session_state.df_form_errors = edited_form

        # TAB 7: BẢNG ĐẦY ĐỦ
        with tabs[6]:
            st.markdown("#### 📋 Bảng Dữ Liệu Đối Chiếu Đầy Đủ (Cột A → AF)")
            c_f1, c_f2, c_f3 = st.columns([2.5, 2.5, 5])
            with c_f1:
                sel_sheet_full = st.selectbox("Sheet:", options=["Tất cả Sheet"] + [r['sheet_name'] for r in all_reports], key="s_sh")
            with c_f2:
                sel_status_full = st.selectbox("Trạng thái:", options=["Tất cả cuộn", "Chỉ hiện cuộn CÓ SAI LỆCH", "Chỉ hiện cuộn OK"], key="s_st")
            with c_f3:
                search_full = st.text_input("🔍 Tìm mã cuộn:", key="s_kw")

            df_view = df_full.copy()
            if sel_sheet_full != "Tất cả Sheet":
                df_view = df_view[df_view['Sheet'] == sel_sheet_full]
            if sel_status_full == "Chỉ hiện cuộn CÓ SAI LỆCH":
                df_view = df_view[df_view['Kết luận nhanh'] == 'CÓ SAI LỆCH']
            elif sel_status_full == "Chỉ hiện cuộn OK":
                df_view = df_view[df_view['Kết luận nhanh'] == 'OK']
            if search_full.strip():
                q = search_full.strip().upper()
                df_view = df_view[df_view['Mã số cuộn (D)'].astype(str).str.upper().str.contains(q)]

            st.dataframe(df_view, use_container_width=True, hide_index=True, height=520)

        # TAB 8: TỔNG HỢP & ĐÁNH GIÁ TỪNG SHEET (CÓ Ô TỔNG LỖI & ĐÁNH GIÁ RÕ RÀNG)
        with tabs[7]:
            st.markdown("#### 📊 Tổng Hợp Lỗi & Đánh Giá Phân Loại Theo Từng Sheet BBNT")
            st.caption("Bảng tổng hợp điểm lưu ý, tổng khối lượng và đánh giá trạng thái kiểm soát của từng sheet.")
            sheet_summary_rows = []
            for r in all_reports:
                sheet_summary_rows.append({
                    'Sheet BBNT': r['sheet_name'],
                    'Sheet Packing List': r['matched_pl'],
                    'Tổng cuộn': r['total_coils'],
                    'Khớp 100%': r['ok_coils_count'],
                    'Cuộn sai lệch': r['error_coils_count'],
                    'Lỗi biểu mẫu BBNT': len(r['bbnt_internal_errors']),
                    'Tổng lỗi (Sheet)': r['error_coils_count'] + len(r['bbnt_internal_errors']),
                    'Đánh giá KSNB': r['sheet_evaluation'],
                    'Tổng KL Net BBNT (kg)': r['sum_net_bbnt'],
                })
            df_summary_sheet = pd.DataFrame(sheet_summary_rows)
            st.dataframe(df_summary_sheet, use_container_width=True, hide_index=True)

        # TAB 9: KẾT LUẬN & KIẾN NGHỊ
        with tabs[8]:
            st.markdown("#### 📝 Soạn Thảo Kết Luận & Kiến Nghị Kiểm Soát")
            st.session_state.general_conclusion = st.text_area("Nội dung kết luận:", value=st.session_state.general_conclusion, height=180)


doi_chieu_hrc_tung_file()
