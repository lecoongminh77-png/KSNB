import datetime
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
    page_title='Kiểm Tra Định Mức Tồn Kho & CLC - KSNB',
    page_icon='🛡️',
    layout='wide',
)

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
        <h1 class="apple-main-title">⚖️ Kiểm Tra Định Mức Tồn Kho & CLC</h1>
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
# DANH MỤC CỐ ĐỊNH (ĐÃ NHÚNG TRỰC TIẾP, KHÔNG CẦN SHEET PHỤ)
# -------------------------------------------------------------
NM_KHAC_LIST = [
    'HS-NHONHOI', 'HS-BINHDUONG', 'NHUNG-XB-XK', 'NHUNGNONG-XB',
    'HS-BINH DUONG', 'HS-NHUNGNONG', 'HS-BINH DUONG ', 'HS-NHUA',
    'HS-NGHEAN', 'NHUNG', 'NHUNGNONG', 'NHUNG-XK', 'NHONHOA-OT'
]

SUB_KHO_033_ACTIVE = ('033-01', '033-02')


# -------------------------------------------------------------
# CÁC HÀM HỖ TRỢ & LOGIC TÍNH TOÁN KSNB (GIỮ NGUYÊN 100% LOGIC GỐC)
# -------------------------------------------------------------
def to_float(val, default=0.0):
    if val is None or val == '':
        return default
    try:
        return float(str(val).replace(',', '').strip())
    except Exception:
        return default


def clean_str_code(val):
    if val is None:
        return ''
    s = str(val).strip()
    s = re.sub(r'\.0$', '', s)
    return s.strip()


def parse_erp_date(val):
    """Chuyển đổi ngày tháng từ ERP (24-FEB-2026 hoặc datetime) sang datetime.date."""
    if val is None or val == '':
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = str(val).strip()
    for fmt in ('%d-%b-%Y', '%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None


def tinh_logic_ksnb_dong(
    ten_hang, phan_loai, ma_lo, kl_net, dv_kd, danh_ten, ngay_nhap, kho_cn, ngay_chot
):
    """
    Thực thi chính xác 100% logic công thức KSNB trên từng dòng dữ liệu
    để vừa đối chiếu tô đỏ trên Excel, vừa hiển thị báo cáo nhanh trên Web.
    """
    chuoi = f"{clean_str_code(ma_lo)}-{int(kl_net) if float(kl_net).is_integer() else kl_net}-{clean_str_code(phan_loai)}"

    # 1. Số ngày tồn kho ($P$12 - BC14)
    d_nhap = parse_erp_date(ngay_nhap)
    so_ngay_ton = (ngay_chot - d_nhap).days if (d_nhap and ngay_chot) else 0

    # 2. Tuổi tồn kho
    if so_ngay_ton < 91:
        tuoi_tk = 'Dưới 90 ngày'
    elif so_ngay_ton < 121:
        tuoi_tk = 'Từ 91 đến 120 ngày'
    elif so_ngay_ton < 181:
        tuoi_tk = 'Từ 121 đến 180 ngày'
    elif so_ngay_ton < 361:
        tuoi_tk = 'Từ 181 đến 360 ngày'
    else:
        tuoi_tk = 'Trên 361 ngày'

    # 3. Đánh số phân loại (Cột P)
    pl_str = clean_str_code(phan_loai)
    so_phan_loai = 1 if pl_str in ('1', 'A2') else 2

    # 4. Đánh tên đơn vị kinh doanh (Cột L - Giữ nguyên 100% công thức IF của bạn)
    ax = '' if dv_kd is None else str(dv_kd)
    ay = '' if danh_ten is None else str(danh_ten)
    ay_up = ay.upper()

    if 'BPDT' in ay_up:
        danh_ten_dv = 'DT_BPDT'
    elif ay.startswith('DTHS') or ay.startswith('DLHS') or ay == 'ĐTHS':
        danh_ten_dv = 'DTHS'
    elif ay.startswith('CN') or ay.startswith('VPDH'):
        danh_ten_dv = 'HTPP'
    elif ax == 'VPDH_MN_PKDTT' and ay in ('KDMN01', 'PKDTT', 'KDTT-MN', 'P.KDTT_MN'):
        danh_ten_dv = 'KDTT'
    elif ax in ('VPDH_MN_NM_BANNGOAI', 'VPDH_MT_NM_BANNGOAI') and ay != '':
        danh_ten_dv = 'Khách hàng ngoài trúng thầu'
    elif ('LỢP XƯỞNG' in ay_up) or ('LOPXUONG' in ay_up):
        danh_ten_dv = 'NM Lợp xưởng'
    elif ax == 'PKDDA_HSH' and ay in ('PKDDA_HSH', 'PKD_HSHome'):
        danh_ten_dv = 'PKDDA_HSH'
    elif ax in ('HSG_PXK', 'VPDH_MN_NM') and ay in ('XK.HSG_OU', 'XK-L2', 'XK.HSG', 'SMC-XK'):
        danh_ten_dv = 'XK'
    elif ay == 'BAOBI':
        danh_ten_dv = 'BTP- BAO BÌ'
    elif (('MM' in ay_up) or ('XBPM' in ay_up)) and ('CXL' not in ay_up):
        danh_ten_dv = 'BTP- Mạ màu & Xẻ băng '
    elif ('VLXD' in ay_up) and ('CXL' not in ay_up):
        danh_ten_dv = 'VLXD'
    elif ay in NM_KHAC_LIST:
        danh_ten_dv = 'Các NM Khác'
    else:
        danh_ten_dv = 'NM THSPM'

    # 5. Đánh loại TÍNH / LOẠI (Cột M - Giữ nguyên 100% công thức IF của bạn)
    sub_033_active = clean_str_code(kho_cn) in SUB_KHO_033_ACTIVE
    cond_loai = (
        ('BPDT' in ay_up)
        or (ay.startswith('DTHS') or ay == 'ĐTHS' or ay.startswith('DLHS'))
        or (ay.startswith('CN') or ay.startswith('VPDH'))
        or (ax == 'VPDH_MN_PKDTT' and ay in ('KDMN01', 'PKDTT', 'KDTT-MN', 'P.KDTT_MN'))
        or (ax == 'VPDH_MN_NM_BANNGOAI' and ay != '')
        or (ax == 'VPDH_MN_NM' and ay in ('NM-LỢP XƯỞNG', 'NM-LOPXUONG'))
        or (ax == 'PKDDA_HSH' and ay in ('PKDDA_HSH', 'PKD_HSHome'))
        or (ax in ('HSG_PXK', 'VPDH_MN_NM') and ay in ('XK.HSG_OU', 'SMC-XK'))
        or sub_033_active
    )
    danh_loai = 'LOẠI' if cond_loai else 'TÍNH'

    # 6. Mức tính phạt CLC (Cột K - Giữ nguyên 100% công thức IF của bạn)
    ten_up = '' if ten_hang is None else str(ten_hang).upper()
    if danh_ten_dv == 'NM Lợp xưởng' or so_ngay_ton < 91 or (('HOA SEN GOLD' in ten_up) and so_ngay_ton < 180):
        muc_clc = 'LOẠI'
    elif so_ngay_ton < 181:
        muc_clc = 'M1'
    else:
        muc_clc = 'M2'

    return {
        'chuoi': chuoi,
        'muc_clc': muc_clc,
        'danh_ten_dv': danh_ten_dv,
        'danh_loai': danh_loai,
        'tuoi_tk': tuoi_tk,
        'so_ngay_ton': so_ngay_ton,
        'so_phan_loai': so_phan_loai,
    }


def tao_file_kiem_tra_dinh_muc(f_erp, f_pcu):
    """
    Đọc file 003.xlsx (ERP) và file Chốt phạt tồn kho của P.Cung ứng,
    xuất ra workbook gồm đúng 4 sheet chính: TTR, ĐỐI CHIẾU, KSNB, P.Cung ứng.
    """
    # ---------------------------------------------------------
    # 1. ĐỌC FILE ERP (003.xlsx)
    # ---------------------------------------------------------
    f_erp.seek(0)
    wb_erp = openpyxl.load_workbook(f_erp, data_only=True)
    ws_erp = wb_erp.active

    # Lấy ngày tồn kho từ dòng 8 (ví dụ: "Ngày tồn kho: 30-09-2026")
    ngay_chot = None
    for r in range(1, 12):
        for c in range(1, 10):
            v = str(ws_erp.cell(r, c).value or '')
            m = re.search(r'(\d{2}-\d{2}-\d{4})', v)
            if m:
                ngay_chot = datetime.datetime.strptime(m.group(1), '%d-%m-%Y').date()
                break
        if ngay_chot:
            break

    erp_headers_r11 = [ws_erp.cell(11, c).value for c in range(1, 47)]
    erp_headers_r12 = [ws_erp.cell(12, c).value for c in range(1, 47)]
    erp_meta_rows = []
    for r in range(1, 11):
        erp_meta_rows.append([ws_erp.cell(r, c).value for c in range(1, 47)])

    erp_data_rows = []
    for r in range(13, ws_erp.max_row + 1):
        kho_vt = ws_erp.cell(r, 1).value
        ma_hang = ws_erp.cell(r, 3).value
        if (kho_vt is None or str(kho_vt).strip() == '') and (ma_hang is None or str(ma_hang).strip() == ''):
            continue
        row_vals = [ws_erp.cell(r, c).value for c in range(1, 47)]
        erp_data_rows.append(row_vals)

    # ---------------------------------------------------------
    # 2. ĐỌC FILE PHÒNG CUNG ỨNG (CHỐT PHẠT TK...)
    # ---------------------------------------------------------
    f_pcu.seek(0)
    wb_pcu = openpyxl.load_workbook(f_pcu, data_only=True)

    ws_pcu_data = wb_pcu['Data']
    if ngay_chot is None:
        ngay_chot = parse_erp_date(ws_pcu_data['AT1'].value) or datetime.date.today()

    pcu_headers_r2 = [ws_pcu_data.cell(2, c).value for c in range(2, 50)]
    pcu_headers_r3 = [ws_pcu_data.cell(3, c).value for c in range(2, 50)]

    pcu_data_rows = []
    dict_pcu_lookup = {}
    for r in range(4, ws_pcu_data.max_row + 1):
        ma_hang = ws_pcu_data.cell(r, 2).value
        ma_lo = ws_pcu_data.cell(r, 17).value
        if ma_hang is None and ma_lo is None:
            continue
        row_vals = [ws_pcu_data.cell(r, c).value for c in range(2, 50)]
        pcu_data_rows.append(row_vals)

        phan_loai = clean_str_code(ws_pcu_data.cell(r, 5).value)
        kl_net = to_float(ws_pcu_data.cell(r, 20).value, 0.0)
        net_str = str(int(kl_net)) if kl_net.is_integer() else str(kl_net)
        key_chuoi = f"{clean_str_code(ma_lo)}-{net_str}-{phan_loai}"

        dict_pcu_lookup[key_chuoi] = {
            'dvkd': str(ws_pcu_data.cell(r, 42).value or '').strip(),
            'tinh_tk': str(ws_pcu_data.cell(r, 44).value or '').strip(),
            'tuoi': str(ws_pcu_data.cell(r, 47).value or '').strip(),
            'clc': str(ws_pcu_data.cell(r, 48).value or '').strip(),
        }

    # Đọc bảng tổng hợp từ sheet 'Tồn kho' của P.Cung ứng (Cột B..K = 2..11)
    ws_pcu_tk = wb_pcu['Tồn kho']
    pcu_tk_table1 = []  # Rows 12..17 (B..K)
    for r in range(12, 18):
        pcu_tk_table1.append([ws_pcu_tk.cell(r, c).value for c in range(2, 12)])

    pcu_tk_table2 = []  # Rows 23..28 (B..K)
    for r in range(23, 29):
        pcu_tk_table2.append([ws_pcu_tk.cell(r, c).value for c in range(2, 12)])

    # Bảng III CLC trong sheet 'Tồn kho' của P.CƯ nằm ở D33 (M1), E33 (M2), F33 (Tổng)
    pcu_tk_clc = [
        ws_pcu_tk.cell(33, 4).value,  # M1
        ws_pcu_tk.cell(33, 5).value,  # M2
        ws_pcu_tk.cell(33, 6).value,  # Tổng
    ]

    # ---------------------------------------------------------
    # 3. KHỞI TẠO WORKBOOK ĐẦU RA (CHỈ GỒM 4 SHEET CHÍNH)
    # ---------------------------------------------------------
    wb_out = openpyxl.Workbook()
    ws_ttr = wb_out.active
    ws_ttr.title = 'TTR'
    ws_dc = wb_out.create_sheet('ĐỐI CHIẾU')
    ws_ksnb = wb_out.create_sheet('KSNB')
    ws_pcu_out = wb_out.create_sheet('P.Cung ứng')

    # Fonts, Fills & Borders chuẩn
    f_title = Font(name='Times New Roman', size=14, bold=True, color='000000')
    f_sec = Font(name='Times New Roman', size=12, bold=True, color='000000')
    f_hdr = Font(name='Times New Roman', size=11, bold=True, color='000000')
    f_data = Font(name='Times New Roman', size=11, bold=False, color='000000')
    f_bold = Font(name='Times New Roman', size=11, bold=True, color='000000')
    f_red = Font(name='Times New Roman', size=11, bold=True, color='9C0006')

    # 3 tông màu phân biệt 3 khu vực: P.Cung ứng (Xanh dương) | KSNB (Xanh lá) | Chênh lệch (Vàng cam)
    fill_pcu_hdr = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    fill_ksnb_green = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    fill_diff_orange = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    fill_sub = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    fill_red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

    thin_s = Side(style='thin', color='000000')
    med_s = Side(style='medium', color='000000')
    border_all = Border(left=thin_s, right=thin_s, top=thin_s, bottom=thin_s)
    border_sep_left = Border(left=med_s, right=thin_s, top=thin_s, bottom=thin_s)

    al_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    al_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    al_right = Alignment(horizontal='right', vertical='center')

    # =========================================================
    # A. GHI SHEET 'P.Cung ứng'
    # =========================================================
    ws_pcu_out['U1'] = f"=SUBTOTAL(9,$U$4:$U${len(pcu_data_rows) + 3})"
    ws_pcu_out['AU1'] = datetime.datetime(ngay_chot.year, ngay_chot.month, ngay_chot.day)
    ws_pcu_out['AU1'].number_format = 'dd/mm/yyyy'

    for idx, h_val in enumerate(pcu_headers_r2, start=3):
        c = ws_pcu_out.cell(2, idx, h_val)
        c.font = f_hdr
        c.fill = fill_pcu_hdr
        c.alignment = al_center
        c.border = border_all
    for idx, h_val in enumerate(pcu_headers_r3, start=3):
        c = ws_pcu_out.cell(3, idx, h_val)
        c.font = f_hdr
        c.fill = fill_pcu_hdr
        c.alignment = al_center
        c.border = border_all

    max_row_ksnb = len(erp_data_rows) + 13
    for idx, r_vals in enumerate(pcu_data_rows, start=4):
        c_a = ws_pcu_out.cell(idx, 1, f"=COUNTIFS(KSNB!$J$14:$J${max_row_ksnb},'P.Cung ứng'!B{idx})")
        c_b = ws_pcu_out.cell(idx, 2, f'=R{idx}&"-"&U{idx}&"-"&F{idx}')
        c_a.font = f_data
        c_b.font = f_data
        for c_offset, val in enumerate(r_vals, start=3):
            cell = ws_pcu_out.cell(idx, c_offset, val)
            cell.font = f_data

    # =========================================================
    # B. GHI SHEET 'KSNB' (KHÔNG PHỤ THUỘC SHEET PHỤ)
    # =========================================================
    max_row_pcu = len(pcu_data_rows) + 3

    ws_ksnb['I3'] = 'Đối chiếu KL & Số lượng'
    ws_ksnb['K3'] = 'Số lượng (cuộn)'
    ws_ksnb['M3'] = 'Khối lượng kg'
    ws_ksnb['K4'] = 'NM'
    ws_ksnb['L4'] = 'P.Cung ứng'
    ws_ksnb['M4'] = 'NM'
    ws_ksnb['N4'] = 'P.Cung ứng'
    ws_ksnb['I6'] = 'Chênh lệch'

    ws_ksnb['K5'] = f'=COUNTIFS($Q$14:$Q${max_row_ksnb},"*?*")'
    ws_ksnb['L5'] = f'=COUNTIFS(\'P.Cung ứng\'!$C$4:$C${max_row_pcu},"<>")'
    ws_ksnb['M5'] = f'=SUM($AK$14:$AK${max_row_ksnb})'
    ws_ksnb['N5'] = f"=SUM('P.Cung ứng'!$U$4:$U${max_row_pcu})"
    ws_ksnb['K6'] = '=K5-L5'
    ws_ksnb['M6'] = '=M5-N5'

    for r_i in range(3, 7):
        for c_i in range(9, 15):
            cell = ws_ksnb.cell(r_i, c_i)
            cell.border = border_all
            cell.font = f_hdr if r_i in (3, 4, 6) else f_data
            if r_i == 5 and c_i in (11, 12, 13, 14):
                cell.number_format = '#,##0'

    for r_idx, m_row in enumerate(erp_meta_rows, start=1):
        for c_offset, val in enumerate(m_row, start=17):
            ws_ksnb.cell(r_idx, c_offset, val).font = f_data

    ws_ksnb['P12'] = datetime.datetime(ngay_chot.year, ngay_chot.month, ngay_chot.day)
    ws_ksnb['P12'].number_format = 'dd/mm/yyyy'
    ws_ksnb['P12'].font = f_hdr
    ws_ksnb['P12'].fill = fill_diff_orange
    ws_ksnb['P12'].alignment = al_center

    ksnb_calc_headers = [
        'Đối chiếu tuổi tồn kho',       # A (1)
        'P.CƯ tính tuổi tồn kho',       # B (2)
        'Đối chiếu tính phạt CLC',      # C (3)
        'P. CƯ tính phạt CLC',          # D (4)
        'Đối chiếu đánh tên',           # E (5)
        'P.CƯ đánh tên',                # F (6)
        'Đối chiếu đánh loại',          # G (7)
        'P.CƯ đánh loại',               # H (8)
        'Đối chiếu cuộn',               # I (9)
        'Chuỗi',                        # J (10)
        'Mức tính phạt CLC',            # K (11)
        'Đánh tên đơn vị kinh doanh',   # L (12)
        'Đánh loại',                    # M (13)
        'Tuổi tồn kho',                 # N (14)
        'Số ngày tồn kho',              # O (15)
        'Đánh số phân loại',            # P (16)
    ]
    for c_i, h_title in enumerate(ksnb_calc_headers, start=1):
        c = ws_ksnb.cell(13, c_i, h_title)
        c.font = f_hdr
        c.fill = fill_diff_orange
        c.alignment = al_center
        c.border = border_all

    for c_offset, val in enumerate(erp_headers_r11, start=17):
        c = ws_ksnb.cell(12, c_offset, val)
        c.font = f_hdr
        c.fill = fill_pcu_hdr
        c.alignment = al_center
        c.border = border_all
    for c_offset, val in enumerate(erp_headers_r12, start=17):
        c = ws_ksnb.cell(13, c_offset, val)
        c.font = f_hdr
        c.fill = fill_pcu_hdr
        c.alignment = al_center
        c.border = border_all

    discrepancy_records = []
    for idx, r_vals in enumerate(erp_data_rows, start=14):
        for c_offset, val in enumerate(r_vals, start=17):
            c = ws_ksnb.cell(idx, c_offset, val)
            c.font = f_data
            c.border = border_all

        # Công thức tự thân 100% (không cần Sheet2 hay NM Khác)
        formulas_a_to_p = {
            1: f'=TEXT(B{idx},0)=TEXT(N{idx},0)',
            2: f"=VLOOKUP(J{idx},'P.Cung ứng'!$B$1:$AX${max_row_pcu},47,0)",
            3: f'=TEXT(D{idx},0)=TEXT(K{idx},0)',
            4: f"=VLOOKUP(J{idx},'P.Cung ứng'!$B$1:$AX${max_row_pcu},48,0)",
            5: f'=TEXT(F{idx},0)=TEXT(L{idx},0)',
            6: f"=VLOOKUP(J{idx},'P.Cung ứng'!$B$1:$AX${max_row_pcu},42,0)",
            7: f'=TEXT(H{idx},0)=TEXT(M{idx},0)',
            8: f"=VLOOKUP(J{idx},'P.Cung ứng'!$B$1:$AX${max_row_pcu},45,0)",
            9: f"=COUNTIFS('P.Cung ứng'!$B$1:$B${max_row_pcu},$J{idx})",
            10: f'=AH{idx}&"-"&AK{idx}&"-"&V{idx}',
            11: f'=IF(Q{idx}="","",IF(OR(L{idx}="NM Lợp xưởng",O{idx}<91,AND(ISNUMBER(SEARCH("HOA SEN GOLD",T{idx})),O{idx}<180)),"LOẠI",IF(O{idx}<181,"M1","M2")))',
            12: (
                f'=IF(Q{idx}="","",IF(ISNUMBER(SEARCH("BPDT",TEXT(AY{idx},0))),"DT_BPDT",'
                f'IF(OR(LEFT(TEXT(AY{idx},0),4)="DTHS",LEFT(TEXT(AY{idx},0),4)="DLHS",TEXT(AY{idx},0)="ĐTHS"),"DTHS",'
                f'IF(OR(LEFT(TEXT(AY{idx},0),2)="CN",LEFT(TEXT(AY{idx},0),4)="VPDH"),"HTPP",'
                f'IF(AND(TEXT(AX{idx},0)="VPDH_MN_PKDTT",OR(TEXT(AY{idx},0)="KDMN01",TEXT(AY{idx},0)="PKDTT",TEXT(AY{idx},0)="KDTT-MN",TEXT(AY{idx},0)="P.KDTT_MN")),"KDTT",'
                f'IF(AND(OR(TEXT(AX{idx},0)="VPDH_MN_NM_BANNGOAI",TEXT(AX{idx},0)="VPDH_MT_NM_BANNGOAI"),TEXT(AY{idx},0)<>""),"Khách hàng ngoài trúng thầu",'
                f'IF(OR(ISNUMBER(SEARCH("LỢP XƯỞNG",TEXT(AY{idx},0))),ISNUMBER(SEARCH("LOPXUONG",TEXT(AY{idx},0)))),"NM Lợp xưởng",'
                f'IF(AND(TEXT(AX{idx},0)="PKDDA_HSH",OR(TEXT(AY{idx},0)="PKDDA_HSH",TEXT(AY{idx},0)="PKD_HSHome")),"PKDDA_HSH",'
                f'IF(AND(OR(TEXT(AX{idx},0)="HSG_PXK",TEXT(AX{idx},0)="VPDH_MN_NM"),OR(TEXT(AY{idx},0)="XK.HSG_OU",TEXT(AY{idx},0)="XK-L2",TEXT(AY{idx},0)="XK.HSG",TEXT(AY{idx},0)="SMC-XK")),"XK",'
                f'IF(TEXT(AY{idx},0)="BAOBI","BTP- BAO BÌ",'
                f'IF(AND(OR(ISNUMBER(SEARCH("MM",TEXT(AY{idx},0))),ISNUMBER(SEARCH("XBPM",TEXT(AY{idx},0)))),ISERR(SEARCH("CXL",TEXT(AY{idx},0)))),"BTP- Mạ màu & Xẻ băng ",'
                f'IF(AND(ISNUMBER(SEARCH("VLXD",TEXT(AY{idx},0))),ISERR(SEARCH("CXL",TEXT(AY{idx},0)))),"VLXD",'
                f'IF(OR(TEXT(AY{idx},0)="HS-NHONHOI",TEXT(AY{idx},0)="HS-BINHDUONG",TEXT(AY{idx},0)="NHUNG-XB-XK",TEXT(AY{idx},0)="NHUNGNONG-XB",TEXT(AY{idx},0)="HS-BINH DUONG",TEXT(AY{idx},0)="HS-NHUNGNONG",TEXT(AY{idx},0)="HS-BINH DUONG ",TEXT(AY{idx},0)="HS-NHUA",TEXT(AY{idx},0)="HS-NGHEAN",TEXT(AY{idx},0)="NHUNG",TEXT(AY{idx},0)="NHUNGNONG",TEXT(AY{idx},0)="NHUNG-XK",TEXT(AY{idx},0)="NHONHOA-OT"),"Các NM Khác","NM THSPM")))))))))))))'
            ),
            13: (
                f'=IF(Q{idx}="","",IF(OR(ISNUMBER(SEARCH("BPDT",TEXT(AY{idx},0))),'
                f'OR(LEFT(TEXT(AY{idx},0),4)="DTHS",TEXT(AY{idx},0)="ĐTHS",LEFT(TEXT(AY{idx},0),4)="DLHS"),'
                f'OR(LEFT(TEXT(AY{idx},0),2)="CN",LEFT(TEXT(AY{idx},0),4)="VPDH"),'
                f'AND(TEXT(AX{idx},0)="VPDH_MN_PKDTT",OR(TEXT(AY{idx},0)="KDMN01",TEXT(AY{idx},0)="PKDTT",TEXT(AY{idx},0)="KDTT-MN",TEXT(AY{idx},0)="P.KDTT_MN")),'
                f'AND(TEXT(AX{idx},0)="VPDH_MN_NM_BANNGOAI",TEXT(AY{idx},0)<>""),'
                f'AND(AX{idx}="VPDH_MN_NM",OR(AY{idx}="NM-LỢP XƯỞNG",AY{idx}="NM-LOPXUONG")),'
                f'AND(TEXT(AX{idx},0)="PKDDA_HSH",OR(TEXT(AY{idx},0)="PKDDA_HSH",TEXT(AY{idx},0)="PKD_HSHome")),'
                f'AND(OR(TEXT(AX{idx},0)="HSG_PXK",TEXT(AX{idx},0)="VPDH_MN_NM"),OR(TEXT(AY{idx},0)="XK.HSG_OU",TEXT(AY{idx},0)="SMC-XK")),'
                f'OR(TEXT(BF{idx},0)="033-01",TEXT(BF{idx},0)="033-02")),"LOẠI","TÍNH"))'
            ),
            14: f'=IF(Q{idx}="","",IF(O{idx}="","",IF(O{idx}<91,"Dưới 90 ngày",IF(O{idx}<121,"Từ 91 đến 120 ngày",IF(O{idx}<181,"Từ 121 đến 180 ngày",IF(O{idx}<361,"Từ 181 đến 360 ngày","Trên 361 ngày"))))))',
            15: f'=$P$12-BC{idx}',
            16: f'=IF(Q{idx}="","",IF(OR(V{idx}="1",V{idx}="A2"),1,2))',
        }

        res_row = tinh_logic_ksnb_dong(
            ten_hang=r_vals[3],
            phan_loai=r_vals[5],
            ma_lo=r_vals[17],
            kl_net=to_float(r_vals[20], 0.0),
            dv_kd=r_vals[33],
            danh_ten=r_vals[34],
            ngay_nhap=r_vals[38],
            kho_cn=r_vals[41],
            ngay_chot=ngay_chot,
        )
        pcu_info = dict_pcu_lookup.get(res_row['chuoi'])

        err_notes = []
        check_flags = {1: True, 3: True, 5: True, 7: True}
        if pcu_info is None:
            err_notes.append('Không tìm thấy chuỗi cuộn bên P.Cung ứng')
            check_flags = {1: False, 3: False, 5: False, 7: False}
        else:
            if pcu_info['tuoi'].strip() != res_row['tuoi_tk'].strip():
                check_flags[1] = False
                err_notes.append(f"Tuổi TK: P.CƯ ({pcu_info['tuoi']}) ≠ KSNB ({res_row['tuoi_tk']})")
            if pcu_info['clc'].strip() != res_row['muc_clc'].strip():
                check_flags[3] = False
                err_notes.append(f"Tính phạt CLC: P.CƯ ({pcu_info['clc']}) ≠ KSNB ({res_row['muc_clc']})")
            if pcu_info['dvkd'].strip() != res_row['danh_ten_dv'].strip():
                check_flags[5] = False
                err_notes.append(f"Đánh tên ĐVKD: P.CƯ ({pcu_info['dvkd']}) ≠ KSNB ({res_row['danh_ten_dv']})")
            if pcu_info['tinh_tk'].strip() != res_row['danh_loai'].strip():
                check_flags[7] = False
                err_notes.append(f"Đánh loại: P.CƯ ({pcu_info['tinh_tk']}) ≠ KSNB ({res_row['danh_loai']})")

        for c_i, f_str in formulas_a_to_p.items():
            cell = ws_ksnb.cell(idx, c_i, f_str)
            cell.font = f_data
            cell.border = border_all
            if c_i in check_flags and not check_flags[c_i]:
                cell.fill = fill_red
                cell.font = f_red

        if err_notes:
            discrepancy_records.append({
                'Dòng Excel (KSNB)': idx,
                'Chuỗi cuộn': res_row['chuoi'],
                'Mã hàng': r_vals[2],
                'Khối lượng Net (kg)': to_float(r_vals[20], 0.0),
                'ĐVKD (ERP)': r_vals[33],
                'Đánh tên (ERP)': r_vals[34],
                'Nội dung lệch giữa P.CƯ & KSNB': ' | '.join(err_notes),
            })

    # =========================================================
    # C. GHI SHEET 'ĐỐI CHIẾU' (GỘP Ô & PHÂN 3 MÀU TÁCH BIỆT RÕ RÀNG)
    # =========================================================
    ws_dc['A3'] = 'BÁO CÁO TỒN KHO THỰC TẾ SO VỚI ĐỊNH MỨC BAN HÀNH'
    ws_dc['A3'].font = f_title
    ws_dc['A4'] = f"Tồn đến ngày {ngay_chot.strftime('%d/%m/%Y')}"
    ws_dc['A4'].font = f_sec

    ws_dc['A7'] = 'I. Tổng tồn kho Nhà máy'
    ws_dc['A7'].font = f_sec
    ws_dc['H7'] = 'ĐVT: KG'
    ws_dc['H7'].font = f_hdr

    age_cols_6 = [
        'Tổng cộng',
        'Dưới 90 ngày',
        'Từ 91 đến 120 ngày',
        'Từ 121 đến 180 ngày',
        'Từ 181 đến 360 ngày',
        'Trên 361 ngày',
    ]

    def setup_dc_table_headers(start_r):
        """Thiết lập tiêu đề 2 dòng có gộp ô (Merge) và tô 3 màu phân vùng: P.CƯ | KSNB | Chênh lệch."""
        ws_dc.cell(start_r, 1, 'Số TT')
        ws_dc.cell(start_r, 2, 'Hạng mục')
        ws_dc.cell(start_r, 3, 'Khối lượng định mức (DM)')
        ws_dc.cell(start_r, 4, 'KHỐI LƯỢNG THỰC TẾ THEO P.CUNG ỨNG (TT)')
        ws_dc.cell(start_r, 10, 'Chênh lệch P.CƯ\n(TT - DM)')
        ws_dc.cell(start_r, 11, 'Ghi chú')
        ws_dc.cell(start_r, 12, 'KHỐI LƯỢNG KIỂM TRA THEO KSNB')
        ws_dc.cell(start_r, 18, 'Chênh lệch KSNB\n(TT - DM)')
        ws_dc.cell(start_r, 19, 'CHÊNH LỆCH (KSNB - P.CUNG ỨNG)')

        for idx_a, lbl in enumerate(age_cols_6):
            ws_dc.cell(start_r + 1, 4 + idx_a, lbl)
            ws_dc.cell(start_r + 1, 12 + idx_a, lbl)
            ws_dc.cell(start_r + 1, 19 + idx_a, lbl)

        # Tô màu & kẻ khung cho cả 2 dòng tiêu đề trước khi Merge
        for r_h in (start_r, start_r + 1):
            for c_h in range(1, 25):
                c = ws_dc.cell(r_h, c_h)
                c.font = f_hdr
                c.alignment = al_center
                if c_h <= 11:
                    c.fill = fill_pcu_hdr
                elif c_h <= 18:
                    c.fill = fill_ksnb_green
                else:
                    c.fill = fill_diff_orange
                c.border = border_sep_left if c_h in (12, 19) else border_all

        # Gộp ô dọc cho các cột đơn và gộp ô ngang cho 3 cụm tiêu đề lớn
        for col_m in (1, 2, 3, 10, 11, 18):
            ws_dc.merge_cells(start_row=start_r, start_column=col_m, end_row=start_r + 1, end_column=col_m)
        ws_dc.merge_cells(start_row=start_r, start_column=4, end_row=start_r, end_column=9)
        ws_dc.merge_cells(start_row=start_r, start_column=12, end_row=start_r, end_column=17)
        ws_dc.merge_cells(start_row=start_r, start_column=19, end_row=start_r, end_column=24)

    # ---------------------------------------------------------
    # Bảng I: Tổng tồn kho Nhà máy (Header tại dòng 10-11, Dữ liệu dòng 12-17)
    # ---------------------------------------------------------
    setup_dc_table_headers(10)
    for idx, row_pcu in enumerate(pcu_tk_table1):
        r_excel = 12 + idx
        is_total = (idx == 0)
        ws_dc.cell(r_excel, 1, 'I.' if is_total else idx)
        ws_dc.cell(r_excel, 2, row_pcu[1])  # Cột C của P.CƯ: Hạng mục
        ws_dc.cell(r_excel, 3, row_pcu[2])  # Cột D của P.CƯ: Định mức

        if is_total:
            for c_i in range(4, 11):
                col_l = get_column_letter(c_i)
                ws_dc.cell(r_excel, c_i, f'=SUM({col_l}13:{col_l}17)')
            for c_i in range(12, 25):
                col_l = get_column_letter(c_i)
                ws_dc.cell(r_excel, c_i, f'=SUM({col_l}13:{col_l}17)')
        else:
            # Cột D..J (4..10) tương ứng index 3..9 trong row_pcu
            for c_off in range(3, 10):
                ws_dc.cell(r_excel, 1 + c_off, row_pcu[c_off])
            ws_dc.cell(r_excel, 12, f'=SUM(M{r_excel}:Q{r_excel})')
            for idx_age, col_ksnb_let in enumerate(['M', 'N', 'O', 'P', 'Q']):
                ws_dc.cell(
                    r_excel,
                    13 + idx_age,
                    f"=SUMIFS(KSNB!$AK$14:$AK${max_row_ksnb},KSNB!$L$14:$L${max_row_ksnb},'ĐỐI CHIẾU'!$B{r_excel},KSNB!$N$14:$N${max_row_ksnb},'ĐỐI CHIẾU'!{col_ksnb_let}$11)",
                )
            ws_dc.cell(r_excel, 18, f'=L{r_excel}-C{r_excel}')
            ws_dc.cell(r_excel, 19, f'=SUM(T{r_excel}:X{r_excel})')
            for idx_diff, (c_k, c_p) in enumerate([('M', 'E'), ('N', 'F'), ('O', 'G'), ('P', 'H'), ('Q', 'I')]):
                ws_dc.cell(r_excel, 20 + idx_diff, f'={c_k}{r_excel}-{c_p}{r_excel}')

        for c_i in range(1, 25):
            cell = ws_dc.cell(r_excel, c_i)
            cell.font = f_bold if is_total else f_data
            cell.border = border_sep_left if c_i in (12, 19) else border_all
            if is_total:
                cell.fill = fill_sub
            if c_i == 1:
                cell.alignment = al_center
            elif c_i == 2:
                cell.alignment = al_left
            elif c_i >= 3:
                cell.number_format = '#,##0'
                cell.alignment = al_right

    # ---------------------------------------------------------
    # Bảng II: Tồn kho hàng L2-A1, XKRL, CXL (Header dòng 21-22, Dữ liệu dòng 23-28)
    # ---------------------------------------------------------
    ws_dc['A19'] = 'II. Tồn kho hàng L2-A1, XKRL, CXL'
    ws_dc['A19'].font = f_sec
    setup_dc_table_headers(21)

    for idx, row_pcu in enumerate(pcu_tk_table2):
        r_excel = 23 + idx
        is_total = (idx == 0)
        ws_dc.cell(r_excel, 1, 'II.' if is_total else idx)
        ws_dc.cell(r_excel, 2, row_pcu[1])  # Hạng mục
        ws_dc.cell(r_excel, 3, row_pcu[2])  # Định mức

        if is_total:
            for c_i in range(4, 11):
                col_l = get_column_letter(c_i)
                ws_dc.cell(r_excel, c_i, f'=SUM({col_l}24:{col_l}28)')
            for c_i in range(12, 25):
                col_l = get_column_letter(c_i)
                ws_dc.cell(r_excel, c_i, f'=SUM({col_l}24:{col_l}28)')
        else:
            for c_off in range(3, 9):
                ws_dc.cell(r_excel, 1 + c_off, row_pcu[c_off])
            if idx == 1:
                ws_dc.cell(r_excel, 10, '=SUM(D24:D28)-C24')
            ws_dc.cell(r_excel, 12, f'=SUM(M{r_excel}:Q{r_excel})')
            for idx_age, col_ksnb_let in enumerate(['M', 'N', 'O', 'P', 'Q']):
                ws_dc.cell(
                    r_excel,
                    13 + idx_age,
                    f"=SUMIFS(KSNB!$AK$14:$AK${max_row_ksnb},KSNB!$L$14:$L${max_row_ksnb},'ĐỐI CHIẾU'!$B{r_excel},KSNB!$N$14:$N${max_row_ksnb},'ĐỐI CHIẾU'!{col_ksnb_let}$22,KSNB!$P$14:$P${max_row_ksnb},\"2\")",
                )
            ws_dc.cell(r_excel, 18, f'=L{r_excel}-C{r_excel}')
            ws_dc.cell(r_excel, 19, f'=SUM(T{r_excel}:X{r_excel})')
            for idx_diff, (c_k, c_p) in enumerate([('M', 'E'), ('N', 'F'), ('O', 'G'), ('P', 'H'), ('Q', 'I')]):
                ws_dc.cell(r_excel, 20 + idx_diff, f'={c_k}{r_excel}-{c_p}{r_excel}')

        for c_i in range(1, 25):
            cell = ws_dc.cell(r_excel, c_i)
            cell.font = f_bold if is_total else f_data
            cell.border = border_sep_left if c_i in (12, 19) else border_all
            if is_total:
                cell.fill = fill_sub
            if c_i == 1:
                cell.alignment = al_center
            elif c_i == 2:
                cell.alignment = al_left
            elif c_i >= 3:
                cell.number_format = '#,##0'
                cell.alignment = al_right

    # ---------------------------------------------------------
    # Bảng III: Tồn kho Chậm Luân Chuyển (CLC) (Dòng 30..33)
    # ---------------------------------------------------------
    ws_dc['A30'] = 'III. Tồn kho CLC'
    ws_dc['A30'].font = f_sec

    # Dòng 30: Tiêu đề phân vùng 3 bên cho bảng CLC
    ws_dc['B30'] = 'THEO P.CUNG ỨNG'
    ws_dc['F30'] = 'THEO KSNB'
    ws_dc['J30'] = 'CHÊNH LỆCH (P.CƯ - KSNB)'

    ws_dc['B31'] = 'Loại'
    ws_dc['C31'] = 'M1'
    ws_dc['D31'] = 'M2'
    ws_dc['E31'] = 'Tổng'
    ws_dc['F31'] = 'Loại'
    ws_dc['G31'] = 'M1'
    ws_dc['H31'] = 'M2'
    ws_dc['I31'] = 'Tổng'
    ws_dc['J31'] = 'M1'
    ws_dc['K31'] = 'M2'
    ws_dc['L31'] = 'Tổng'

    ws_dc['B32'] = 'Số ngày tồn kho'
    ws_dc['C32'] = '91-180'
    ws_dc['D32'] = '181- Trên 360'
    ws_dc['F32'] = 'Số ngày tồn kho'
    ws_dc['G32'] = 'Từ 91 đến 180 ngày'
    ws_dc['H32'] = '>181 ngày '
    ws_dc['J32'] = 'Từ 91 đến 180 ngày'
    ws_dc['K32'] = '>181 ngày '

    ws_dc['B33'] = 'Khối lượng (kg)'
    ws_dc['C33'] = pcu_tk_clc[0]
    ws_dc['D33'] = pcu_tk_clc[1]
    ws_dc['E33'] = pcu_tk_clc[2]
    ws_dc['F33'] = 'Khối lượng (kg)'
    ws_dc['G33'] = f"=SUMIFS(KSNB!$AK$14:$AK${max_row_ksnb},KSNB!$K$14:$K${max_row_ksnb},'ĐỐI CHIẾU'!G$31,KSNB!$M$14:$M${max_row_ksnb},\"TÍNH\")"
    ws_dc['H33'] = f"=SUMIFS(KSNB!$AK$14:$AK${max_row_ksnb},KSNB!$K$14:$K${max_row_ksnb},'ĐỐI CHIẾU'!H$31,KSNB!$M$14:$M${max_row_ksnb},\"TÍNH\")"
    ws_dc['I33'] = '=SUM(G33:H33)'
    ws_dc['J33'] = '=C33-G33'
    ws_dc['K33'] = '=D33-H33'
    ws_dc['L33'] = '=SUM(J33:K33)'

    for r_i in range(30, 34):
        for c_i in range(2, 13):
            c = ws_dc.cell(r_i, c_i)
            c.border = border_sep_left if c_i in (6, 10) else border_all
            c.font = f_hdr if r_i in (30, 31, 32) else f_bold
            if r_i in (30, 31, 32):
                c.alignment = al_center
                if c_i <= 5:
                    c.fill = fill_pcu_hdr
                elif c_i <= 9:
                    c.fill = fill_ksnb_green
                else:
                    c.fill = fill_diff_orange
            elif c_i in (3, 4, 5, 7, 8, 9, 10, 11, 12):
                c.number_format = '#,##0'
                c.alignment = al_right

    ws_dc.merge_cells('B30:E30')
    ws_dc.merge_cells('F30:I30')
    ws_dc.merge_cells('J30:L30')
    ws_dc.merge_cells('E31:E32')
    ws_dc.merge_cells('I31:I32')
    ws_dc.merge_cells('L31:L32')

    # =========================================================
    # D. GHI SHEET 'TTR' (TỜ TRÌNH TỔNG HỢP PHẠT)
    # =========================================================
    ws_ttr['C1'] = 'BÁO CÁO ĐỊNH MỨC TỒN KHO VÀ CHẬM LUÂN CHUYỂN'
    ws_ttr['C1'].font = f_title
    ws_ttr['L1'] = datetime.datetime(ngay_chot.year, ngay_chot.month, ngay_chot.day)
    ws_ttr['L1'].number_format = 'dd/mm/yyyy'
    ws_ttr['C2'] = '=+"Số liệu chốt hết "&TEXT(KSNB!P12,"dd/mm/yyyy")'
    ws_ttr['C2'].font = f_sec

    ws_ttr['A3'] = 'I.'
    ws_ttr['B3'] = 'Định Mức Tồn Kho'
    ws_ttr['A3'].font = f_sec
    ws_ttr['B3'].font = f_sec
    ws_ttr['H4'] = 'Đvt: kg'

    ttr_cols = [
        'Nội dung',
        'Tổng',
        'NM THSPM',
        'BTP- Mạ màu & Xẻ băng ',
        'BTP- BAO BÌ',
        'VLXD',
        'TXL-ĐT',
        'KHCU',
        'Các NM Khác',
    ]
    for idx, h_lbl in enumerate(ttr_cols, start=2):
        c = ws_ttr.cell(5, idx, h_lbl)
        c.font = f_hdr
        c.fill = fill_pcu_hdr
        c.alignment = al_center
        c.border = border_all

    ws_ttr['B6'] = 'Định mức theo TTr số 099/TTr/P.CƯ/2023'
    ws_ttr['C6'] = '=SUM(D6:J6)'
    dm_vals = [2000000, 3500000, 140000, 2000000, 0, 0, 0]
    for idx, v in enumerate(dm_vals, start=4):
        ws_ttr.cell(6, idx, v)

    ws_ttr['B7'] = '="Số liệu tồn kho "&TEXT(KSNB!P12,"dd/mm/yyyy")'
    ws_ttr['C7'] = '=SUM(D7:J7)'
    ws_ttr['D7'] = "=SUMIFS('ĐỐI CHIẾU'!$L$13:$L$17,'ĐỐI CHIẾU'!$B$13:$B$17,TTR!D$5)"
    ws_ttr['E7'] = "=SUMIFS('ĐỐI CHIẾU'!$L$13:$L$17,'ĐỐI CHIẾU'!$B$13:$B$17,TTR!E$5)"
    ws_ttr['F7'] = "=SUMIFS('ĐỐI CHIẾU'!$L$13:$L$17,'ĐỐI CHIẾU'!$B$13:$B$17,TTR!F$5)"
    ws_ttr['G7'] = "=SUMIFS('ĐỐI CHIẾU'!$L$13:$L$17,'ĐỐI CHIẾU'!$B$13:$B$17,TTR!G$5)"
    ws_ttr['H7'] = "=SUMIFS('ĐỐI CHIẾU'!$D$13:$D$17,'ĐỐI CHIẾU'!$B$13:$B$17,TTR!H$5)"
    ws_ttr['I7'] = "=SUMIFS('ĐỐI CHIẾU'!$D$13:$D$17,'ĐỐI CHIẾU'!$B$13:$B$17,TTR!I$5)"
    ws_ttr['J7'] = "=SUMIFS('ĐỐI CHIẾU'!$L$13:$L$17,'ĐỐI CHIẾU'!$B$13:$B$17,TTR!J$5)"

    ws_ttr['B8'] = 'Tờ trình xin giảm phạt tồn kho hàng VLXD, BTP'
    ws_ttr['B9'] = 'Chênh lệch'
    ws_ttr['C9'] = '=SUM(D9:J9)'
    for idx, col_l in enumerate(['D', 'E', 'F', 'G', 'H', 'I', 'J'], start=4):
        ws_ttr.cell(9, idx, f'={col_l}7-{col_l}6-{col_l}8')

    ws_ttr['B10'] = 'Đơn giá (đồng/kg)'
    ws_ttr['C10'] = 7
    ws_ttr['B11'] = 'Thành tiền (vnđ)'
    ws_ttr['C11'] = '=IF(C9<0,0,C9*C10)'
    ws_ttr['B12'] = 'Kết Luận:'
    ws_ttr['C12'] = '=+"Tổng tồn kho thực tế"& IF(C11>0," vượt "," không vượt ")& "tổng định mức ban hành nên mức phạt bằng: "&TEXT(C11,"#,##0")&" đồng."'

    for r_i in range(6, 12):
        for c_i in range(2, 11):
            c = ws_ttr.cell(r_i, c_i)
            c.border = border_all
            c.font = f_bold if r_i in (9, 11) else f_data
            if c_i >= 3:
                c.number_format = '#,##0'
                c.alignment = al_right

    # Mục II: Định Mức Tồn Kho Hàng Loại 2
    ws_ttr['A14'] = 'II.'
    ws_ttr['B14'] = 'Định Mức Tồn Kho Hàng Loại 2'
    ws_ttr['A14'].font = f_sec
    ws_ttr['B14'].font = f_sec

    ws_ttr['B16'] = 'Nội dung'
    ws_ttr['C16'] = 'Tổng'
    ws_ttr['B17'] = 'Định mức theo TT (kg)'
    ws_ttr['C17'] = 1800000
    ws_ttr['B18'] = '="Tồn kho thực tế "&TEXT(KSNB!P12,"dd/mm/yyyy")&" (kg)"'
    ws_ttr['C18'] = "=SUM('ĐỐI CHIẾU'!$L$24:$L$28)"
    ws_ttr['B19'] = 'Chênh lệch (kg)'
    ws_ttr['C19'] = '=C18-C17'
    ws_ttr['B20'] = 'Đơn giá (Đồng/kg)'
    ws_ttr['C20'] = 7
    ws_ttr['B21'] = 'Thành tiền (VNĐ)'
    ws_ttr['C21'] = '=IF(C19<0,0,C19*C20)'
    ws_ttr['B22'] = 'Kết Luận:'
    ws_ttr['C22'] = '=IF(C21>0,"Tổng tồn kho thực tế vượt tổng định mức ban hành nên mức phạt bằng: "&TEXT(C21,"#,##0")&" đồng.","Tổng tồn kho thực tế không vượt Tổng định mức ban hành nên mức phạt bằng 0")'
    ws_ttr['B23'] = 'Số tiền phạt'
    ws_ttr['D23'] = 'đồng'

    for r_i in range(16, 22):
        for c_i in (2, 3):
            c = ws_ttr.cell(r_i, c_i)
            c.border = border_all
            c.font = f_hdr if r_i in (16, 21) else f_data
            if r_i == 16:
                c.fill = fill_pcu_hdr
            if c_i == 3 and r_i >= 17:
                c.number_format = '#,##0'

    # Mục III: Tồn Kho Chậm Luân Chuyển
    ws_ttr['A25'] = 'III.'
    ws_ttr['B25'] = 'Tồn Kho Chậm Luân Chuyển'
    ws_ttr['A25'].font = f_sec
    ws_ttr['B25'].font = f_sec

    ws_ttr['B26'] = 'Loại'
    ws_ttr['C26'] = 'M1'
    ws_ttr['D26'] = 'M2'
    ws_ttr['E26'] = 'Tổng'
    ws_ttr['B27'] = 'Số ngày tồn kho'
    ws_ttr['C27'] = 'Từ 91-180 ngày '
    ws_ttr['D27'] = '>180 ngày '
    ws_ttr['B28'] = 'Đơn giá (Đồng/kg)'
    ws_ttr['C28'] = 5
    ws_ttr['D28'] = 10
    ws_ttr['B29'] = 'Khối lượng (kg)'
    ws_ttr['C29'] = "=SUMIFS('ĐỐI CHIẾU'!$G$33:$H$33,'ĐỐI CHIẾU'!$G$31:$H$31,TTR!C$26)"
    ws_ttr['D29'] = "=SUMIFS('ĐỐI CHIẾU'!$G$33:$H$33,'ĐỐI CHIẾU'!$G$31:$H$31,TTR!D$26)"
    ws_ttr['E29'] = '=SUM(C29:D29)'
    ws_ttr['B30'] = 'Thành tiền (VNĐ)'
    ws_ttr['C30'] = '=C29*C28'
    ws_ttr['D30'] = '=D29*D28'
    ws_ttr['E30'] = '=SUM(C30:D30)'
    ws_ttr['B31'] = 'Kết Luận:'
    ws_ttr['C31'] = '=IF(E30>0,"Tiền phạt chậm luân chuyển là:"&" "&TEXT(E30,"#,##0")&" đồng","Hàng tồn kho chậm luân chuyển nhỏ hơn 90 ngày nên mức phạt bằng 0")'

    for r_i in range(26, 31):
        for c_i in range(2, 6):
            c = ws_ttr.cell(r_i, c_i)
            c.border = border_all
            c.font = f_hdr if r_i in (26, 30) else f_data
            if r_i == 26:
                c.fill = fill_pcu_hdr
            if r_i >= 28 and c_i >= 3:
                c.number_format = '#,##0'

    # Mục IV: Tổng Tiền Phạt
    ws_ttr['A33'] = 'IV.'
    ws_ttr['B33'] = 'Tổng Tiền Phạt'
    ws_ttr['A33'].font = f_sec
    ws_ttr['B33'].font = f_sec

    ws_ttr['B34'] = '="Tổng tiền phạt tháng "&MONTH(L1)&" là:"'
    ws_ttr['C34'] = '=C11+C21+E30'
    ws_ttr['D34'] = 'đồng'
    ws_ttr['B35'] = 'Trong đó:'
    ws_ttr['B36'] = '- Mức phạt cho BGĐ (30%):'
    ws_ttr['C36'] = '=C34*30%'
    ws_ttr['D36'] = 'đồng'
    ws_ttr['B37'] = '- Mức phạt cho P.KHCU NM (70%):'
    ws_ttr['C37'] = '=C34*70%'
    ws_ttr['D37'] = 'đồng'

    for r_i in (34, 36, 37):
        ws_ttr.cell(r_i, 2).font = f_bold
        c_val = ws_ttr.cell(r_i, 3)
        c_val.font = f_bold
        c_val.number_format = '#,##0'

    # Căn chỉnh độ rộng cột
    for ws_tmp in (ws_ttr, ws_dc, ws_ksnb, ws_pcu_out):
        for col_idx in range(1, ws_tmp.max_column + 1):
            ws_tmp.column_dimensions[get_column_letter(col_idx)].width = 18
    ws_ttr.column_dimensions['B'].width = 42
    ws_dc.column_dimensions['A'].width = 8
    ws_dc.column_dimensions['B'].width = 28

    output = io.BytesIO()
    wb_out.save(output)

    # ---------------------------------------------------------
    # 4. TỔNG HỢP SỐ LIỆU HIỂN THỊ NHANH TRÊN WEB STREAMLIT
    # ---------------------------------------------------------
    kl_nm_total = sum(to_float(r[20], 0.0) for r in erp_data_rows)
    kl_pcu_total = sum(to_float(r[18], 0.0) for r in pcu_data_rows)

    tt_tong_nm = to_float(pcu_tk_table1[0][3], 0.0)
    dm_tong_nm = 7640000.0
    phat_dm = max(0.0, (tt_tong_nm - dm_tong_nm) * 7.0)

    tt_l2 = to_float(pcu_tk_table2[0][3], 0.0)
    dm_l2 = 1800000.0
    phat_l2 = max(0.0, (tt_l2 - dm_l2) * 7.0)

    m1_kl = to_float(pcu_tk_clc[0], 0.0)
    m2_kl = to_float(pcu_tk_clc[1], 0.0)
    phat_clc = m1_kl * 5.0 + m2_kl * 10.0

    tong_phat = phat_dm + phat_l2 + phat_clc

    summary_rows = [
        {
            'Hạng mục kiểm tra': 'I. Tổng tồn kho Nhà máy (Vượt ĐM)',
            'Định mức (kg)': f'{dm_tong_nm:,.0f}',
            'Thực tế chốt (kg)': f'{tt_tong_nm:,.2f}',
            'Chênh lệch / Tính phạt (kg)': f'{tt_tong_nm - dm_tong_nm:,.2f}',
            'Đơn giá (đ/kg)': '7',
            'Thành tiền phạt (VNĐ)': f'{phat_dm:,.0f}',
        },
        {
            'Hạng mục kiểm tra': 'II. Tồn kho Hàng Loại 2 (L2-A1, XKRL, CXL)',
            'Định mức (kg)': f'{dm_l2:,.0f}',
            'Thực tế chốt (kg)': f'{tt_l2:,.2f}',
            'Chênh lệch / Tính phạt (kg)': f'{tt_l2 - dm_l2:,.2f}',
            'Đơn giá (đ/kg)': '7',
            'Thành tiền phạt (VNĐ)': f'{phat_l2:,.0f}',
        },
        {
            'Hạng mục kiểm tra': 'III.1. Tồn kho Chậm luân chuyển M1 (91-180 ngày)',
            'Định mức (kg)': '-',
            'Thực tế chốt (kg)': f'{m1_kl:,.2f}',
            'Chênh lệch / Tính phạt (kg)': f'{m1_kl:,.2f}',
            'Đơn giá (đ/kg)': '5',
            'Thành tiền phạt (VNĐ)': f'{m1_kl * 5.0:,.0f}',
        },
        {
            'Hạng mục kiểm tra': 'III.2. Tồn kho Chậm luân chuyển M2 (>180 ngày)',
            'Định mức (kg)': '-',
            'Thực tế chốt (kg)': f'{m2_kl:,.2f}',
            'Chênh lệch / Tính phạt (kg)': f'{m2_kl:,.2f}',
            'Đơn giá (đ/kg)': '10',
            'Thành tiền phạt (VNĐ)': f'{m2_kl * 10.0:,.0f}',
        },
    ]

    return {
        'excel_bytes': output.getvalue(),
        'ngay_chot_str': ngay_chot.strftime('%d/%m/%Y'),
        'file_out_name': f"KIỂM TRA ĐỊNH MỨC TỒN KHO & CLC NM THSPM T{ngay_chot.strftime('%m.%Y')}.xlsx",
        'count_nm': len(erp_data_rows),
        'count_pcu': len(pcu_data_rows),
        'kl_nm': kl_nm_total,
        'kl_pcu': kl_pcu_total,
        'phat_dm': phat_dm,
        'phat_l2': phat_l2,
        'phat_clc': phat_clc,
        'tong_phat': tong_phat,
        'summary_rows': summary_rows,
        'discrepancy_records': discrepancy_records,
    }


# -------------------------------------------------------------
# CHỨC NĂNG CHÍNH (GIAO DIỆN UPLOAD & KẾT QUẢ)
# -------------------------------------------------------------
def kiem_tra_dinh_muc_ton_kho():
    st.caption('Kiểm soát nội bộ • Tự động đối chiếu dữ liệu ERP với P.Cung ứng và xuất báo cáo Tờ trình phạt tồn kho')
    st.markdown('### 📥 Nạp Dữ Liệu Nguồn')

    col_upload1, col_upload2 = st.columns(2, gap='large')

    with col_upload1:
        with st.container(border=True):
            st.markdown(
                """
<div class="upload-title">1️⃣ Dữ liệu Tồn kho theo lô (ERP)</div>
<div class="guide-box">
<div class="guide-header">📌 1. Báo cáo Tồn kho theo lô (cuộn) - File 003</div>
<div class="guide-text">
Xuất từ hệ thống ERP (Kho: 031-036-038)<br>
Từ loại hàng: 3.01.01 đến 4.08.02<br>
Lưu ý: Nạp trực tiếp file gốc xuất từ ERP (code tự động loại bỏ dòng trống cuối file)
</div>
</div>
""",
                unsafe_allow_html=True,
            )
            f_erp = st.file_uploader('1. File Tồn kho theo lô ERP (003.xlsx)', type=['xlsx', 'xls'])

    with col_upload2:
        with st.container(border=True):
            st.markdown(
                """
<div class="upload-title">2️⃣ Dữ liệu Chốt phạt của P.Cung ứng</div>
<div class="guide-box">
<div class="guide-header">📌 2. File Chốt phạt tồn kho (P.Cung ứng gửi)</div>
<div class="guide-text">
File tổng hợp của Phòng Cung ứng (ví dụ: CHỐT PHẠT TK T09.2026.xlsx)<br>
Bao gồm đầy đủ các sheet: <b>Tồn kho</b>, <b>Data</b>, <b>Đánh tên</b>, <b>Mã hh</b>
</div>
</div>
""",
                unsafe_allow_html=True,
            )
            f_pcu = st.file_uploader('2. File Chốt phạt tồn kho P.Cung ứng', type=['xlsx', 'xls'])

    st.write('')
    if f_erp is not None and f_pcu is not None:
        if st.button('🚀 Bắt đầu kiểm tra & xuất báo cáo tự động', use_container_width=True):
            with st.spinner('Đang đối chiếu dữ liệu từng cuộn và lập bảng tính phạt...'):
                try:
                    res = tao_file_kiem_tra_dinh_muc(f_erp, f_pcu)

                    st.success(
                        f"✅ Đã đối soát hoàn tất số liệu chốt ngày **{res['ngay_chot_str']}**! "
                        f"File xuất ra gồm đúng 4 sheet chính (`TTR`, `ĐỐI CHIẾU`, `KSNB`, `P.Cung ứng`) và đã phân vùng màu rõ ràng."
                    )
                    st.download_button(
                        f"📥 Tải File Kết Quả ({res['file_out_name']})",
                        data=res['excel_bytes'],
                        file_name=res['file_out_name'],
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True,
                    )

                    # ---------------------------------------------------------
                    # HIỂN THỊ BÁO CÁO NHANH KẾT QUẢ KIỂM SOÁT TRÊN WEB
                    # ---------------------------------------------------------
                    st.markdown('### 📊 Báo Cáo Nhanh Kết Quả Kiểm Soát Tồn Kho & CLC')
                    with st.container(border=True):
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric(
                            '📦 Số cuộn đối chiếu (NM / P.CƯ)',
                            f"{res['count_nm']:,} / {res['count_pcu']:,}",
                            delta=f"Lệch {res['count_nm'] - res['count_pcu']} cuộn",
                            delta_color='off' if res['count_nm'] == res['count_pcu'] else 'inverse',
                        )
                        c2.metric(
                            '⚖️ Tổng KL đối chiếu (kg)',
                            f"{res['kl_nm']:,.2f}",
                            delta=f"Lệch {res['kl_nm'] - res['kl_pcu']:,.2f} kg",
                            delta_color='off' if round(res['kl_nm'], 2) == round(res['kl_pcu'], 2) else 'inverse',
                        )
                        c3.metric(
                            '💰 Tổng tiền phạt tháng',
                            f"{res['tong_phat']:,.0f} đ",
                            delta=f"BGĐ (30%): {res['tong_phat']*0.3:,.0f} đ",
                            delta_color='off',
                        )
                        c4.metric(
                            '🔍 Cuộn lệch giữa P.CƯ & KSNB',
                            f"{len(res['discrepancy_records'])} cuộn",
                            delta='Khớp 100%' if len(res['discrepancy_records']) == 0 else 'Cần rà soát',
                            delta_color='normal' if len(res['discrepancy_records']) == 0 else 'inverse',
                        )

                        st.markdown('#### 📋 Bảng Tổng Hợp Tính Phạt (Tờ Trình)')
                        st.dataframe(
                            pd.DataFrame(res['summary_rows']),
                            use_container_width=True,
                            hide_index=True,
                        )

                        if res['discrepancy_records']:
                            st.warning(
                                f"🚨 **Phát hiện {len(res['discrepancy_records'])} cuộn có chênh lệch cách phân loại giữa P.Cung ứng và công thức KSNB** "
                                f"(Đã tô đỏ tại các cột đối chiếu trong sheet `KSNB`):"
                            )
                            st.dataframe(
                                pd.DataFrame(res['discrepancy_records']),
                                use_container_width=True,
                                hide_index=True,
                            )
                        else:
                            st.success('🎉 Toàn bộ các cuộn đều khớp 100% giữa P.Cung ứng và KSNB!')

                except Exception as e:
                    st.error(f'Lỗi xử lý dữ liệu: {e}')
    else:
        st.info('👉 Vui lòng nạp đủ cả 2 file (File Tồn kho ERP `003` và File Chốt phạt của P.Cung ứng) để bắt đầu kiểm tra.')


kiem_tra_dinh_muc_ton_kho()
