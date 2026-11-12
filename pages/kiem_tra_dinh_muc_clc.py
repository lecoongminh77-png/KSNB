import datetime
import io
import re
import time
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN CHUẨN APPLE & THANH TIỆN ÍCH CỐ ĐỊNH
# -------------------------------------------------------------
st.set_page_config(
    page_title='Tính Phạt Tồn Kho Thành Phẩm Bán Thành Phẩm',
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
        padding: 15px 0 10px 0;
        margin-bottom: 5px;
    }
    .apple-main-title {
        font-size: 38px;
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
        transition: all 0.2s ease !important;
    }
    div.stButton > button:first-child:hover {
        background: #0077ed !important;
        color: #ffffff !important;
        transform: translateY(-1px);
    }

    /* Nút Download nổi bật */
    div.stDownloadButton > button:first-child {
        background: #1e7e34 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding: 12px 24px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        box-shadow: 0 6px 16px rgba(30, 126, 52, 0.25) !important;
    }
    div.stDownloadButton > button:first-child:hover {
        background: #218838 !important;
        color: #ffffff !important;
    }

    /* Container Streamlit Thẻ Apple */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff;
        border-radius: 18px;
        padding: 18px;
        border: 1px solid rgba(0, 0, 0, 0.07);
        box-shadow: 0 4px 18px rgba(0, 0, 0, 0.03);
    }

    /* Thẻ KPI Bento Apple */
    .bento-card {
        background: #ffffff;
        border-radius: 16px;
        padding: 18px 20px;
        border: 1px solid rgba(0, 0, 0, 0.06);
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.02);
        height: 100%;
    }
    .bento-card-highlight {
        background: linear-gradient(135deg, #0071e3 0%, #005bb5 100%);
        color: #ffffff !important;
        border-radius: 16px;
        padding: 18px 20px;
        box-shadow: 0 8px 24px rgba(0, 113, 227, 0.22);
        height: 100%;
    }
    .bento-label {
        font-size: 13px;
        font-weight: 600;
        color: #86868b;
        margin-bottom: 6px;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }
    .bento-label-white {
        font-size: 13px;
        font-weight: 600;
        color: rgba(255, 255, 255, 0.85);
        margin-bottom: 6px;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }
    .bento-value {
        font-size: 26px;
        font-weight: 700;
        color: #1d1d1f;
        margin-bottom: 6px;
    }
    .bento-value-white {
        font-size: 28px;
        font-weight: 800;
        color: #ffffff;
        margin-bottom: 6px;
    }
    .bento-sub {
        font-size: 13px;
        font-weight: 500;
        color: #515154;
    }
    .bento-sub-white {
        font-size: 13px;
        font-weight: 500;
        color: rgba(255, 255, 255, 0.92);
    }

    /* Thành phần chữ trong thẻ Upload */
    .upload-title {
        font-size: 18px;
        font-weight: 700;
        color: #1d1d1f;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .guide-box {
        background: #fbfbfd;
        border-radius: 12px;
        padding: 14px 16px;
        margin-bottom: 14px;
        border: 1px solid rgba(0,0,0,0.05);
    }
    .guide-header {
        font-weight: 700;
        color: #0071e3;
        font-size: 14px;
        margin-bottom: 6px;
    }
    .guide-text {
        font-size: 13px;
        color: #515154;
        line-height: 1.55;
        margin: 0;
    }

    /* Tuỳ chỉnh Thanh Tabs kiểu Segmented Control */
    button[data-baseweb="tab"] {
        font-size: 15px !important;
        font-weight: 600 !important;
        padding: 10px 18px !important;
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
    "<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 8px 0 16px 0;'>",
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
    "<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>",
    unsafe_allow_html=True,
)


# -------------------------------------------------------------
# DANH MỤC CỐ ĐỊNH (ĐÃ NHÚNG TRỰC TIẾP VÀO CÔNG THỨC)
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
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace(',', '').strip())
    except Exception:
        return default


def clean_str_code(val):
    if val is None:
        return ''
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
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
    """Thực thi chính xác 100% logic công thức KSNB của bạn trên từng dòng dữ liệu."""
    net_num = float(kl_net)
    net_str = str(int(net_num)) if net_num.is_integer() else str(net_num)
    chuoi = f"{clean_str_code(ma_lo)}-{net_str}-{clean_str_code(phan_loai)}"

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

    # 4. Đánh tên đơn vị kinh doanh (Cột L)
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

    # 5. Đánh loại TÍNH / LOẠI (Cột M)
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

    # 6. Mức tính phạt CLC (Cột K)
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


# -------------------------------------------------------------
# BƯỚC 1: ĐỌC & XỬ LÝ SIÊU TỐC VÀO BỘ NHỚ ĐỆM (CÓ THANH % TIẾN TRÌNH)
# -------------------------------------------------------------
def doc_va_phan_tich_du_lieu(f_erp, f_pcu, progress_bar):
    t_start = time.time()

    # 1. Đọc siêu tốc File ERP (003.xlsx)
    progress_bar.progress(10, text='📥 [10%] Đang đọc dữ liệu Tồn kho theo lô từ ERP (File 003)...')
    f_erp.seek(0)
    wb_erp = openpyxl.load_workbook(f_erp, data_only=True, read_only=True)
    ws_erp = wb_erp.active

    all_erp_rows = list(ws_erp.iter_rows(values_only=True))
    wb_erp.close()

    ngay_chot = None
    for r_i in range(min(12, len(all_erp_rows))):
        row_v = all_erp_rows[r_i]
        for c_i in range(min(10, len(row_v))):
            v = str(row_v[c_i] or '')
            m = re.search(r'(\d{2}-\d{2}-\d{4})', v)
            if m:
                ngay_chot = datetime.datetime.strptime(m.group(1), '%d-%m-%Y').date()
                break
        if ngay_chot:
            break

    def pad_row(r_tuple, length=46):
        lst = list(r_tuple[:length]) if r_tuple else []
        if len(lst) < length:
            lst.extend([None] * (length - len(lst)))
        return lst

    erp_meta_rows = [pad_row(all_erp_rows[i], 46) for i in range(min(10, len(all_erp_rows)))]
    erp_headers_r11 = pad_row(all_erp_rows[10], 46) if len(all_erp_rows) > 10 else [None] * 46
    erp_headers_r12 = pad_row(all_erp_rows[11], 46) if len(all_erp_rows) > 11 else [None] * 46

    erp_data_rows = []
    for r_tuple in all_erp_rows[12:]:
        if not r_tuple:
            continue
        kho_vt = r_tuple[0] if len(r_tuple) > 0 else None
        ma_hang = r_tuple[2] if len(r_tuple) > 2 else None
        if (kho_vt is None or str(kho_vt).strip() == '') and (ma_hang is None or str(ma_hang).strip() == ''):
            continue
        erp_data_rows.append(pad_row(r_tuple, 46))

    # 2. Đọc siêu tốc File Phòng Cung ứng
    progress_bar.progress(30, text='📥 [30%] Đang đọc dữ liệu Chốt phạt của Phòng Cung ứng...')
    f_pcu.seek(0)
    wb_pcu = openpyxl.load_workbook(f_pcu, data_only=True, read_only=True)
    ws_pcu_data = wb_pcu['Data']
    all_pcu_rows = list(ws_pcu_data.iter_rows(values_only=True))

    if ngay_chot is None and len(all_pcu_rows) > 0 and len(all_pcu_rows[0]) >= 46:
        ngay_chot = parse_erp_date(all_pcu_rows[0][45]) or datetime.date.today()
    elif ngay_chot is None:
        ngay_chot = datetime.date.today()

    def slice_pcu_row(r_tuple):
        lst = list(r_tuple[1:49]) if r_tuple and len(r_tuple) > 1 else []
        if len(lst) < 48:
            lst.extend([None] * (48 - len(lst)))
        return lst

    pcu_headers_r2 = slice_pcu_row(all_pcu_rows[1]) if len(all_pcu_rows) > 1 else [None] * 48
    pcu_headers_r3 = slice_pcu_row(all_pcu_rows[2]) if len(all_pcu_rows) > 2 else [None] * 48

    pcu_data_rows = []
    dict_pcu_lookup = {}
    for r_tuple in all_pcu_rows[3:]:
        if not r_tuple:
            continue
        ma_hang = r_tuple[1] if len(r_tuple) > 1 else None
        ma_lo = r_tuple[16] if len(r_tuple) > 16 else None
        if ma_hang is None and ma_lo is None:
            continue
        row_vals = slice_pcu_row(r_tuple)
        pcu_data_rows.append(row_vals)

        phan_loai = clean_str_code(row_vals[3])
        kl_net = to_float(row_vals[18], 0.0)
        net_str = str(int(kl_net)) if kl_net.is_integer() else str(kl_net)
        key_chuoi = f"{clean_str_code(ma_lo)}-{net_str}-{phan_loai}"

        dict_pcu_lookup[key_chuoi] = {
            'dvkd': str(row_vals[40] or '').strip(),
            'tinh_tk': str(row_vals[41] or '').strip(),  # Cột AQ - Tính TK
            'tuoi': str(row_vals[45] or '').strip(),
            'clc': str(row_vals[46] or '').strip(),
        }

    # Đọc sheet 'Tồn kho' của P.Cung ứng
    ws_pcu_tk = wb_pcu['Tồn kho']
    tk_rows = list(ws_pcu_tk.iter_rows(min_row=1, max_row=35, max_col=15, values_only=True))
    wb_pcu.close()

    pcu_tk_table1 = [list(tk_rows[r_i][1:11]) for r_i in range(11, 17)]
    pcu_tk_table2 = [list(tk_rows[r_i][1:11]) for r_i in range(22, 28)]
    pcu_tk_clc = [tk_rows[32][3], tk_rows[32][4], tk_rows[32][5]]

    # 3. Tính toán đối chiếu từng cuộn và cập nhật thanh %
    total_coils = len(erp_data_rows)
    web_all_records = []
    discrepancy_records = []
    ksnb_flags_list = []

    age_order = ['Dưới 90 ngày', 'Từ 91 đến 120 ngày', 'Từ 121 đến 180 ngày', 'Từ 181 đến 360 ngày', 'Trên 361 ngày']
    dvkd_order = ['NM THSPM', 'BTP- Mạ màu & Xẻ băng ', 'BTP- BAO BÌ', 'VLXD', 'Các NM Khác']
    ksnb_agg_t1 = {dv: {ag: 0.0 for ag in age_order} for dv in dvkd_order}
    ksnb_agg_t2 = {dv: {ag: 0.0 for ag in age_order} for dv in dvkd_order}
    ksnb_clc_m1 = 0.0
    ksnb_clc_m2 = 0.0

    step_update = max(1, total_coils // 5)
    for i, r_vals in enumerate(erp_data_rows):
        if i % step_update == 0:
            pct = 40 + int((i / max(1, total_coils)) * 45)
            progress_bar.progress(pct, text=f'🔍 [{pct}%] Đang đối chiếu chéo KSNB & P.CƯ: {i:,} / {total_coils:,} cuộn...')

        idx_excel = i + 14
        kl_net = to_float(r_vals[20], 0.0)
        res_row = tinh_logic_ksnb_dong(
            ten_hang=r_vals[3],
            phan_loai=r_vals[5],
            ma_lo=r_vals[17],
            kl_net=kl_net,
            dv_kd=r_vals[33],
            danh_ten=r_vals[34],
            ngay_nhap=r_vals[38],
            kho_cn=r_vals[41],
            ngay_chot=ngay_chot,
        )

        dv_key = res_row['danh_ten_dv']
        ag_key = res_row['tuoi_tk']
        if dv_key in ksnb_agg_t1 and ag_key in ksnb_agg_t1[dv_key]:
            ksnb_agg_t1[dv_key][ag_key] += kl_net
            if res_row['so_phan_loai'] == 2:
                ksnb_agg_t2[dv_key][ag_key] += kl_net

        if res_row['danh_loai'] == 'TÍNH':
            if res_row['muc_clc'] == 'M1':
                ksnb_clc_m1 += kl_net
            elif res_row['muc_clc'] == 'M2':
                ksnb_clc_m2 += kl_net

        pcu_info = dict_pcu_lookup.get(res_row['chuoi'])
        err_types = []
        err_notes = []
        check_flags = {1: True, 3: True, 5: True, 7: True}

        if pcu_info is None:
            err_types.append('Lệch Cuộn')
            err_notes.append('Không tìm thấy chuỗi cuộn bên P.Cung ứng')
            check_flags = {1: False, 3: False, 5: False, 7: False}
            pcu_tuoi, pcu_clc, pcu_dvkd, pcu_loai = '', '', '', ''
        else:
            pcu_tuoi = pcu_info['tuoi']
            pcu_clc = pcu_info['clc']
            pcu_dvkd = pcu_info['dvkd']
            pcu_loai = pcu_info['tinh_tk']

            if pcu_tuoi.strip() != res_row['tuoi_tk'].strip():
                check_flags[1] = False
                err_types.append('Lệch Tuổi TK')
                err_notes.append(f"Tuổi TK: P.CƯ ({pcu_tuoi}) ≠ KSNB ({res_row['tuoi_tk']})")
            if pcu_clc.strip() != res_row['muc_clc'].strip():
                check_flags[3] = False
                err_types.append('Lệch Phạt CLC')
                err_notes.append(f"Phạt CLC: P.CƯ ({pcu_clc}) ≠ KSNB ({res_row['muc_clc']})")
            if pcu_dvkd.strip() != res_row['danh_ten_dv'].strip():
                check_flags[5] = False
                err_types.append('Lệch Đánh tên ĐVKD')
                err_notes.append(f"Đánh tên: P.CƯ ({pcu_dvkd}) ≠ KSNB ({res_row['danh_ten_dv']})")
            if pcu_loai.strip() != res_row['danh_loai'].strip():
                check_flags[7] = False
                err_types.append('Lệch Đánh loại')
                err_notes.append(f"Đánh loại: P.CƯ ({pcu_loai}) ≠ KSNB ({res_row['danh_loai']})")

        ksnb_flags_list.append((res_row, pcu_info, check_flags))

        rec_full = {
            'Dòng Excel': idx_excel,
            'Kho': clean_str_code(r_vals[0]),
            'Chuỗi cuộn': res_row['chuoi'],
            'KL Net (kg)': kl_net,
            'ĐVKD (ERP)': str(r_vals[33] or ''),
            'Đánh tên (ERP)': str(r_vals[34] or ''),
            'KSNB: Đánh tên': res_row['danh_ten_dv'].strip(),
            'P.CƯ: Đánh tên': pcu_dvkd,
            'KSNB: Đánh loại': res_row['danh_loai'],
            'P.CƯ: Đánh loại': pcu_loai,
            'KSNB: Phạt CLC': res_row['muc_clc'],
            'P.CƯ: Phạt CLC': pcu_clc,
            'Trạng thái': '✅ Khớp 100%' if not err_notes else '❌ Có lệch',
            'Nhóm lỗi': ', '.join(err_types) if err_types else 'Khớp',
            'Chi tiết chênh lệch': ' | '.join(err_notes) if err_notes else '',
        }
        web_all_records.append(rec_full)

        if err_notes:
            discrepancy_records.append({
                'Đã rà soát': False,
                'Dòng Excel': idx_excel,
                'Kho': clean_str_code(r_vals[0]),
                'Chuỗi cuộn': res_row['chuoi'],
                'KL Net (kg)': kl_net,
                'ĐVKD (ERP)': str(r_vals[33] or ''),
                'Đánh tên (ERP)': str(r_vals[34] or ''),
                'KSNB: Đánh tên': res_row['danh_ten_dv'].strip(),
                'P.CƯ: Đánh tên': pcu_dvkd,
                'KSNB: Đánh loại': res_row['danh_loai'],
                'P.CƯ: Đánh loại': pcu_loai,
                'KSNB: Phạt CLC': res_row['muc_clc'],
                'P.CƯ: Phạt CLC': pcu_clc,
                'Nhóm lỗi': ', '.join(err_types),
                'Ghi chú KSNB': ' | '.join(err_notes),
            })

    progress_bar.progress(95, text='📊 [95%] Đang tổng hợp bảng ĐỐI CHIẾU & Tờ trình (TTR)...')

    elapsed = round(time.time() - t_start, 2)
    progress_bar.progress(100, text=f'✅ [100%] Hoàn tất đối chiếu {total_coils:,} cuộn trong {elapsed} giây!')

    return {
        'ngay_chot': ngay_chot,
        'ngay_chot_str': ngay_chot.strftime('%d/%m/%Y'),
        'file_out_name': f"KIỂM TRA ĐỊNH MỨC TỒN KHO & CLC NM THSPM T{ngay_chot.strftime('%m.%Y')}.xlsx",
        'erp_meta_rows': erp_meta_rows,
        'erp_headers_r11': erp_headers_r11,
        'erp_headers_r12': erp_headers_r12,
        'erp_data_rows': erp_data_rows,
        'pcu_headers_r2': pcu_headers_r2,
        'pcu_headers_r3': pcu_headers_r3,
        'pcu_data_rows': pcu_data_rows,
        'pcu_tk_table1': pcu_tk_table1,
        'pcu_tk_table2': pcu_tk_table2,
        'pcu_tk_clc': pcu_tk_clc,
        'ksnb_flags_list': ksnb_flags_list,
        'ksnb_agg_t1': ksnb_agg_t1,
        'ksnb_agg_t2': ksnb_agg_t2,
        'ksnb_clc_m1': ksnb_clc_m1,
        'ksnb_clc_m2': ksnb_clc_m2,
        'web_all_records': web_all_records,
        'discrepancy_records': discrepancy_records,
        'count_nm': len(erp_data_rows),
        'count_pcu': len(pcu_data_rows),
        'kl_nm': sum(to_float(r[20], 0.0) for r in erp_data_rows),
        'kl_pcu': sum(to_float(r[18], 0.0) for r in pcu_data_rows),
        'elapsed_sec': elapsed,
    }


# -------------------------------------------------------------
# BƯỚC 2: HÀM XUẤT FILE EXCEL SIÊU TỐC (GIỮ NGUYÊN 100% CÔNG THỨC)
# -------------------------------------------------------------
def xuat_workbook_excel(data_bundle, giam_phat_dict=None):
    if giam_phat_dict is None:
        giam_phat_dict = {}

    ngay_chot = data_bundle['ngay_chot']
    erp_meta_rows = data_bundle['erp_meta_rows']
    erp_headers_r11 = data_bundle['erp_headers_r11']
    erp_headers_r12 = data_bundle['erp_headers_r12']
    erp_data_rows = data_bundle['erp_data_rows']
    pcu_headers_r2 = data_bundle['pcu_headers_r2']
    pcu_headers_r3 = data_bundle['pcu_headers_r3']
    pcu_data_rows = data_bundle['pcu_data_rows']
    pcu_tk_table1 = data_bundle['pcu_tk_table1']
    pcu_tk_table2 = data_bundle['pcu_tk_table2']
    pcu_tk_clc = data_bundle['pcu_tk_clc']
    ksnb_flags_list = data_bundle['ksnb_flags_list']

    wb_out = openpyxl.Workbook()
    ws_ttr = wb_out.active
    ws_ttr.title = 'TTR'
    ws_dc = wb_out.create_sheet('ĐỐI CHIẾU')
    ws_ksnb = wb_out.create_sheet('KSNB')
    ws_pcu_out = wb_out.create_sheet('P.Cung ứng')

    f_title = Font(name='Times New Roman', size=14, bold=True, color='000000')
    f_sec = Font(name='Times New Roman', size=12, bold=True, color='000000')
    f_hdr = Font(name='Times New Roman', size=11, bold=True, color='000000')
    f_data = Font(name='Times New Roman', size=11, bold=False, color='000000')
    f_bold = Font(name='Times New Roman', size=11, bold=True, color='000000')
    f_red = Font(name='Times New Roman', size=11, bold=True, color='9C0006')

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

    max_row_ksnb = len(erp_data_rows) + 13
    max_row_pcu = len(pcu_data_rows) + 3

    # =========================================================
    # A. GHI SHEET 'P.Cung ứng' BẰNG APPEND()
    # =========================================================
    row1_pcu = [None] * 50
    row1_pcu[20] = f"=SUBTOTAL(9,$U$4:$U${max_row_pcu})"
    row1_pcu[46] = datetime.datetime(ngay_chot.year, ngay_chot.month, ngay_chot.day)
    ws_pcu_out.append(row1_pcu)
    ws_pcu_out.append([None, None] + pcu_headers_r2)
    ws_pcu_out.append([None, None] + pcu_headers_r3)

    ws_pcu_out['AU1'].number_format = 'dd/mm/yyyy'
    for r_h in (2, 3):
        for c_h in range(3, 51):
            c = ws_pcu_out.cell(r_h, c_h)
            c.font = f_hdr
            c.fill = fill_pcu_hdr
            c.alignment = al_center
            c.border = border_all

    for idx, r_vals in enumerate(pcu_data_rows, start=4):
        f_a = f"=COUNTIFS(KSNB!$J$14:$J${max_row_ksnb},'P.Cung ứng'!B{idx})"
        f_b = f'=R{idx}&"-"&U{idx}&"-"&F{idx}'
        ws_pcu_out.append([f_a, f_b] + r_vals)

    # =========================================================
    # B. GHI SHEET 'KSNB' BẰNG APPEND() + TÔ ĐỎ Ô LỖI
    # =========================================================
    for r_idx in range(1, 12):
        m_row = erp_meta_rows[r_idx - 1] if r_idx - 1 < len(erp_meta_rows) else [None] * 46
        ws_ksnb.append([None] * 16 + m_row)

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

    row12_ksnb = [None] * 15 + [datetime.datetime(ngay_chot.year, ngay_chot.month, ngay_chot.day)] + erp_headers_r11
    ws_ksnb.append(row12_ksnb)

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
    ws_ksnb.append(ksnb_calc_headers + erp_headers_r12)

    ws_ksnb['P12'].number_format = 'dd/mm/yyyy'
    ws_ksnb['P12'].font = f_hdr
    ws_ksnb['P12'].fill = fill_diff_orange
    ws_ksnb['P12'].alignment = al_center

    for c_i in range(1, 63):
        c12 = ws_ksnb.cell(12, c_i)
        c13 = ws_ksnb.cell(13, c_i)
        if c_i >= 17:
            c12.font = f_hdr
            c12.fill = fill_pcu_hdr
            c12.alignment = al_center
            c12.border = border_all
        c13.font = f_hdr
        c13.fill = fill_diff_orange if c_i <= 16 else fill_pcu_hdr
        c13.alignment = al_center
        c13.border = border_all

    cells_to_highlight_red = []
    for i, r_vals in enumerate(erp_data_rows):
        idx = i + 14
        _, _, check_flags = ksnb_flags_list[i]

        f_list_a_to_p = [
            f'=TEXT(B{idx},0)=TEXT(N{idx},0)',
            f"=VLOOKUP(J{idx},'P.Cung ứng'!$B$1:$AX${max_row_pcu},47,0)",
            f'=TEXT(D{idx},0)=TEXT(K{idx},0)',
            f"=VLOOKUP(J{idx},'P.Cung ứng'!$B$1:$AX${max_row_pcu},48,0)",
            f'=TEXT(F{idx},0)=TEXT(L{idx},0)',
            f"=VLOOKUP(J{idx},'P.Cung ứng'!$B$1:$AX${max_row_pcu},42,0)",
            f'=TEXT(H{idx},0)=TEXT(M{idx},0)',
            f"=VLOOKUP(J{idx},'P.Cung ứng'!$B$1:$AX${max_row_pcu},45,0)",
            f"=COUNTIFS('P.Cung ứng'!$B$1:$B${max_row_pcu},$J{idx})",
            f'=AH{idx}&"-"&AK{idx}&"-"&V{idx}',
            f'=IF(Q{idx}="","",IF(OR(L{idx}="NM Lợp xưởng",O{idx}<91,AND(ISNUMBER(SEARCH("HOA SEN GOLD",T{idx})),O{idx}<180)),"LOẠI",IF(O{idx}<181,"M1","M2")))',
            (
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
            (
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
            f'=IF(Q{idx}="","",IF(O{idx}="","",IF(O{idx}<91,"Dưới 90 ngày",IF(O{idx}<121,"Từ 91 đến 120 ngày",IF(O{idx}<181,"Từ 121 đến 180 ngày",IF(O{idx}<361,"Từ 181 đến 360 ngày","Trên 361 ngày"))))))',
            f'=$P$12-BC{idx}',
            f'=IF(Q{idx}="","",IF(OR(V{idx}="1",V{idx}="A2"),1,2))',
        ]
        ws_ksnb.append(f_list_a_to_p + r_vals)

        for c_flag, is_ok in check_flags.items():
            if not is_ok:
                cells_to_highlight_red.append((idx, c_flag))

    for r_err, c_err in cells_to_highlight_red:
        cell_err = ws_ksnb.cell(r_err, c_err)
        cell_err.fill = fill_red
        cell_err.font = f_red

    # =========================================================
    # C. GHI SHEET 'ĐỐI CHIẾU' (GỘP Ô & PHÂN 3 MÀU RÕ RÀNG)
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
        'Tổng cộng', 'Dưới 90 ngày', 'Từ 91 đến 120 ngày',
        'Từ 121 đến 180 ngày', 'Từ 181 đến 360 ngày', 'Trên 361 ngày'
    ]

    def setup_dc_table_headers(start_r):
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

        for col_m in (1, 2, 3, 10, 11, 18):
            ws_dc.merge_cells(start_row=start_r, start_column=col_m, end_row=start_r + 1, end_column=col_m)
        ws_dc.merge_cells(start_row=start_r, start_column=4, end_row=start_r, end_column=9)
        ws_dc.merge_cells(start_row=start_r, start_column=12, end_row=start_r, end_column=17)
        ws_dc.merge_cells(start_row=start_r, start_column=19, end_row=start_r, end_column=24)

    # Bảng I: Tổng tồn kho
    setup_dc_table_headers(10)
    for idx, row_pcu in enumerate(pcu_tk_table1):
        r_excel = 12 + idx
        is_total = (idx == 0)
        ws_dc.cell(r_excel, 1, 'I.' if is_total else idx)
        ws_dc.cell(r_excel, 2, row_pcu[1])
        ws_dc.cell(r_excel, 3, row_pcu[2])

        if is_total:
            for c_i in range(4, 11):
                col_l = get_column_letter(c_i)
                ws_dc.cell(r_excel, c_i, f'=SUM({col_l}13:{col_l}17)')
            for c_i in range(12, 25):
                col_l = get_column_letter(c_i)
                ws_dc.cell(r_excel, c_i, f'=SUM({col_l}13:{col_l}17)')
        else:
            for c_off in range(3, 10):
                ws_dc.cell(r_excel, 1 + c_off, row_pcu[c_off])
            ws_dc.cell(r_excel, 12, f'=SUM(M{r_excel}:Q{r_excel})')
            for idx_age, col_ksnb_let in enumerate(['M', 'N', 'O', 'P', 'Q']):
                ws_dc.cell(r_excel, 13 + idx_age, f"=SUMIFS(KSNB!$AK$14:$AK${max_row_ksnb},KSNB!$L$14:$L${max_row_ksnb},'ĐỐI CHIẾU'!$B{r_excel},KSNB!$N$14:$N${max_row_ksnb},'ĐỐI CHIẾU'!{col_ksnb_let}$11)")
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

    # Bảng II: Tồn kho Loại 2
    ws_dc['A19'] = 'II. Tồn kho hàng L2-A1, XKRL, CXL'
    ws_dc['A19'].font = f_sec
    setup_dc_table_headers(21)

    for idx, row_pcu in enumerate(pcu_tk_table2):
        r_excel = 23 + idx
        is_total = (idx == 0)
        ws_dc.cell(r_excel, 1, 'II.' if is_total else idx)
        ws_dc.cell(r_excel, 2, row_pcu[1])
        ws_dc.cell(r_excel, 3, row_pcu[2])

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
                ws_dc.cell(r_excel, 13 + idx_age, f"=SUMIFS(KSNB!$AK$14:$AK${max_row_ksnb},KSNB!$L$14:$L${max_row_ksnb},'ĐỐI CHIẾU'!$B{r_excel},KSNB!$N$14:$N${max_row_ksnb},'ĐỐI CHIẾU'!{col_ksnb_let}$22,KSNB!$P$14:$P${max_row_ksnb},\"2\")")
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

    # Bảng III: Tồn kho CLC
    ws_dc['A30'] = 'III. Tồn kho CLC'
    ws_dc['A30'].font = f_sec
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
    # D. GHI SHEET 'TTR' (ĐỒNG BỘ SỐ LIỆU XIN GIẢM PHẠT TỪ WEB)
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

    ttr_cols = ['Nội dung', 'Tổng', 'NM THSPM', 'BTP- Mạ màu & Xẻ băng ', 'BTP- BAO BÌ', 'VLXD', 'TXL-ĐT', 'KHCU', 'Các NM Khác']
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
    ws_ttr['C8'] = '=SUM(D8:J8)'
    giam_keys = ['NM THSPM', 'BTP- Mạ màu & Xẻ băng ', 'BTP- BAO BÌ', 'VLXD', 'TXL-ĐT', 'KHCU', 'Các NM Khác']
    for idx, gk in enumerate(giam_keys, start=4):
        val_gp = to_float(giam_phat_dict.get(gk, 0.0), 0.0)
        if val_gp != 0:
            ws_ttr.cell(8, idx, val_gp)

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

    for ws_tmp in (ws_ttr, ws_dc):
        for col_idx in range(1, 25):
            ws_tmp.column_dimensions[get_column_letter(col_idx)].width = 18
    for col_idx in range(1, 17):
        ws_ksnb.column_dimensions[get_column_letter(col_idx)].width = 20
    ws_ttr.column_dimensions['B'].width = 42
    ws_dc.column_dimensions['A'].width = 8
    ws_dc.column_dimensions['B'].width = 28

    output = io.BytesIO()
    wb_out.save(output)
    return output.getvalue()


# -------------------------------------------------------------
# GIAO DIỆN CHÍNH & BẢNG ĐIỀU KHIỂN TƯƠNG TÁC TRỰC TIẾP TRÊN WEB
# -------------------------------------------------------------
def kiem_tra_dinh_muc_ton_kho():
    st.caption('Kiểm soát nội bộ • 1 Ổ tải file duy nhất. Rà soát trực tiếp trên Web. Xuất báo cáo Excel siêu tốc.')

    with st.container(border=True):
        st.markdown(
            """
<div class="upload-title">📥 Nạp Dữ Liệu Nguồn (Kéo thả cùng lúc 2 file)</div>
<div class="guide-box">
<div class="guide-text">
👉 <b>Quét chọn 2 file rồi kéo thả vào ô bên dưới:</b><br>
1. <b>Báo cáo Tồn kho theo lô (File 003 ERP):</b> Kho: 031-036-038 • Loại hàng: 3.01.01 đến 4.08.02<br>
2. <b>File Chốt phạt của P.Cung ứng:</b> Chứa đủ sheet <i>Tồn kho, Data, Đánh tên, Mã hh</i><br>
<span style="color:#0071e3;">Hệ thống sẽ tự động nhận diện đâu là file ERP, đâu là file P.Cung ứng.</span>
</div>
</div>
""",
            unsafe_allow_html=True,
        )
        
        uploaded_files = st.file_uploader(
            'Chọn 2 file (003 ERP và P.Cung ứng)', 
            type=['xlsx', 'xls'], 
            accept_multiple_files=True,
            label_visibility="collapsed"
        )
        
        f_erp = None
        f_pcu = None
        
        if uploaded_files:
            if len(uploaded_files) == 2:
                for f in uploaded_files:
                    f.seek(0)
                    try:
                        xl = pd.ExcelFile(f)
                        sheets = xl.sheet_names
                        f.seek(0)
                        if "Tồn kho" in sheets and "Data" in sheets:
                            f_pcu = f
                        else:
                            f_erp = f
                    except Exception as e:
                        st.error(f"Lỗi đọc file {f.name}: {e}")
            elif len(uploaded_files) > 2:
                st.warning("⚠️ Vui lòng chỉ kéo thả đúng 2 file (1 file ERP 003 và 1 file Chốt phạt P.CƯ).")
            elif len(uploaded_files) == 1:
                st.info("👉 Vui lòng kéo thả thêm 1 file nữa (Cần đủ cả 2 file để đối chiếu).")

    st.write('')
    if f_erp is not None and f_pcu is not None:
        if st.button('🚀 Chạy Đối Chiếu & Mở Bảng Điều Khiển Tương Tác', use_container_width=True):
            prog = st.progress(0, text='⏳ Đang khởi động bộ xử lý dữ liệu siêu tốc...')
            try:
                bundle = doc_va_phan_tich_du_lieu(f_erp, f_pcu, prog)
                st.session_state['clc_bundle'] = bundle
                st.session_state['clc_giam_phat'] = {
                    'NM THSPM': 0.0,
                    'BTP- Mạ màu & Xẻ băng ': 0.0,
                    'BTP- BAO BÌ': 0.0,
                    'VLXD': 0.0,
                }
                st.session_state['clc_excel_bytes'] = xuat_workbook_excel(
                    bundle, st.session_state['clc_giam_phat']
                )
            except Exception as e:
                st.error(f'Lỗi xử lý dữ liệu: {e}')

    # ---------------------------------------------------------
    # KHU VỰC THAO TÁC TRỰC TIẾP TRÊN WEB (CHẠY TỨC THÌ)
    # ---------------------------------------------------------
    if 'clc_bundle' in st.session_state:
        res = st.session_state['clc_bundle']
        gp_dict = st.session_state.get('clc_giam_phat', {})
        tong_giam_phat_kg = sum(to_float(v, 0.0) for v in gp_dict.values())

        tt_tong_nm = to_float(res['pcu_tk_table1'][0][3], 0.0)
        dm_tong_nm = 7640000.0
        cl_tong_nm = tt_tong_nm - dm_tong_nm - tong_giam_phat_kg
        phat_dm = max(0.0, cl_tong_nm * 7.0)

        tt_l2 = to_float(res['pcu_tk_table2'][0][3], 0.0)
        dm_l2 = 1800000.0
        cl_l2 = tt_l2 - dm_l2
        phat_l2 = max(0.0, cl_l2 * 7.0)

        m1_kl = res['ksnb_clc_m1']
        m2_kl = res['ksnb_clc_m2']
        phat_clc = m1_kl * 5.0 + m2_kl * 10.0

        tong_phat = phat_dm + phat_l2 + phat_clc

        st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 18px 0;'>", unsafe_allow_html=True)

        col_title_res, col_dl_btn = st.columns([6.5, 3.5])
        with col_title_res:
            st.markdown(
                f"### 📊 Bảng Điều Khiển Kiểm Soát Tồn Kho & CLC *(Chốt ngày {res['ngay_chot_str']} • Xử lý trong {res['elapsed_sec']}s)*"
            )
        with col_dl_btn:
            st.download_button(
                f"📥 Xuất File Excel Hoàn Chỉnh ({res['file_out_name']})",
                data=st.session_state['clc_excel_bytes'],
                file_name=res['file_out_name'],
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
            )

        k1, k2, k3, k4 = st.columns(4, gap='medium')
        with k1:
            diff_c = res['count_nm'] - res['count_pcu']
            diff_kl = res['kl_nm'] - res['kl_pcu']
            status_txt = '✓ Khớp 100% số cuộn & KL với P.CƯ' if (diff_c == 0 and round(diff_kl, 2) == 0) else f'⚠️ Lệch {diff_c} cuộn ({diff_kl:+,.1f} kg)'
            st.markdown(
                f"""
<div class="bento-card">
    <div class="bento-label">📦 TỔNG CUỘN & KHỐI LƯỢNG</div>
    <div class="bento-value">{res['count_nm']:,} cuộn</div>
    <div class="bento-sub">{res['kl_nm']:,.2f} kg<br><span style="color:#1e7e34;font-weight:600;">{status_txt}</span></div>
</div>
""",
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                f"""
<div class="bento-card">
    <div class="bento-label">⚖️ PHẠT VƯỢT ĐỊNH MỨC (I + II)</div>
    <div class="bento-value">{phat_dm + phat_l2:,.0f} đ</div>
    <div class="bento-sub">• Định mức chung: <b>{phat_dm:,.0f} đ</b><br>• Hàng Loại 2: <b>{phat_l2:,.0f} đ</b></div>
</div>
""",
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                f"""
<div class="bento-card">
    <div class="bento-label">⏳ PHẠT CHẬM LUÂN CHUYỂN (III)</div>
    <div class="bento-value">{phat_clc:,.0f} đ</div>
    <div class="bento-sub">• M1 (91-180 ngày): <b>{m1_kl * 5:,.0f} đ</b><br>• M2 (>180 ngày): <b>{m2_kl * 10:,.0f} đ</b></div>
</div>
""",
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                f"""
<div class="bento-card-highlight">
    <div class="bento-label-white">💰 TỔNG TIỀN PHẠT THÁNG</div>
    <div class="bento-value-white">{tong_phat:,.0f} đ</div>
    <div class="bento-sub-white">• BGĐ (30%): <b>{tong_phat * 0.3:,.0f} đ</b><br>• P.KHCU NM (70%): <b>{tong_phat * 0.7:,.0f} đ</b></div>
</div>
""",
                unsafe_allow_html=True,
            )

        st.write('')

        tab1, tab2, tab3, tab4 = st.tabs([
            f"🚨 1. Rà Soát Cuộn Lệch ({len(res['discrepancy_records'])} cuộn)",
            '📄 2. Tờ Trình Tính Phạt (TTR)',
            '⚖️ 3. Bảng Đối Chiếu 3 Bên (ĐỐI CHIẾU)',
            f"🔍 4. Tra Cứu Toàn Bộ ({res['count_nm']:,} cuộn)",
        ])

        with tab1:
            with st.container(border=True):
                df_disc = pd.DataFrame(res['discrepancy_records'])
                if df_disc.empty:
                    st.success('🎉 Toàn bộ các cuộn đều khớp 100% giữa P.Cung ứng và công thức KSNB!')
                else:
                    c_f1, c_f2, c_f3 = st.columns([4, 3, 3])
                    with c_f1:
                        loai_lech_chon = st.radio(
                            '🎯 Nhấp chọn nhóm cuộn lệch cần kiểm tra:',
                            options=[
                                f'Tất cả ({len(df_disc)})',
                                f"Lệch Đánh loại TÍNH/LOẠI ({df_disc['Nhóm lỗi'].str.contains('Lệch Đánh loại').sum()})",
                                f"Lệch Đánh tên ĐVKD ({df_disc['Nhóm lỗi'].str.contains('Lệch Đánh tên').sum()})",
                                f"Lệch Phạt CLC ({df_disc['Nhóm lỗi'].str.contains('Lệch Phạt CLC').sum()})",
                            ],
                            horizontal=True,
                        )
                    with c_f2:
                        kho_list_err = ['Tất cả'] + sorted(df_disc['Kho'].unique().tolist())
                        chon_kho_err = st.selectbox('🏭 Lọc theo Kho:', kho_list_err, key='err_kho')
                    with c_f3:
                        tim_ma_err = st.text_input('🔎 Tìm nhanh theo Chuỗi / Đánh tên:', key='err_search')

                    df_disc_view = df_disc.copy()
                    if 'Lệch Đánh loại' in loai_lech_chon:
                        df_disc_view = df_disc_view[df_disc_view['Nhóm lỗi'].str.contains('Lệch Đánh loại')]
                    elif 'Lệch Đánh tên' in loai_lech_chon:
                        df_disc_view = df_disc_view[df_disc_view['Nhóm lỗi'].str.contains('Lệch Đánh tên')]
                    elif 'Lệch Phạt CLC' in loai_lech_chon:
                        df_disc_view = df_disc_view[df_disc_view['Nhóm lỗi'].str.contains('Lệch Phạt CLC')]

                    if chon_kho_err != 'Tất cả':
                        df_disc_view = df_disc_view[df_disc_view['Kho'] == chon_kho_err]
                    if tim_ma_err.strip():
                        kw = tim_ma_err.strip().upper()
                        df_disc_view = df_disc_view[
                            df_disc_view['Chuỗi cuộn'].str.upper().str.contains(kw)
                            | df_disc_view['Đánh tên (ERP)'].str.upper().str.contains(kw)
                        ]

                    st.caption(
                        f"💡 Đang hiển thị **{len(df_disc_view)} cuộn** (Tổng KL: **{df_disc_view['KL Net (kg)'].sum():,.0f} kg**). "
                        f"Bạn có thể tích chọn ô **Đã rà soát** hoặc nhấp đúp vào cột **Ghi chú KSNB** để sửa/ghi chú trực tiếp trên bảng."
                    )

                    st.data_editor(
                        df_disc_view,
                        use_container_width=True,
                        hide_index=True,
                        height=420,
                        column_config={
                            'Đã rà soát': st.column_config.CheckboxColumn('☑️ Đã rà soát', width='small'),
                            'KL Net (kg)': st.column_config.NumberColumn('KL Net (kg)', format='%d'),
                            'Ghi chú KSNB': st.column_config.TextColumn('📝 Ghi chú KSNB (Nhấp để sửa)', width='large'),
                        },
                        disabled=[
                            'Dòng Excel', 'Kho', 'Chuỗi cuộn', 'KL Net (kg)', 'ĐVKD (ERP)',
                            'Đánh tên (ERP)', 'KSNB: Đánh tên', 'P.CƯ: Đánh tên',
                            'KSNB: Đánh loại', 'P.CƯ: Đánh loại', 'KSNB: Phạt CLC', 'P.CƯ: Phạt CLC', 'Nhóm lỗi'
                        ],
                        key='editor_discrepancy',
                    )

        with tab2:
            with st.container(border=True):
                st.markdown('#### 📄 I. Định Mức Tồn Kho Nhà Máy & Tờ Trình Xin Giảm Phạt')
                st.caption('💡 Nhập số kg xin giảm phạt vào ô bên dưới, hệ thống sẽ tự động cập nhật lại số tiền phạt và làm mới lại file Excel xuất ra.')

                col_gp1, col_gp2, col_gp3, col_gp4, col_gp_btn = st.columns([2, 2, 2, 2, 2])
                with col_gp1:
                    gp_nm = st.number_input('Giảm phạt NM THSPM (kg):', min_value=0.0, value=float(gp_dict.get('NM THSPM', 0.0)), step=1000.0)
                with col_gp2:
                    gp_btp = st.number_input('Giảm phạt BTP Mạ màu & XB (kg):', min_value=0.0, value=float(gp_dict.get('BTP- Mạ màu & Xẻ băng ', 0.0)), step=1000.0)
                with col_gp3:
                    gp_bb = st.number_input('Giảm phạt BTP Bao bì (kg):', min_value=0.0, value=float(gp_dict.get('BTP- BAO BÌ', 0.0)), step=1000.0)
                with col_gp4:
                    gp_vlxd = st.number_input('Giảm phạt VLXD (kg):', min_value=0.0, value=float(gp_dict.get('VLXD', 0.0)), step=1000.0)
                with col_gp_btn:
                    st.write('')
                    if st.button('🔄 Cập nhật vào Tờ trình', use_container_width=True):
                        st.session_state['clc_giam_phat'] = {
                            'NM THSPM': gp_nm,
                            'BTP- Mạ màu & Xẻ băng ': gp_btp,
                            'BTP- BAO BÌ': gp_bb,
                            'VLXD': gp_vlxd,
                        }
                        st.session_state['clc_excel_bytes'] = xuat_workbook_excel(
                            res, st.session_state['clc_giam_phat']
                        )
                        st.rerun()

                t1_rows = res['pcu_tk_table1']
                df_ttr_1 = pd.DataFrame([
                    {
                        'Nội dung': '1. Định mức theo TTr số 099/TTr/P.CƯ/2023 (kg)',
                        'Tổng cộng': 7640000.0,
                        'NM THSPM': 2000000.0,
                        'BTP- Mạ màu & Xẻ băng': 3500000.0,
                        'BTP- BAO BÌ': 140000.0,
                        'VLXD': 2000000.0,
                        'Các NM Khác': 0.0,
                    },
                    {
                        'Nội dung': f"2. Số liệu tồn kho thực tế {res['ngay_chot_str']} (kg)",
                        'Tổng cộng': to_float(t1_rows[0][3]),
                        'NM THSPM': to_float(t1_rows[1][3]),
                        'BTP- Mạ màu & Xẻ băng': to_float(t1_rows[2][3]),
                        'BTP- BAO BÌ': to_float(t1_rows[3][3]),
                        'VLXD': to_float(t1_rows[4][3]),
                        'Các NM Khác': to_float(t1_rows[5][3]),
                    },
                    {
                        'Nội dung': '3. Tờ trình xin giảm phạt tồn kho (kg)',
                        'Tổng cộng': tong_giam_phat_kg,
                        'NM THSPM': gp_dict.get('NM THSPM', 0.0),
                        'BTP- Mạ màu & Xẻ băng': gp_dict.get('BTP- Mạ màu & Xẻ băng ', 0.0),
                        'BTP- BAO BÌ': gp_dict.get('BTP- BAO BÌ', 0.0),
                        'VLXD': gp_dict.get('VLXD', 0.0),
                        'Các NM Khác': 0.0,
                    },
                    {
                        'Nội dung': '4. Chênh lệch vượt định mức (kg)',
                        'Tổng cộng': cl_tong_nm,
                        'NM THSPM': to_float(t1_rows[1][3]) - 2000000.0 - gp_dict.get('NM THSPM', 0.0),
                        'BTP- Mạ màu & Xẻ băng': to_float(t1_rows[2][3]) - 3500000.0 - gp_dict.get('BTP- Mạ màu & Xẻ băng ', 0.0),
                        'BTP- BAO BÌ': to_float(t1_rows[3][3]) - 140000.0 - gp_dict.get('BTP- BAO BÌ', 0.0),
                        'VLXD': to_float(t1_rows[4][3]) - 2000000.0 - gp_dict.get('VLXD', 0.0),
                        'Các NM Khác': to_float(t1_rows[5][3]),
                    },
                ])
                st.dataframe(
                    df_ttr_1.style.format({c: '{:,.2f}' for c in df_ttr_1.columns if c != 'Nội dung'}),
                    use_container_width=True,
                    hide_index=True,
                )
                st.info(f"📌 **Kết luận Mục I:** Đơn giá phạt **7 đ/kg** ➔ Thành tiền phạt vượt định mức chung = **{phat_dm:,.0f} đồng**.")

                col_ttr_sub1, col_ttr_sub2 = st.columns(2, gap='large')
                with col_ttr_sub1:
                    st.markdown('#### 📄 II. Định Mức Tồn Kho Hàng Loại 2')
                    df_ttr_2 = pd.DataFrame([
                        {'Chỉ tiêu': 'Định mức theo TT (kg)', 'Giá trị': f'{dm_l2:,.0f}'},
                        {'Chỉ tiêu': f"Tồn kho thực tế {res['ngay_chot_str']} (kg)", 'Giá trị': f'{tt_l2:,.2f}'},
                        {'Chỉ tiêu': 'Chênh lệch vượt định mức (kg)', 'Giá trị': f'{cl_l2:,.2f}'},
                        {'Chỉ tiêu': 'Đơn giá phạt (Đồng/kg)', 'Giá trị': '7'},
                        {'Chỉ tiêu': '💰 Thành tiền phạt Hàng Loại 2 (VNĐ)', 'Giá trị': f'{phat_l2:,.0f}'},
                    ])
                    st.dataframe(df_ttr_2, use_container_width=True, hide_index=True)

                with col_ttr_sub2:
                    st.markdown('#### 📄 III. Tồn Kho Chậm Luân Chuyển (CLC)')
                    df_ttr_3 = pd.DataFrame([
                        {
                            'Chỉ tiêu': 'Số ngày tồn kho',
                            'Loại M1': 'Từ 91 - 180 ngày',
                            'Loại M2': '> 180 ngày',
                            'Tổng cộng': '-',
                        },
                        {
                            'Chỉ tiêu': 'Đơn giá phạt (Đồng/kg)',
                            'Loại M1': '5',
                            'Loại M2': '10',
                            'Tổng cộng': '-',
                        },
                        {
                            'Chỉ tiêu': 'Khối lượng tồn kho CLC (kg)',
                            'Loại M1': f'{m1_kl:,.2f}',
                            'Loại M2': f'{m2_kl:,.2f}',
                            'Tổng cộng': f'{m1_kl + m2_kl:,.2f}',
                        },
                        {
                            'Chỉ tiêu': '💰 Thành tiền phạt CLC (VNĐ)',
                            'Loại M1': f'{m1_kl * 5:,.0f}',
                            'Loại M2': f'{m2_kl * 10:,.0f}',
                            'Tổng cộng': f'{phat_clc:,.0f}',
                        },
                    ])
                    st.dataframe(df_ttr_3, use_container_width=True, hide_index=True)

        with tab3:
            with st.container(border=True):
                chon_bang_dc = st.radio(
                    '📊 Chọn bảng đối chiếu cần xem:',
                    options=[
                        'Bảng I: Tổng Tồn Kho Nhà Máy',
                        'Bảng II: Tồn Kho Hàng Loại 2 (L2-A1, XKRL, CXL)',
                        'Bảng III: Tồn Kho Chậm Luân Chuyển (M1 & M2)',
                    ],
                    horizontal=True,
                )

                age_order = ['Dưới 90 ngày', 'Từ 91 đến 120 ngày', 'Từ 121 đến 180 ngày', 'Từ 181 đến 360 ngày', 'Trên 361 ngày']
                if 'Bảng I:' in chon_bang_dc or 'Bảng II:' in chon_bang_dc:
                    is_b1 = 'Bảng I:' in chon_bang_dc
                    src_pcu = res['pcu_tk_table1'] if is_b1 else res['pcu_tk_table2']
                    src_ksnb = res['ksnb_agg_t1'] if is_b1 else res['ksnb_agg_t2']

                    dc_rows_web = []
                    for idx_r in range(1, 6):
                        r_p = src_pcu[idx_r]
                        hm = str(r_p[1]).strip()
                        dm_val = to_float(r_p[2], 0.0)
                        pcu_tot = to_float(r_p[3], 0.0)

                        k_hm = next((k for k in src_ksnb.keys() if k.strip().upper() == hm.upper()), hm)
                        ksnb_dict_age = src_ksnb.get(k_hm, {ag: 0.0 for ag in age_order})
                        ksnb_tot = sum(ksnb_dict_age.values())

                        dc_rows_web.append({
                            'Hạng mục': hm,
                            'Định mức (DM)': dm_val,
                            'P.CƯ: Tổng TT': pcu_tot,
                            'KSNB: Tổng TT': ksnb_tot,
                            'Lệch Tổng (KSNB - P.CƯ)': ksnb_tot - pcu_tot,
                            'KSNB: < 90 ngày': ksnb_dict_age['Dưới 90 ngày'],
                            'KSNB: 91-120 ngày': ksnb_dict_age['Từ 91 đến 120 ngày'],
                            'KSNB: 121-180 ngày': ksnb_dict_age['Từ 121 đến 180 ngày'],
                            'KSNB: 181-360 ngày': ksnb_dict_age['Từ 181 đến 360 ngày'],
                            'KSNB: > 361 ngày': ksnb_dict_age['Trên 361 ngày'],
                        })

                    df_dc_web = pd.DataFrame(dc_rows_web)
                    tot_row = {
                        'Hạng mục': 'TỔNG CỘNG TOÀN NHÀ MÁY',
                        'Định mức (DM)': df_dc_web['Định mức (DM)'].sum() if is_b1 else 1800000.0,
                        'P.CƯ: Tổng TT': df_dc_web['P.CƯ: Tổng TT'].sum(),
                        'KSNB: Tổng TT': df_dc_web['KSNB: Tổng TT'].sum(),
                        'Lệch Tổng (KSNB - P.CƯ)': df_dc_web['Lệch Tổng (KSNB - P.CƯ)'].sum(),
                        'KSNB: < 90 ngày': df_dc_web['KSNB: < 90 ngày'].sum(),
                        'KSNB: 91-120 ngày': df_dc_web['KSNB: 91-120 ngày'].sum(),
                        'KSNB: 121-180 ngày': df_dc_web['KSNB: 121-180 ngày'].sum(),
                        'KSNB: 181-360 ngày': df_dc_web['KSNB: 181-360 ngày'].sum(),
                        'KSNB: > 361 ngày': df_dc_web['KSNB: > 361 ngày'].sum(),
                    }
                    df_dc_web = pd.concat([pd.DataFrame([tot_row]), df_dc_web], ignore_index=True)

                    num_cols = [c for c in df_dc_web.columns if c != 'Hạng mục']
                    st.dataframe(
                        df_dc_web.style.format({c: '{:,.2f}' for c in num_cols}),
                        use_container_width=True,
                        hide_index=True,
                    )
                else:
                    p_m1 = to_float(res['pcu_tk_clc'][0], 0.0)
                    p_m2 = to_float(res['pcu_tk_clc'][1], 0.0)
                    p_tot = to_float(res['pcu_tk_clc'][2], 0.0)
                    k_m1 = res['ksnb_clc_m1']
                    k_m2 = res['ksnb_clc_m2']
                    k_tot = k_m1 + k_m2

                    df_clc_web = pd.DataFrame([
                        {
                            'Nhóm CLC': 'M1 (Từ 91 đến 180 ngày)',
                            'Theo P.Cung ứng (kg)': p_m1,
                            'Theo KSNB (kg)': k_m1,
                            'Chênh lệch (P.CƯ - KSNB)': p_m1 - k_m1,
                        },
                        {
                            'Nhóm CLC': 'M2 (Trên 180 ngày)',
                            'Theo P.Cung ứng (kg)': p_m2,
                            'Theo KSNB (kg)': k_m2,
                            'Chênh lệch (P.CƯ - KSNB)': p_m2 - k_m2,
                        },
                        {
                            'Nhóm CLC': 'TỔNG CỘNG TỒN KHO CLC',
                            'Theo P.Cung ứng (kg)': p_tot,
                            'Theo KSNB (kg)': k_tot,
                            'Chênh lệch (P.CƯ - KSNB)': p_tot - k_tot,
                        },
                    ])
                    st.dataframe(
                        df_clc_web.style.format({
                            'Theo P.Cung ứng (kg)': '{:,.2f}',
                            'Theo KSNB (kg)': '{:,.2f}',
                            'Chênh lệch (P.CƯ - KSNB)': '{:,.2f}',
                        }),
                        use_container_width=True,
                        hide_index=True,
                    )

        with tab4:
            with st.container(border=True):
                df_all = pd.DataFrame(res['web_all_records'])
                s1, s2, s3, s4 = st.columns([2.2, 2.5, 2.3, 3.0])
                with s1:
                    f_kho = st.selectbox('🏭 Kho vị trí:', ['Tất cả'] + sorted(df_all['Kho'].unique().tolist()))
                with s2:
                    f_dvkd = st.selectbox('🏢 KSNB Đánh tên ĐVKD:', ['Tất cả'] + sorted(df_all['KSNB: Đánh tên'].unique().tolist()))
                with s3:
                    f_loai = st.selectbox('📌 KSNB Đánh loại:', ['Tất cả', 'TÍNH', 'LOẠI'])
                with s4:
                    f_kw = st.text_input('🔎 Tìm Mã lô / Mã hàng / Tên hàng:')

                df_view = df_all
                if f_kho != 'Tất cả':
                    df_view = df_view[df_view['Kho'] == f_kho]
                if f_dvkd != 'Tất cả':
                    df_view = df_view[df_view['KSNB: Đánh tên'] == f_dvkd]
                if f_loai != 'Tất cả':
                    df_view = df_view[df_view['KSNB: Đánh loại'] == f_loai]
                if f_kw.strip():
                    kw_u = f_kw.strip().upper()
                    df_view = df_view[
                        df_view['Chuỗi cuộn'].str.upper().str.contains(kw_u)
                        | df_view['Mã hàng'].str.upper().str.contains(kw_u)
                        | df_view['Tên hàng'].str.upper().str.contains(kw_u)
                    ]

                st.caption(
                    f"📦 Kết quả lọc: **{len(df_view):,} cuộn** • Tổng khối lượng Net: **{df_view['KL Net (kg)'].sum():,.2f} kg**"
                )
                st.dataframe(
                    df_view,
                    use_container_width=True,
                    hide_index=True,
                    height=440,
                )


kiem_tra_dinh_muc_ton_kho()
