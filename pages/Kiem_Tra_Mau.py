import io
import re
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st


def read_auto_header(file_obj, keywords, max_scan=15):
    """Tự động tìm dòng chứa tiêu đề dựa vào từ khóa, tránh lỗi cố định skiprows."""
    df_raw = pd.read_excel(file_obj, header=None, nrows=max_scan)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
        if any(k.upper() in row_str for k in keywords):
            header_idx = idx
            break
    file_obj.seek(0)
    df = pd.read_excel(file_obj, skiprows=header_idx)
    return df


def doi_chieu_du_lieu(f_bbnt, f_pl, f_tc, f_bv, f_opm):
    # 1. ĐỌC BBNT
    df_bbnt = read_auto_header(f_bbnt, ['MÃ SỐ CUỘN', 'MÃ CUỘN', 'SỐ PHIẾU CÂN'])
    # Xóa dòng trống mã cuộn và dòng tổng kết
    col_cuon = next(c for c in df_bbnt.columns if 'CUỘN' in str(c).upper())
    col_pcan = next(c for c in df_bbnt.columns if 'PHIẾU CÂN' in str(c).upper())
    col_net = next(c for c in df_bbnt.columns if 'CHỨNG TỪ GIAO HÀNG' in str(c).upper() and 'GROSS' not in str(c).upper())
    col_gross = next(c for c in df_bbnt.columns if 'GROSS' in str(c).upper())
    col_canban = next(c for c in df_bbnt.columns if 'CÂN BÀN' in str(c).upper())
    col_dsai = next(c for c in df_bbnt.columns if 'CHÊNH LỆCH' in str(c).upper() and '%' in str(c))

    df_bbnt = df_bbnt.dropna(subset=[col_cuon]).copy()
    df_bbnt = df_bbnt[~df_bbnt[df_bbnt.columns[0]].astype(str).str.contains('Tổng cộng|Kết luận', case=False, na=False)]

    base_cols = df_bbnt.columns[:13].tolist()
    res_df = df_bbnt[base_cols].copy()

    ma_cuon = df_bbnt[col_cuon].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
    phieu_can_bbnt = df_bbnt[col_pcan].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    ma_hoa = ma_cuon + 'A'

    # 2. ĐỌC PACKING LIST
    dict_pl = {}
    if f_pl is not None:
        df_pl = read_auto_header(f_pl, ['MÃ SỐ NSX', 'MÃ SỐ HSPM', 'KL(NET)'])
        c_pl_ma = next(c for c in df_pl.columns if 'MÃ SỐ NSX' in str(c).upper() or 'MÃ SỐ' in str(c).upper())
        c_pl_kl = next(c for c in df_pl.columns if 'NET' in str(c).upper())
        
        df_pl_clean = df_pl.dropna(subset=[c_pl_ma, c_pl_kl]).copy()
        pl_key = df_pl_clean[c_pl_ma].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper() + 'A'
        pl_val = pd.to_numeric(df_pl_clean[c_pl_kl], errors='coerce').fillna(0)
        dict_pl = pd.DataFrame({'Key': pl_key, 'Net': pl_val}).groupby('Key')['Net'].sum().to_dict()

    # 3. ĐỌC NHẬP BẢO VỆ
    dict_bv_phieu = {}
    dict_bv_kl = {}
    if f_bv is not None:
        df_bv = read_auto_header(f_bv, ['MÃ PHIẾU CÂN', 'KHỐI LƯỢNG PHIẾU CÂN', 'TÊN HÀNG HÓA'])
        col_bv_hang = next(c for c in df_bv.columns if 'TÊN HÀNG HÓA' in str(c).upper() or 'MÃ HÀNG HÓA' in str(c).upper())
        col_bv_pcan = next(c for c in df_bv.columns if 'MÃ PHIẾU CÂN' in str(c).upper())
        col_bv_kl = next(c for c in df_bv.columns if 'KHỐI LƯỢNG' in str(c).upper())

        for _, r in df_bv.iterrows():
            txt = str(r[col_bv_hang]).upper()
            p = str(r[col_bv_pcan]).replace('.0', '').strip()
            kl = r[col_bv_kl]
            for m in ma_cuon:
                if m in txt and m not in dict_bv_phieu:
                    dict_bv_phieu[m] = p
                    dict_bv_kl[m] = kl

    # 4. ĐỌC TRẠM CÂN
    dict_tc_kl = {}
    dict_tc_phanloai = {}
    if f_tc is not None:
        df_tc = read_auto_header(f_tc, ['PHIẾU CÂN', 'KHỐI LƯỢNG PHIẾU CÂN', 'PHÂN LOẠI'])
        col_tc_p = next(c for c in df_tc.columns if str(c).strip().upper() == 'PHIẾU CÂN')
        col_tc_kl = next(c for c in df_tc.columns if 'KHỐI LƯỢNG' in str(c).upper())
        col_tc_pl = next((c for c in df_tc.columns if 'PHÂN LOẠI' in str(c).upper()), None)

        tc_p = df_tc[col_tc_p].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        tc_kl = pd.to_numeric(df_tc[col_tc_kl], errors='coerce').fillna(0)
        dict_tc_kl = dict(zip(tc_p, tc_kl))
        if col_tc_pl:
            dict_tc_phanloai = dict(zip(tc_p, df_tc[col_tc_pl].fillna('Nhập hàng').astype(str)))

    # 5. ĐỌC ERP OPM 105
    opm_035_po, opm_035_rtv, opm_031, opm_03g_po, opm_03g_rtv = {}, {}, {}, {}, {}
    if f_opm is not None:
        df_opm = read_auto_header(f_opm, ['LOẠI GIAO DỊCH', 'MÃ LÔ', 'KHO'])
        # Lọc bỏ dòng phụ nếu có nhãn [1], [2]
        if len(df_opm) > 0 and '[' in str(df_opm.iloc[0, 1]):
            df_opm = df_opm.iloc[1:].copy()

        c_opm_loai = next(c for c in df_opm.columns if 'LOẠI GIAO DỊCH' in str(c).upper())
        c_opm_malo = next(c for c in df_opm.columns if 'MÃ LÔ' in str(c).upper())
        c_opm_kl = next(c for c in df_opm.columns if 'SỐ LƯỢNG' in str(c).upper() or 'KHỐI LƯỢNG' in str(c).upper())
        c_opm_kho = next(c for c in df_opm.columns if 'KHO' in str(c).upper() and 'NGÀY' not in str(c).upper())

        df_opm_clean = df_opm.dropna(subset=[c_opm_malo, c_opm_loai]).copy()
        lo_str = df_opm_clean[c_opm_malo].astype(str).str.strip().str.upper()
        df_opm_clean['Key'] = lo_str.apply(lambda x: x if x.endswith('A') else x + 'A')
        df_opm_clean['KL'] = pd.to_numeric(df_opm_clean[c_opm_kl], errors='coerce').fillna(0)
        df_opm_clean['Kho_Clean'] = df_opm_clean[c_opm_kho].astype(str).str.strip()

        # Tổng hợp theo nghiệp vụ chuẩn của mẫu Excel
        opm_035_po = df_opm_clean[(df_opm_clean['Kho_Clean'] == '035') & (df_opm_clean[c_opm_loai].str.contains('PO Receipt', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
        opm_035_rtv = df_opm_clean[(df_opm_clean['Kho_Clean'] == '035') & (df_opm_clean[c_opm_loai].str.contains('Return to Vendor', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
        opm_031 = df_opm_clean[(df_opm_clean['Kho_Clean'].isin(['03G', '03H'])) & (df_opm_clean[c_opm_loai].str.contains('Intransit Receipt', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
        opm_03g_po = df_opm_clean[(df_opm_clean['Kho_Clean'] == '03G') & (df_opm_clean[c_opm_loai].str.contains('PO Receipt', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()
        opm_03g_rtv = df_opm_clean[(df_opm_clean['Kho_Clean'] == '03G') & (df_opm_clean[c_opm_loai].str.contains('Return to Vendor', case=False, na=False))].groupby('Key')['KL'].sum().to_dict()

    # 6. TÍNH CÁC CỘT THEO ĐÚNG CÔNG THỨC EXCEL MẪU
    res_df['NHẬP/XUẤT'] = phieu_can_bbnt.map(lambda x: dict_tc_phanloai.get(x, 'Nhập hàng'))
    res_df['Mã hóa'] = ma_hoa
    res_df['Khối lượng net theo PL'] = ma_hoa.map(dict_pl).fillna(0).astype(float)

    net_ct = pd.to_numeric(df_bbnt[col_net], errors='coerce').fillna(0).astype(float)
    res_df['Đối chiếu'] = net_ct - res_df['Khối lượng net theo PL']

    res_df['KL theo bc nhập hàng'] = ma_cuon.map(dict_bv_kl).fillna(0).astype(float)
    res_df['Số phiếu cân'] = ma_cuon.map(dict_bv_phieu).fillna('')
    res_df['Đối chiếu số phiếu cân'] = phieu_can_bbnt == res_df['Số phiếu cân'].astype(str)

    # Khối lượng trạm cân: tra theo số phiếu cân của BẢO VỆ (theo đúng công thức VLOOKUP V10 của Excel)
    res_df['KL theo bc trạm cân'] = res_df['Số phiếu cân'].map(dict_tc_kl).fillna(0).astype(float)
    res_df['Đối chiếu KL TC & KL BCBV'] = res_df['KL theo bc trạm cân'] == res_df['KL theo bc nhập hàng']

    can_ban = pd.to_numeric(df_bbnt[col_canban], errors='coerce').fillna(0).astype(float)
    res_df['Đối chiếu KL TC & KL BBNT'] = res_df['KL theo bc trạm cân'] == can_ban

    gross_ct = pd.to_numeric(df_bbnt[col_gross], errors='coerce').replace(0, pd.NA).fillna(1).astype(float)
    res_df['Chênh lệch (kg)'] = res_df['KL theo bc trạm cân'] - gross_ct
    res_df['Chênh lệch dung sai'] = res_df['Chênh lệch (kg)'] / gross_ct

    dsai_bbnt = pd.to_numeric(df_bbnt[col_dsai], errors='coerce').fillna(0).astype(float)
    res_df['Đối chiếu dung sai'] = res_df['Chênh lệch dung sai'].round(6) == dsai_bbnt.round(6)
    res_df['So với quy định ( >+-0.3%)'] = res_df['Chênh lệch dung sai'].apply(lambda x: 'VUOT DUNG SAI' if abs(x) > 0.003 else 'OK')

    res_df['KL Nhập kho 035'] = ma_hoa.map(lambda x: opm_035_po.get(x, 0) - abs(opm_035_rtv.get(x, 0)))
    res_df['Chênh lệch '] = res_df['Khối lượng net theo PL'] - res_df['KL Nhập kho 035']

    res_df['KL Nhập kho 031'] = ma_hoa.map(lambda x: opm_031.get(x, 0))
    res_df['Chênh lệch  '] = res_df.apply(lambda r: 0 if r['KL Nhập kho 031'] == 0 else r['Khối lượng net theo PL'] - r['KL Nhập kho 031'], axis=1)

    res_df['KL Nhập kho 03G-3H'] = ma_hoa.map(lambda x: opm_03g_po.get(x, 0) - abs(opm_03g_rtv.get(x, 0)))
    res_df['Chênh lệch   '] = res_df['Khối lượng net theo PL'] - res_df['KL Nhập kho 03G-3H']

    return res_df
