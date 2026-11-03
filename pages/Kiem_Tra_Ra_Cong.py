import io
import unicodedata
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="Kiểm Tra Hàng Ra Cổng - KSNB", page_icon="🚛", layout="wide")

# CSS Giao diện
st.markdown("""
<style>
    [data-testid="stSidebarNav"] { display: none !important; }
    .stApp { background-color: #f5f5f7; font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif; color: #1d1d1f; }
    .apple-header-box { display: flex; justify-content: space-between; align-items: flex-end; padding: 20px 0 10px 0; margin-bottom: 5px; }
    .apple-main-title { font-size: 38px; font-weight: 700; letter-spacing: -0.025em; margin: 0; }
    .apple-sub-text { font-size: 16px; font-weight: 600; text-align: right; margin: 0; line-height: 1.35; }
    .apple-sub-text span { color: #0071e3; font-size: 14px; display: block; margin-top: 2px; font-weight: 500; }
    div.stButton > button { background: #0071e3 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; padding: 10px 24px !important; font-weight: 600 !important; font-size: 15px !important; box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2) !important; }
    .upload-box { background-color: white; padding: 15px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.03); margin-bottom: 15px; border: 1px solid #e2e8f0; }
    .section-title { font-size: 18px; font-weight: 600; margin-bottom: 10px; color: #1d1d1f; }
    .progress-text { font-size: 15px; font-weight: 600; color: #0071e3; margin-bottom: 8px; }
</style>
""", unsafe_allow_html=True)

col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Trang chủ", use_container_width=True):
        st.switch_page("app_ksnb.py")

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)

st.markdown("""
<div class="apple-header-box">
    <div><h1 class="apple-main-title">🚛 Kiểm Tra Hàng Ra Cổng (Trạm Cân)</h1></div>
    <div><div class="apple-sub-text">Đối soát dung sai & Chênh lệch Net/Gross.<span>Lê Công Minh Workspace ›</span></div></div>
</div>
<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>
""", unsafe_allow_html=True)

# KHU VỰC TẢI FILE
st.markdown('<div class="section-title">📥 1. Dữ liệu Trạm Cân (Thực tế)</div>', unsafe_allow_html=True)
f_tramcan_nm = st.file_uploader("File Trạm Cân (Nhà Máy)", type=["xlsx", "xls", "xlsb"], key="tc_nm")

st.markdown('<div class="section-title" style="margin-top: 20px;">📥 2. Dữ liệu ERP (Lý thuyết)</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    f_105_cu = st.file_uploader("1. OPM 105 (HSG_KS_INV_Cung Ứng)", type=["xlsx", "xls"])
    f_021_cu = st.file_uploader("3. INV 021 (HSG_KS_INV_Cung Ứng)", type=["xlsx", "xls"])
    st.markdown('</div>', unsafe_allow_html=True)
with col2:
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    f_105_admin = st.file_uploader("2. OPM 105 (HSH_HSSC... Admin)", type=["xlsx", "xls"])
    f_021_admin = st.file_uploader("4. INV 021 (HSH_HSSC... Admin)", type=["xlsx", "xls"])
    st.markdown('</div>', unsafe_allow_html=True)

# HÀM BỎ DẤU TIẾNG VIỆT ĐỂ SO SÁNH TÊN CỘT (CHỐNG LỖI 'MĂ KHO', 'MĂ LÔ')
def remove_accents(input_str):
    s = str(input_str).strip().lower()
    s = s.replace('đ', 'd')
    nfkd_form = unicodedata.normalize('NFKD', s)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

# HÀM LÀM SẠCH CHUỖI CHUNG (Xóa khoảng trắng, xóa .0 ở cuối số, xóa nan, in hoa)
def clean_key_series(series):
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r'\.0$', '', regex=True)
        .str.upper()
        .replace(['NAN', 'NONE', '<NA>'], '', regex=False)
        .fillna('')
    )

# HÀM TÌM CỘT AN TOÀN (Bỏ qua khoảng trắng thừa, hoa/thường VÀ dấu tiếng Việt như Mã vs Mă)
def get_clean_col(df, col_name):
    target = remove_accents(col_name)
    matched = [c for c in df.columns if remove_accents(c) == target]
    if matched:
        return clean_key_series(df[matched[0]])
    return pd.Series('', index=df.index)

# HÀM LỌC DỮ LIỆU ERP 105 (Lấy Gross từ cột AN - cột thứ 40 / index 39)
def xuly_105(df, is_cungung=True):
    try:
        header_idx = df.astype(str).apply(lambda row: row.str.contains('Loại Giao Dịch', case=False, na=False).any(), axis=1).idxmax()
        if pd.isna(header_idx): return pd.DataFrame()
        df.columns = df.iloc[header_idx].astype(str).str.strip()
        
        # Nếu dòng dưới tiêu đề là dòng đánh số thứ tự cột [1], [2]... [40] thì bỏ qua luôn dòng đó
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
        if not df.empty and str(df.iloc[0, 0]).strip().startswith('[1]'):
            df = df.iloc[1:].reset_index(drop=True)
        
        # Lấy chính xác cột thứ 40 (Cột AN trong Excel, index = 39) hoặc cột có chữ 'Gross'
        gross_cols = [i for i, c in enumerate(df.columns) if 'gross' in remove_accents(c)]
        if gross_cols:
            raw_gross = df.iloc[:, gross_cols[0]]
        elif df.shape[1] > 39:
            raw_gross = df.iloc[:, 39]
        else:
            raw_gross = pd.Series(0, index=df.index)
            
        # Xử lý trường hợp số Gross có dấu phẩy hoặc chấm ngăn cách hàng nghìn
        gross_an_series = pd.to_numeric(
            raw_gross.astype(str).str.strip().str.replace(',', '', regex=False), 
            errors='coerce'
        ).fillna(0)
            
        df = df.loc[:, df.columns.notna() & (df.columns != 'nan')]
        df = df.loc[:, ~df.columns.duplicated()]
        df['Gross_AN'] = gross_an_series
        
        col_lgd_list = [c for c in df.columns if remove_accents(c) == 'loai giao dich']
        if df.empty or not col_lgd_list: return df
        col_lgd = col_lgd_list[0]
        
        lgd_clean = df[col_lgd].astype(str).str.strip()
        
        if is_cungung:
            remove_gd = ['COGS Recognition', 'WIP Return', 'WIP Issue', 'WIP Completion Return', 'WIP Completion', 'WIP Byproduct Return', 'WIP Byproduct Completion', 'Residual Qty Receipt', 'Residual Qty Issue']
            df = df[~lgd_clean.isin(remove_gd)]
            col_mh_list = [c for c in df.columns if remove_accents(c) == 'ma mat hang']
            if col_mh_list:
                df = df[df[col_mh_list[0]].astype(str).str.strip().str.startswith(('3', '4'), na=False)]
        else:
            df = df[~lgd_clean.isin(['COGS Recognition'])]
            
        df = df[df[col_lgd].notna()].copy()
        
        ma_kho = get_clean_col(df, 'Mã Kho')
        kv_kho = get_clean_col(df, 'Khu Vực Kho')
        locator = get_clean_col(df, 'Locator')
        lpn = get_clean_col(df, 'LPN')
        ma_lo = get_clean_col(df, 'Mã Lô')
        loai_gd = get_clean_col(df, 'Loại Giao Dịch')
        
        # Ưu tiên lấy Mã Lô (nếu Mã Lô trống mới lấy LPN)
        lot_or_lpn = pd.Series(np.where(ma_lo != '', ma_lo, lpn), index=df.index)
        
        # Tầng 1: Chuỗi chuẩn 5 thành phần (Kho + Khu vực + Locator + Mã Lô/LPN + Loại Giao Dịch)
        df['Chuỗi'] = ma_kho + kv_kho + locator + lot_or_lpn + loai_gd
        
        # Tầng 2: Chuỗi ép đúng cột Mã Lô + Loại Giao Dịch
        df['Chuỗi_MaLo'] = ma_kho + kv_kho + locator + ma_lo + loai_gd
        
        # Tầng 3: Chuỗi dự phòng khi Locator lệch nhau nhưng đúng Kho + Khu vực + Mã Lô + Loại Giao Dịch
        df['Chuỗi_Kho_Lo_GD'] = ma_kho + kv_kho + lot_or_lpn + loai_gd
        
        # Tầng 4: Chuỗi cho hàng rời không quản lý Lô/LPN
        df['Chuỗi_KhongLo'] = ma_kho + kv_kho + locator + loai_gd
        
        # Chuỗi phục vụ riêng cho việc lấy thông tin Pick SO
        df['Chuỗi_Pick'] = ma_kho + kv_kho + locator + lot_or_lpn + "SALES ORDER PICK"
        df['Chuỗi_Pick_MaLo'] = ma_kho + kv_kho + locator + ma_lo + "SALES ORDER PICK"
        return df
    except: return pd.DataFrame()

def xuly_021(df, is_cungung=True):
    try:
        header_idx = df.astype(str).apply(lambda row: row.str.contains('Số chứng từ', case=False, na=False).any(), axis=1).idxmax()
        if pd.isna(header_idx): return pd.DataFrame()
        df.columns = df.iloc[header_idx].astype(str).str.strip()
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
        df = df.loc[:, df.columns.notna() & (df.columns != 'nan')]
        df = df.loc[:, ~df.columns.duplicated()]
        
        col_lgd_list = [c for c in df.columns if remove_accents(c) == 'loai giao dich']
        if df.empty or not col_lgd_list: return df
        col_lgd = col_lgd_list[0]
        lgd_clean = df[col_lgd].astype(str).str.strip()
        
        if is_cungung: remove_gd = ['WIP Issue', 'WIP Completion Return', 'Return to vender', 'Return to vendor']
        else: remove_gd = ['Return to Vendor', 'Return to vender']
        df = df[~lgd_clean.isin(remove_gd)].copy()
        return df
    except: return pd.DataFrame()

# HÀM TÔ MÀU EXCEL
def style_excel_cells(val):
    val_str = str(val).strip()
    if '🔴' in val_str or val_str in ['Vượt dung sai']:
        return 'background-color: #FFC7CE; color: #9C0006' 
    elif '🟢' in val_str or val_str in ['OK']:
        return 'background-color: #C6EFCE; color: #006100' 
    return ''

# XỬ LÝ CHÍNH
if st.button("🚀 Bắt đầu đối soát & Xuất Báo Cáo", use_container_width=True):
    if not f_tramcan_nm:
        st.warning("⚠️ Vui lòng nạp File Trạm Cân Nhà Máy!")
    else:
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # --- BƯỚC 1: ĐỌC DỮ LIỆU TRẠM CÂN & LỌC PHIẾU LỚN NHẤT KHI TRÙNG CHỨNG TỪ ---
            progress_text.markdown('<div class="progress-text">⏳ Đang dọn dẹp dữ liệu Trạm Cân (Lọc phiếu lớn nhất)... (10%)</div>', unsafe_allow_html=True)
            progress_bar.progress(10)
            xls_tc = pd.ExcelFile(f_tramcan_nm)
            sheet_tc = 'tramcan' if 'tramcan' in xls_tc.sheet_names else 0
            df_tc_raw = pd.read_excel(f_tramcan_nm, sheet_name=sheet_tc)
            
            df_tc_full = df_tc_raw.dropna(subset=['PHIẾU CÂN']).copy()
            
            # Quét dữ liệu 4 cổng chéo nhau theo PHIẾU CÂN
            dict_cong_can = {}
            for _, row in df_tc_full.iterrows():
                pc = str(row.get('PHIẾU CÂN', '')).strip().upper()
                cong = str(row.get('BÀN CÂN RA', '')).strip().upper()
                kl = pd.to_numeric(row.get('KHỐI LƯỢNG PHIẾU CÂN', np.nan), errors='coerce')
                if pc and pc != 'NAN' and pd.notna(kl):
                    if pc not in dict_cong_can: dict_cong_can[pc] = {}
                    dict_cong_can[pc][cong] = kl

            # Xử lý trùng chứng từ: Chỉ giữ lại phiếu cân có khối lượng lớn nhất
            col_chungtu = 'Lấy phiếu xuất.1' if 'Lấy phiếu xuất.1' in df_tc_full.columns else 'SỐ CHỨNG TỪ'
            df_tc_full['KL_TAM'] = pd.to_numeric(df_tc_full['KHỐI LƯỢNG PHIẾU CÂN'], errors='coerce').fillna(0)
            df_tc_full['CT_TAM'] = clean_key_series(df_tc_full[col_chungtu])
            
            df_has_ct = df_tc_full[df_tc_full['CT_TAM'] != ''].copy()
            df_no_ct = df_tc_full[~df_tc_full.index.isin(df_has_ct.index)].copy()
            
            df_has_ct = df_has_ct.sort_values(by='KL_TAM', ascending=False).drop_duplicates(subset=['CT_TAM'], keep='first')
            
            df_tc_main = pd.concat([df_has_ct, df_no_ct], ignore_index=True)
            df_tc_main = df_tc_main.drop_duplicates(subset=['PHIẾU CÂN']).drop(columns=['KL_TAM', 'CT_TAM'])
            
            tc_chungtu_list = [str(x).strip().upper() for x in df_tc_main[col_chungtu].tolist() if pd.notna(x)]

            # --- BƯỚC 2: ĐỌC VÀ LỌC ERP ---
            progress_text.markdown('<div class="progress-text">⏳ Đang xử lý bộ lọc cho ERP 105 và 021... (30%)</div>', unsafe_allow_html=True)
            progress_bar.progress(30)
            
            frames_105, frames_021 = [], []
            if f_105_cu: frames_105.append(xuly_105(pd.read_excel(f_105_cu, header=None), True))
            if f_105_admin: frames_105.append(xuly_105(pd.read_excel(f_105_admin, header=None), False))
            df_105_all = pd.concat(frames_105, ignore_index=True) if frames_105 else pd.DataFrame()
            
            if f_021_cu: frames_021.append(xuly_021(pd.read_excel(f_021_cu, header=None), True))
            if f_021_admin: frames_021.append(xuly_021(pd.read_excel(f_021_admin, header=None), False))
            df_021_all = pd.concat(frames_021, ignore_index=True) if frames_021 else pd.DataFrame()

            # --- TÍNH TOÁN SHEET 021 (KHÓA CHẶT THEO DÒNG LOẠI GIAO DỊCH SANG 105) ---
            progress_text.markdown('<div class="progress-text">⏳ Đang tự động hóa VLOOKUP đúng dòng giao dịch cho bảng 021... (55%)</div>', unsafe_allow_html=True)
            progress_bar.progress(55)
            
            df_021_export = pd.DataFrame()
            if not df_021_all.empty:
                df_021_export = df_021_all.copy()
                
                def safe_str(col_name):
                    return get_clean_col(df_021_export, col_name)

                # Tự động bắt cả 'Mã kho' lẫn 'Mă kho', 'Mã lô' lẫn 'Mă lô' nhờ hàm remove_accents
                ma_kho_021 = safe_str('Mã kho')
                sub_021 = safe_str('Subinventory')
                loc_021 = safe_str('Locator')
                lpn_021 = safe_str('LPN')
                malo_021 = safe_str('Mã lô')
                loai_gd_021 = safe_str('Loại giao dịch')
                
                # Ưu tiên lấy Mã lô (nếu Mã lô trống mới lấy LPN)
                lot_or_lpn_021 = pd.Series(np.where(malo_021 != '', malo_021, lpn_021), index=df_021_export.index)

                # Các tầng chuỗi của 021 (Tất cả đều khóa chặt theo loai_gd_021)
                df_021_export['Chuỗi 2 '] = ma_kho_021 + sub_021 + loc_021 + lot_or_lpn_021 + loai_gd_021
                s_021_malo = ma_kho_021 + sub_021 + loc_021 + malo_021 + loai_gd_021
                s_021_kho_lo_gd = ma_kho_021 + sub_021 + lot_or_lpn_021 + loai_gd_021
                s_021_nolo = ma_kho_021 + sub_021 + loc_021 + loai_gd_021
                
                df_021_export['Chuỗi Pick'] = ma_kho_021 + sub_021 + loc_021 + lot_or_lpn_021 + "SALES ORDER PICK"
                s_pick_malo = ma_kho_021 + sub_021 + loc_021 + malo_021 + "SALES ORDER PICK"

                if not df_105_all.empty and 'Chuỗi' in df_105_all.columns:
                    # Sắp xếp ưu tiên dòng có Gross_AN > 0 lên đầu trước khi drop_duplicates
                    df_105_sorted = df_105_all.sort_values(by='Gross_AN', ascending=False)
                    
                    map_105_t1 = df_105_sorted[df_105_sorted['Chuỗi'] != ""].drop_duplicates(subset=['Chuỗi'], keep='first').set_index('Chuỗi')
                    map_105_t2 = df_105_sorted[df_105_sorted['Chuỗi_MaLo'] != ""].drop_duplicates(subset=['Chuỗi_MaLo'], keep='first').set_index('Chuỗi_MaLo')
                    map_105_t3 = df_105_sorted[df_105_sorted['Chuỗi_Kho_Lo_GD'] != ""].drop_duplicates(subset=['Chuỗi_Kho_Lo_GD'], keep='first').set_index('Chuỗi_Kho_Lo_GD')
                    map_105_t4 = df_105_sorted[df_105_sorted['Chuỗi_KhongLo'] != ""].drop_duplicates(subset=['Chuỗi_KhongLo'], keep='first').set_index('Chuỗi_KhongLo')
                    
                    map_pick_t1 = df_105_sorted[df_105_sorted['Chuỗi_Pick'] != ""].drop_duplicates(subset=['Chuỗi_Pick'], keep='first').set_index('Chuỗi_Pick')
                    map_pick_t2 = df_105_sorted[df_105_sorted['Chuỗi_Pick_MaLo'] != ""].drop_duplicates(subset=['Chuỗi_Pick_MaLo'], keep='first').set_index('Chuỗi_Pick_MaLo')

                    def map_fallback(col_105):
                        if col_105 not in df_105_sorted.columns:
                            return pd.Series("", index=df_021_export.index)
                        res = df_021_export['Chuỗi 2 '].map(map_105_t1[col_105])
                        res = res.fillna(s_021_malo.map(map_105_t2[col_105]))
                        res = res.fillna(s_021_kho_lo_gd.map(map_105_t3[col_105]))
                        mask_empty_lot = (lot_or_lpn_021 == '')
                        if mask_empty_lot.any():
                            res.loc[mask_empty_lot] = res.loc[mask_empty_lot].fillna(s_021_nolo.loc[mask_empty_lot].map(map_105_t4[col_105]))
                        return res

                    df_021_export['Ngày xuất'] = map_fallback('Ngày Giao Dịch').fillna("")
                    df_021_export['Thời gian thao tác'] = map_fallback('Creation Date').fillna("")
                    df_021_export['Người thực hiện'] = map_fallback('Created by User').fillna("")
                    
                    if 'Gross_AN' in df_105_sorted.columns:
                        g_t1 = pd.to_numeric(df_021_export['Chuỗi 2 '].map(map_105_t1['Gross_AN']), errors='coerce')
                        g_t2 = pd.to_numeric(s_021_malo.map(map_105_t2['Gross_AN']), errors='coerce')
                        g_t3 = pd.to_numeric(s_021_kho_lo_gd.map(map_105_t3['Gross_AN']), errors='coerce')
                        g_t4 = pd.to_numeric(s_021_nolo.map(map_105_t4['Gross_AN']), errors='coerce')
                        
                        gross_final = g_t1.where(g_t1.notna() & (g_t1 != 0), g_t2)
                        gross_final = gross_final.where(gross_final.notna() & (gross_final != 0), g_t3)
                        
                        mask_empty_lot = (lot_or_lpn_021 == '')
                        if mask_empty_lot.any():
                            gross_final.loc[mask_empty_lot] = gross_final.loc[mask_empty_lot].where(
                                gross_final.loc[mask_empty_lot].notna() & (gross_final.loc[mask_empty_lot] != 0),
                                g_t4.loc[mask_empty_lot]
                            )
                            
                        df_021_export['Số Gross'] = gross_final.fillna(0)
                    else:
                        df_021_export['Số Gross'] = 0
                        
                    def map_pick_fallback(col_105):
                        if col_105 not in df_105_sorted.columns:
                            return pd.Series("", index=df_021_export.index)
                        res = df_021_export['Chuỗi Pick'].map(map_pick_t1[col_105])
                        res = res.fillna(s_pick_malo.map(map_pick_t2[col_105]))
                        return res.fillna("")

                    df_021_export['Thời gian pick SO'] = map_pick_fallback('Creation Date')
                    df_021_export['Người Pick SO'] = map_pick_fallback('Created by User')
                else:
                    df_021_export['Ngày xuất'] = df_021_export['Thời gian thao tác'] = df_021_export['Người thực hiện'] = ""
                    df_021_export['Số Gross'] = 0
                    df_021_export['Thời gian pick SO'] = df_021_export['Người Pick SO'] = ""

                df_021_export['Chuyển text'] = safe_str('Số chứng từ') + "A" 
                df_021_export['Chuỗi 1'] = df_021_export['Chuyển text'] + safe_str('Loại giao dịch')
                
                col_sct_list = [c for c in df_021_export.columns if remove_accents(c) == 'so chung tu']
                col_sct_021 = col_sct_list[0] if col_sct_list else 'Số chứng từ'
                df_021_export['Check ra cổng'] = df_021_export[col_sct_021].apply(lambda x: 1 if any(str(x).strip().upper() in str(tc_val) for tc_val in tc_chungtu_list) else 0)
                
                cols_10_first = ['Check ra cổng', 'Thời gian pick SO', 'Người Pick SO', 'Ngày xuất', 'Thời gian thao tác', 'Người thực hiện', 'Số Gross', 'Chuỗi 2 ', 'Chuỗi 1', 'Chuyển text']
                final_021_cols = cols_10_first + [c for c in df_021_all.columns if c not in cols_10_first and c != 'Chuỗi 2 ' and c != 'Chuỗi Pick']
                df_021_export = df_021_export[final_021_cols]

            # --- TÍNH TOÁN SHEET KSKT ---
            progress_text.markdown('<div class="progress-text">⏳ Đang ráp dữ liệu 4 cổng chéo nhau & VLOOKUP Net/Gross... (75%)</div>', unsafe_allow_html=True)
            progress_bar.progress(75)
            
            dict_net, dict_gross = {}, {}
            if not df_021_export.empty:
                df_calc = df_021_export.copy()
                col_sct_calc = [c for c in df_calc.columns if remove_accents(c) == 'so chung tu']
                if col_sct_calc:
                    phieu_series = clean_key_series(df_calc[col_sct_calc[0]])
                    col_sl_calc = [c for c in df_calc.columns if remove_accents(c) == 'so luong']
                    if col_sl_calc: dict_net = pd.to_numeric(df_calc[col_sl_calc[0]], errors='coerce').groupby(phieu_series).sum().to_dict()
                    if 'Số Gross' in df_calc.columns: dict_gross = pd.to_numeric(df_calc['Số Gross'], errors='coerce').groupby(phieu_series).sum().to_dict()

            final_cols = ['STT', 'Số phiếu cân', 'Bàn cân', 'Số xe', 'Giờ vào', 'Giờ ra', 'Tên hàng hóa', 'Khối lượng vào', 'Khối lượng ra', 'Khối lượng hàng hóa', 'Số chứng từ', '1', '2', '3', '4', '5', 'Khối lượng Net', 'Khối lượng Gross', '1-CỔNG TRƯỚC', '2-CỔNG TRƯỚC', '1-CỔNG SAU', '2-CỔNG SAU', 'KL bàn cân lớn nhất', 'Chênh lệch', 'Dung sai', 'Chênh lệch Max - Min', 'Đánh giá KSNB', 'Ghi chú']
            df_out = pd.DataFrame(columns=final_cols)
            
            def get_tc_main_col(col_names):
                for c in col_names:
                    if c in df_tc_main.columns: return df_tc_main[c]
                return ""

            df_out['STT'] = range(1, len(df_tc_main) + 1)
            df_out['Số phiếu cân'] = get_tc_main_col(['PHIẾU CÂN', 'Số phiếu cân'])
            df_out['Bàn cân'] = get_tc_main_col(['BÀN CÂN RA', 'CỔNG'])
            df_out['Số xe'] = get_tc_main_col(['BS RA 1', 'BS VÀO 1', 'Số xe'])
            df_out['Giờ vào'] = get_tc_main_col(['NGÀY GIỜ VÀO', 'Giờ vào'])
            df_out['Giờ ra'] = get_tc_main_col(['NGÀY GIỜ RA', 'Giờ ra'])
            df_out['Tên hàng hóa'] = get_tc_main_col(['TÊN HÀNG HÓA'])
            df_out['Khối lượng vào'] = pd.to_numeric(get_tc_main_col(['KL VÀO']), errors='coerce').fillna(0)
            df_out['Khối lượng ra'] = pd.to_numeric(get_tc_main_col(['KL RA']), errors='coerce').fillna(0)
            df_out['Khối lượng hàng hóa'] = pd.to_numeric(get_tc_main_col(['KHỐI LƯỢNG PHIẾU CÂN']), errors='coerce').fillna(0)
            
            if 'Lấy phiếu xuất.1' in df_tc_main.columns: df_out['Số chứng từ'] = df_tc_main['Lấy phiếu xuất.1']
            else: df_out['Số chứng từ'] = get_tc_main_col(['SỐ CHỨNG TỪ'])
            
            for i in range(1, 6): df_out[str(i)] = ""
                
            split_vouchers = df_out['Số chứng từ'].astype(str).replace('nan', '').str.replace('+', ',', regex=False).str.split(',', expand=True)
            for i in range(min(5, len(split_vouchers.columns))):
                df_out[str(i+1)] = split_vouchers[i].str.strip()

            df_out['Khối lượng Net'] = 0.0
            df_out['Khối lượng Gross'] = 0.0
            df_out['1-CỔNG TRƯỚC'] = np.nan
            df_out['2-CỔNG TRƯỚC'] = np.nan
            df_out['1-CỔNG SAU'] = np.nan
            df_out['2-CỔNG SAU'] = np.nan

            for idx, row in df_out.iterrows():
                net_sum = 0
                gross_sum = 0
                c1t, c2t, c1s, c2s = np.nan, np.nan, np.nan, np.nan
                
                for i in range(1, 6):
                    v = str(row[str(i)]).replace('.0', '').strip().upper()
                    if v and v != 'NAN':
                        net_sum += dict_net.get(v, 0)
                        gross_sum += dict_gross.get(v, 0)
                
                pc_key = str(row['Số phiếu cân']).strip().upper()
                if pc_key in dict_cong_can:
                    c1t = dict_cong_can[pc_key].get('1-CỔNG TRƯỚC', c1t)
                    c2t = dict_cong_can[pc_key].get('2-CỔNG TRƯỚC', c2t)
                    c1s = dict_cong_can[pc_key].get('1-CỔNG SAU', c1s)
                    c2s = dict_cong_can[pc_key].get('2-CỔNG SAU', c2s)

                df_out.at[idx, 'Khối lượng Net'] = net_sum
                df_out.at[idx, 'Khối lượng Gross'] = gross_sum
                df_out.at[idx, '1-CỔNG TRƯỚC'] = c1t
                df_out.at[idx, '2-CỔNG TRƯỚC'] = c2t
                df_out.at[idx, '1-CỔNG SAU'] = c1s
                df_out.at[idx, '2-CỔNG SAU'] = c2s

            cong_cols = ['1-CỔNG TRƯỚC', '2-CỔNG TRƯỚC', '1-CỔNG SAU', '2-CỔNG SAU']
            temp_cong = df_out[cong_cols].replace(0, np.nan)
            df_out['KL bàn cân lớn nhất'] = temp_cong.max(axis=1).fillna(df_out['Khối lượng hàng hóa'])
            df_out['Chênh lệch Max - Min'] = temp_cong.max(axis=1) - temp_cong.min(axis=1)

            df_out['Chênh lệch'] = df_out['Khối lượng Gross'] - df_out['KL bàn cân lớn nhất']
            df_out['Dung sai'] = np.where(df_out['Khối lượng Gross'] != 0, (df_out['Chênh lệch'] / df_out['Khối lượng Gross']).round(4), 0)
            df_out['Dung sai'] = df_out['Dung sai'].apply(lambda x: 'Vượt dung sai' if abs(x) > 0.003 else 'OK')
            
            def danh_gia_rui_ro(row):
                if row['Khối lượng Gross'] == 0 and row['Khối lượng Net'] == 0: return '🔴 Rủi ro (Thiếu dữ liệu ERP)'
                if pd.notna(row['Chênh lệch Max - Min']) and row['Chênh lệch Max - Min'] > 30: return '🔴 Rủi ro (Bất thường bàn cân)'
                if row['Dung sai'] == 'Vượt dung sai': return '🔴 Rủi ro Cao (Vượt dung sai)'
                return '🟢 An toàn'
                    
            df_out['Đánh giá KSNB'] = df_out.apply(danh_gia_rui_ro, axis=1)

            # --- BƯỚC 6: XUẤT FILE & AUTO FILTER ---
            progress_text.markdown('<div class="progress-text">⏳ Đang đóng gói & Cài đặt bộ lọc Excel tự động... (95%)</div>', unsafe_allow_html=True)
            progress_bar.progress(95)
            
            try:
                style_out = df_out.style.map(style_excel_cells, subset=['Dung sai', 'Đánh giá KSNB'])
                style_021 = df_021_export.style.map(style_excel_cells, subset=['Check ra cổng']) if not df_021_export.empty else df_021_export
            except AttributeError:
                style_out = df_out.style.applymap(style_excel_cells, subset=['Dung sai', 'Đánh giá KSNB'])
                style_021 = df_021_export.style.applymap(style_excel_cells, subset=['Check ra cổng']) if not df_021_export.empty else df_021_export
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                style_out.to_excel(writer, index=False, sheet_name="KSKTDulieucan-NM")
                ws1 = writer.sheets["KSKTDulieucan-NM"]
                ws1.freeze_panes = "A2"
                ws1.auto_filter.ref = ws1.dimensions
                
                if not df_021_export.empty:
                    style_021.to_excel(writer, index=False, sheet_name="021")
                    ws2 = writer.sheets["021"]
                    ws2.freeze_panes = "A2"
                    ws2.auto_filter.ref = ws2.dimensions

            progress_text.markdown('<div class="progress-text" style="color: #28a745;">✅ Đã hoàn tất! Các cổng cân đã được điền số liệu. File Excel đã tích hợp sẵn bộ lọc (Filter).</div>', unsafe_allow_html=True)
            progress_bar.progress(100)
            
            st.download_button(
                label="📥 Tải File Báo Cáo Chuẩn KSNB (Đã Tô Màu)",
                data=output.getvalue(),
                file_name="Bao_Cao_KSNB_Ra_Cong.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            progress_text.markdown(f'<div class="progress-text" style="color: #dc3545;">❌ Lỗi trong quá trình xử lý: {e}</div>', unsafe_allow_html=True)
