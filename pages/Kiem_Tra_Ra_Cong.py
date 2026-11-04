import io
import unicodedata
import pandas as pd
import numpy as np
import streamlit as st
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="Kiểm Tra Hàng Ra Cổng - KSNB", page_icon="🚛", layout="wide")

# CSS Giao diện Apple Minimalist Pastel
st.markdown("""
<style>
    [data-testid="stSidebarNav"] { display: none !important; }
    .stApp { background-color: #f5f5f7; font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif; color: #1d1d1f; }
    .apple-header-box { display: flex; justify-content: space-between; align-items: flex-end; padding: 15px 0 10px 0; margin-bottom: 5px; }
    .apple-main-title { font-size: 36px; font-weight: 700; letter-spacing: -0.025em; margin: 0; color: #1d1d1f; }
    .apple-sub-text { font-size: 15px; font-weight: 600; text-align: right; margin: 0; line-height: 1.35; color: #424245; }
    .apple-sub-text span { color: #0071e3; font-size: 13px; display: block; margin-top: 2px; font-weight: 500; }
    div.stButton > button { background: #0071e3 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; padding: 10px 24px !important; font-weight: 600 !important; font-size: 15px !important; box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2) !important; transition: all 0.2s ease; }
    div.stButton > button:hover { background: #0077ED !important; transform: translateY(-1px); }
    .upload-box { background-color: #ffffff; padding: 16px; border-radius: 16px; box-shadow: 0 2px 12px rgba(0,0,0,0.03); margin-bottom: 15px; border: 1px solid #e5e5ea; }
    .section-title { font-size: 17px; font-weight: 600; margin-bottom: 10px; color: #1d1d1f; }
    .progress-text { font-size: 14px; font-weight: 600; color: #0071e3; margin-bottom: 8px; }
    
    /* KPI Cards Pastel */
    .kpi-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 14px; margin: 18px 0; }
    .kpi-card { background: #ffffff; border-radius: 18px; padding: 18px 20px; border: 1px solid #e5e5ea; box-shadow: 0 4px 16px rgba(0,0,0,0.03); }
    .kpi-title { font-size: 13px; font-weight: 600; color: #6e6e73; margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.02em; }
    .kpi-value { font-size: 28px; font-weight: 700; letter-spacing: -0.02em; margin-bottom: 4px; }
    .kpi-sub { font-size: 12px; font-weight: 500; color: #86868b; }
    
    .pastel-blue { background: linear-gradient(135deg, #ffffff 0%, #f0f7ff 100%); border-left: 5px solid #0071e3; }
    .pastel-green { background: linear-gradient(135deg, #ffffff 0%, #edfaf2 100%); border-left: 5px solid #28cd41; }
    .pastel-red { background: linear-gradient(135deg, #ffffff 0%, #fff0f2 100%); border-left: 5px solid #ff3b30; }
    .pastel-amber { background: linear-gradient(135deg, #ffffff 0%, #fff9eb 100%); border-left: 5px solid #ff9500; }
    .pastel-purple { background: linear-gradient(135deg, #ffffff 0%, #f7f2ff 100%); border-left: 5px solid #af52de; }
    
    .ratio-bar-wrap { background: #ffffff; padding: 16px 20px; border-radius: 16px; border: 1px solid #e5e5ea; margin-bottom: 20px; }
    .ratio-bar { display: flex; height: 12px; border-radius: 999px; overflow: hidden; background: #f2f2f7; margin-top: 8px; }
</style>
""", unsafe_allow_html=True)

col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Trang chủ", use_container_width=True):
        st.switch_page("app_ksnb.py")

st.markdown("""
<div class="apple-header-box">
    <div><h1 class="apple-main-title">🚛 Kiểm Tra Hàng Ra Cổng (Trạm Cân)</h1></div>
    <div><div class="apple-sub-text">Đối soát dung sai & Chênh lệch Net/Gross.<span>Lê Công Minh Workspace ›</span></div></div>
</div>
<hr style='border: none; border-top: 1px solid #e5e5ea; margin: 12px 0 20px 0;'>
""", unsafe_allow_html=True)

# KHU VỰC TẢI FILE
with st.expander("📂 Khu vực nạp dữ liệu Trạm Cân & ERP (Click để thu gọn/mở rộng)", expanded=True):
    st.markdown('<div class="section-title">📥 1. Dữ liệu Trạm Cân (Thực tế)</div>', unsafe_allow_html=True)
    f_tramcan_nm = st.file_uploader("File Trạm Cân (Nhà Máy)", type=["xlsx", "xls", "xlsb"], key="tc_nm")

    st.markdown('<div class="section-title" style="margin-top: 15px;">📥 2. Dữ liệu ERP (Lý thuyết)</div>', unsafe_allow_html=True)
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

# CÁC HÀM TIỆN ÍCH CHUẨN HÓA DỮ LIỆU
def remove_accents(input_str):
    s = str(input_str).strip().lower().replace('đ', 'd')
    nfkd_form = unicodedata.normalize('NFKD', s)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def clean_key_series(series):
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r'\.0$', '', regex=True)
        .str.upper()
        .replace(['NAN', 'NONE', '<NA>'], '', regex=False)
        .fillna('')
    )

# Hàm chuẩn hóa khóa Số chứng từ để gom nhóm chính xác (kể cả khi nhiều chứng từ nối dấu + khác thứ tự/khoảng trắng)
def normalize_ct_group_key(val):
    s = str(val).strip().upper()
    if not s or s in ['NAN', 'NONE', '<NA>']:
        return ''
    parts = [p.strip().replace('.0', '') for p in s.replace('+', ',').split(',') if p.strip()]
    if not parts:
        return ''
    return "+".join(sorted(parts))

def get_clean_col(df, col_name):
    target = remove_accents(col_name)
    matched = [c for c in df.columns if remove_accents(c) == target]
    if matched:
        return clean_key_series(df[matched[0]])
    return pd.Series('', index=df.index)

@st.cache_data(show_spinner=False)
def read_and_process_105(file_bytes, is_cungung=True):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str)
        header_idx = df.apply(lambda row: row.astype(str).str.contains('Loại Giao Dịch', case=False, na=False).any(), axis=1).idxmax()
        if pd.isna(header_idx): return pd.DataFrame()
        df.columns = df.iloc[header_idx].astype(str).str.strip()
        
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
        if not df.empty:
            first_row_vals = df.iloc[0].astype(str).str.strip().values
            if any(val.startswith('[') and val.endswith(']') for val in first_row_vals):
                df = df.iloc[1:].reset_index(drop=True)
        
        gross_cols = [i for i, c in enumerate(df.columns) if 'gross' in remove_accents(c)]
        if gross_cols:
            raw_gross = df.iloc[:, gross_cols[0]]
        elif df.shape[1] > 39:
            raw_gross = df.iloc[:, 39]
        else:
            raw_gross = pd.Series(0, index=df.index)
            
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
        
        lot_or_lpn = pd.Series(np.where(ma_lo != '', ma_lo, lpn), index=df.index)
        
        df['Chuỗi'] = ma_kho + kv_kho + locator + lot_or_lpn + loai_gd
        df['Chuỗi_MaLo'] = ma_kho + kv_kho + locator + ma_lo + loai_gd
        df['Chuỗi_Kho_Lo_GD'] = ma_kho + kv_kho + lot_or_lpn + loai_gd
        df['Chuỗi_KhongLo'] = ma_kho + kv_kho + locator + loai_gd
        
        df['Chuỗi_Pick'] = ma_kho + kv_kho + locator + lot_or_lpn + "SALES ORDER PICK"
        df['Chuỗi_Pick_MaLo'] = ma_kho + kv_kho + locator + ma_lo + "SALES ORDER PICK"
        return df
    except:
        return pd.DataFrame()

@st.cache_data(show_spinner=False)
def read_and_process_021(file_bytes, is_cungung=True):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str)
        header_idx = df.apply(lambda row: row.astype(str).str.contains('Số chứng từ', case=False, na=False).any(), axis=1).idxmax()
        if pd.isna(header_idx): return pd.DataFrame()
        df.columns = df.iloc[header_idx].astype(str).str.strip()
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
        
        if not df.empty:
            first_row_vals = df.iloc[0].astype(str).str.strip().values
            if any(val.startswith('[') and val.endswith(']') for val in first_row_vals):
                df = df.iloc[1:].reset_index(drop=True)

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
    except:
        return pd.DataFrame()

# XỬ LÝ CHÍNH
if st.button("🚀 Bắt đầu đối soát & Hiển thị Dashboard", use_container_width=True):
    if not f_tramcan_nm:
        st.warning("⚠️ Vui lòng nạp File Trạm Cân Nhà Máy!")
    else:
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # --- BƯỚC 1: ĐỌC DỮ LIỆU TRẠM CÂN & LỌC DUY NHẤT PHIẾU CÂN LỚN NHẤT CHO MỖI CHỨNG TỪ ---
            progress_text.markdown('<div class="progress-text">⏳ Đang lọc lấy Phiếu cân có khối lượng lớn nhất theo Số chứng từ... (15%)</div>', unsafe_allow_html=True)
            progress_bar.progress(15)
            
            xls_tc = pd.ExcelFile(f_tramcan_nm)
            sheet_tc = 'tramcan' if 'tramcan' in xls_tc.sheet_names else 0
            df_tc_raw = pd.read_excel(f_tramcan_nm, sheet_name=sheet_tc)
            df_tc_full = df_tc_raw.dropna(subset=['PHIẾU CÂN']).copy()
            
            df_tc_full['PC_CLEAN'] = clean_key_series(df_tc_full['PHIẾU CÂN'])
            df_tc_full['CONG_CLEAN'] = clean_key_series(df_tc_full.get('BÀN CÂN RA', pd.Series('', index=df_tc_full.index)))
            df_tc_full['KL_TAM'] = pd.to_numeric(df_tc_full.get('KHỐI LƯỢNG PHIẾU CÂN', 0), errors='coerce').fillna(0)
            
            # Quét 4 bàn cân theo từng Phiếu cân
            pivot_cong = df_tc_full.pivot_table(
                index='PC_CLEAN', columns='CONG_CLEAN', values='KL_TAM', aggfunc='max'
            )
            # Tính số cân lớn nhất của từng Phiếu cân để so sánh khi trùng Số chứng từ
            pc_max_kl = df_tc_full.groupby('PC_CLEAN')['KL_TAM'].max().to_dict()
            df_tc_full['MAX_KL_OF_PC'] = df_tc_full['PC_CLEAN'].map(pc_max_kl).fillna(df_tc_full['KL_TAM'])

            col_chungtu = 'Lấy phiếu xuất.1' if 'Lấy phiếu xuất.1' in df_tc_full.columns else 'SỐ CHỨNG TỪ'
            df_tc_full['CT_KEY'] = df_tc_full[col_chungtu].apply(normalize_ct_group_key)
            
            # 1. Với các phiếu có Số chứng từ: Gom theo Số chứng từ (CT_KEY), sắp xếp giảm dần theo số cân lớn nhất và CHỈ GIỮ LẠI 1 PHIẾU CÂN LỚN NHẤT
            df_has_ct = df_tc_full[df_tc_full['CT_KEY'] != ''].copy()
            df_has_ct = (
                df_has_ct.sort_values(by=['MAX_KL_OF_PC', 'KL_TAM'], ascending=[False, False])
                .drop_duplicates(subset=['CT_KEY'], keep='first')
            )
            
            # 2. Với các phiếu không có Số chứng từ: Chỉ giữ lại dòng có khối lượng lớn nhất theo từng Mã phiếu cân
            df_no_ct = df_tc_full[df_tc_full['CT_KEY'] == ''].copy()
            df_no_ct = (
                df_no_ct.sort_values(by='KL_TAM', ascending=False)
                .drop_duplicates(subset=['PC_CLEAN'], keep='first')
            )
            
            # Gộp lại và đảm bảo mỗi Phiếu cân chỉ xuất hiện 1 dòng duy nhất
            df_tc_main = pd.concat([df_has_ct, df_no_ct], ignore_index=True)
            df_tc_main = (
                df_tc_main.sort_values(by='MAX_KL_OF_PC', ascending=False)
                .drop_duplicates(subset=['PC_CLEAN'], keep='first')
                .reset_index(drop=True)
            )
            
            # Tập hợp các chứng từ đơn lẻ trên trạm cân để đánh dấu Check ra cổng O(1)
            all_tc_vouchers_set = set(
                df_tc_main[col_chungtu].dropna().astype(str)
                .str.replace('+', ',', regex=False).str.split(',').explode()
                .str.strip().str.upper()
            ) - {'', 'NAN'}

            # --- BƯỚC 2: ĐỌC VÀ LỌC ERP ---
            progress_text.markdown('<div class="progress-text">⏳ Đang đọc và chuẩn hóa dữ liệu ERP 105 & 021... (40%)</div>', unsafe_allow_html=True)
            progress_bar.progress(40)
            
            frames_105, frames_021 = [], []
            if f_105_cu: frames_105.append(read_and_process_105(f_105_cu.getvalue(), True))
            if f_105_admin: frames_105.append(read_and_process_105(f_105_admin.getvalue(), False))
            df_105_all = pd.concat(frames_105, ignore_index=True) if frames_105 else pd.DataFrame()
            
            if f_021_cu: frames_021.append(read_and_process_021(f_021_cu.getvalue(), True))
            if f_021_admin: frames_021.append(read_and_process_021(f_021_admin.getvalue(), False))
            df_021_all = pd.concat(frames_021, ignore_index=True) if frames_021 else pd.DataFrame()

            # --- BƯỚC 3: ÁNH XẠ ĐA TẦNG 021 SANG 105 ---
            progress_text.markdown('<div class="progress-text">⏳ Đang thực hiện VLOOKUP đa tầng tốc độ cao... (65%)</div>', unsafe_allow_html=True)
            progress_bar.progress(65)
            
            df_021_export = pd.DataFrame()
            if not df_021_all.empty:
                df_021_export = df_021_all.copy()
                
                ma_kho_021 = get_clean_col(df_021_export, 'Mã kho')
                sub_021 = get_clean_col(df_021_export, 'Subinventory')
                loc_021 = get_clean_col(df_021_export, 'Locator')
                lpn_021 = get_clean_col(df_021_export, 'LPN')
                malo_021 = get_clean_col(df_021_export, 'Mã lô')
                loai_gd_021 = get_clean_col(df_021_export, 'Loại giao dịch')
                sct_021_clean = get_clean_col(df_021_export, 'Số chứng từ')
                
                lot_or_lpn_021 = pd.Series(np.where(malo_021 != '', malo_021, lpn_021), index=df_021_export.index)

                df_021_export['Chuỗi 2 '] = ma_kho_021 + sub_021 + loc_021 + lot_or_lpn_021 + loai_gd_021
                s_021_malo = ma_kho_021 + sub_021 + loc_021 + malo_021 + loai_gd_021
                s_021_kho_lo_gd = ma_kho_021 + sub_021 + lot_or_lpn_021 + loai_gd_021
                s_021_nolo = ma_kho_021 + sub_021 + loc_021 + loai_gd_021
                
                df_021_export['Chuỗi Pick'] = ma_kho_021 + sub_021 + loc_021 + lot_or_lpn_021 + "SALES ORDER PICK"
                s_pick_malo = ma_kho_021 + sub_021 + loc_021 + malo_021 + "SALES ORDER PICK"

                if not df_105_all.empty and 'Chuỗi' in df_105_all.columns:
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

                df_021_export['Chuyển text'] = sct_021_clean + "A" 
                df_021_export['Chuỗi 1'] = df_021_export['Chuyển text'] + loai_gd_021
                df_021_export['Check ra cổng'] = np.where(sct_021_clean.isin(all_tc_vouchers_set), 1, 0)
                
                cols_10_first = ['Check ra cổng', 'Thời gian pick SO', 'Người Pick SO', 'Ngày xuất', 'Thời gian thao tác', 'Người thực hiện', 'Số Gross', 'Chuỗi 2 ', 'Chuỗi 1', 'Chuyển text']
                final_021_cols = cols_10_first + [c for c in df_021_all.columns if c not in cols_10_first and c not in ['Chuỗi 2 ', 'Chuỗi Pick']]
                df_021_export = df_021_export[final_021_cols]

            # --- BƯỚC 4: TÍNH TOÁN BẢNG TỔNG HỢP KSKT ---
            progress_text.markdown('<div class="progress-text">⏳ Đang đối soát dung sai với Phiếu cân lớn nhất... (85%)</div>', unsafe_allow_html=True)
            progress_bar.progress(85)
            
            dict_net, dict_gross, dict_xe_021 = {}, {}, {}
            if not df_021_export.empty:
                col_sct_calc = [c for c in df_021_export.columns if remove_accents(c) == 'so chung tu']
                col_sl_calc = [c for c in df_021_export.columns if remove_accents(c) == 'so luong']
                col_tc_calc = [c for c in df_021_export.columns if remove_accents(c) == 'so tham chieu']
                
                if col_sct_calc:
                    phieu_series = clean_key_series(df_021_export[col_sct_calc[0]])
                    if col_sl_calc:
                        sl_num = pd.to_numeric(df_021_export[col_sl_calc[0]].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)
                        dict_net = sl_num.groupby(phieu_series).sum().to_dict()
                    if 'Số Gross' in df_021_export.columns:
                        dict_gross = pd.to_numeric(df_021_export['Số Gross'], errors='coerce').fillna(0).groupby(phieu_series).sum().to_dict()
                    if col_tc_calc:
                        df_034 = df_021_export[phieu_series.str.startswith('PX034')].copy()
                        if not df_034.empty:
                            p_034 = clean_key_series(df_034[col_sct_calc[0]])
                            xe_034 = clean_key_series(df_034[col_tc_calc[0]])
                            dict_xe_021 = xe_034[xe_034 != ''].groupby(p_034).nunique().to_dict()

            final_cols = ['STT', 'Số phiếu cân', 'Bàn cân', 'Số xe', 'Giờ vào', 'Giờ ra', 'Tên hàng hóa', 'Khối lượng vào', 'Khối lượng ra', 'Khối lượng hàng hóa', 'Số chứng từ', '1', '2', '3', '4', '5', 'Khối lượng Net', 'Khối lượng Gross', '1-CỔNG TRƯỚC', '2-CỔNG TRƯỚC', '1-CỔNG SAU', '2-CỔNG SAU', 'KL bàn cân lớn nhất', 'Chênh lệch', 'Dung sai', 'Chênh lệch Max - Min', 'Đánh giá KSNB', 'Ghi chú']
            df_out = pd.DataFrame(index=df_tc_main.index, columns=final_cols)
            
            def get_tc_main_col(col_names):
                for c in col_names:
                    if c in df_tc_main.columns: return df_tc_main[c]
                return pd.Series("", index=df_tc_main.index)

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
            df_out['Số chứng từ'] = df_tc_main[col_chungtu]
            
            split_vouchers = df_out['Số chứng từ'].astype(str).replace('nan', '').str.replace('+', ',', regex=False).str.split(',', expand=True)
            net_total = pd.Series(0.0, index=df_out.index)
            gross_total = pd.Series(0.0, index=df_out.index)
            
            for i in range(1, 6):
                if i - 1 < split_vouchers.shape[1]:
                    col_v = clean_key_series(split_vouchers[i - 1])
                    df_out[str(i)] = col_v
                    net_total += col_v.map(dict_net).fillna(0)
                    gross_total += col_v.map(dict_gross).fillna(0)
                else:
                    df_out[str(i)] = ""

            df_out['Khối lượng Net'] = net_total
            df_out['Khối lượng Gross'] = gross_total
            
            pc_keys = clean_key_series(df_out['Số phiếu cân'])
            for cong_name in ['1-CỔNG TRƯỚC', '2-CỔNG TRƯỚC', '1-CỔNG SAU', '2-CỔNG SAU']:
                if cong_name in pivot_cong.columns:
                    df_out[cong_name] = pc_keys.map(pivot_cong[cong_name])
                else:
                    df_out[cong_name] = np.nan

            cong_cols = ['1-CỔNG TRƯỚC', '2-CỔNG TRƯỚC', '1-CỔNG SAU', '2-CỔNG SAU']
            temp_cong = df_out[cong_cols].replace(0, np.nan)
            
            # Lấy số cân lớn nhất của Phiếu cân đại diện đó làm chuẩn để so sánh với Gross ERP
            df_out['KL bàn cân lớn nhất'] = np.maximum(
                temp_cong.max(axis=1).fillna(0),
                df_out['Khối lượng hàng hóa']
            )
            df_out['Chênh lệch Max - Min'] = temp_cong.max(axis=1) - temp_cong.min(axis=1)

            df_out['Chênh lệch'] = df_out['Khối lượng Gross'] - df_out['KL bàn cân lớn nhất']
            dung_sai_num = np.where(df_out['Khối lượng Gross'] != 0, (df_out['Chênh lệch'] / df_out['Khối lượng Gross']).round(4), 0)
            df_out['Dung sai'] = np.where(np.abs(dung_sai_num) > 0.003, 'Vượt dung sai', 'OK')
            
            def phan_tich_ghi_chu(row):
                notes = []
                raw_sct = str(row['Số chứng từ']).strip()
                if '+' in raw_sct or ',' in raw_sct:
                    parts = [p.strip().upper() for p in raw_sct.replace('+', ',').split(',') if p.strip()]
                    if len(parts) > len(set(parts)):
                        notes.append("⚠️ Lỗi bảo vệ: Nhập lặp chứng từ")
                for i in range(1, 6):
                    v = str(row[str(i)]).strip().upper()
                    if dict_xe_021.get(v, 0) > 1:
                        notes.append(f"⚠️ Lỗi kho 034: Chứng từ {v} lẫn nhiều biển số xe")
                    if v.startswith('PXHAB') and row['Khối lượng Net'] > 0 and row['Dung sai'] == 'Vượt dung sai':
                        notes.append(f"ℹ️ Dính kho HAB ({v} thiếu Gross trên 105)")
                return " | ".join(dict.fromkeys(notes))

            df_out['Ghi chú'] = df_out.apply(phan_tich_ghi_chu, axis=1)

            cond_missing = (df_out['Khối lượng Gross'] == 0) & (df_out['Khối lượng Net'] == 0)
            cond_scale_err = df_out['Chênh lệch Max - Min'].notna() & (df_out['Chênh lệch Max - Min'] > 30)
            cond_tol_err = df_out['Dung sai'] == 'Vượt dung sai'
            
            df_out['Đánh giá KSNB'] = np.select(
                [cond_missing, cond_scale_err, cond_tol_err],
                ['🔴 Rủi ro (Thiếu dữ liệu ERP)', '🔴 Rủi ro (Bất thường bàn cân)', '🔴 Rủi ro Cao (Vượt dung sai)'],
                default='🟢 An toàn'
            )

            st.session_state['df_out'] = df_out
            st.session_state['df_021_export'] = df_021_export

            # --- BƯỚC 5: ĐÓNG GÓI EXCEL TỐC ĐỘ CAO BẰNG OPENPYXL ---
            progress_text.markdown('<div class="progress-text">⏳ Đang đóng gói Excel tốc độ cao... (95%)</div>', unsafe_allow_html=True)
            progress_bar.progress(95)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_out.to_excel(writer, index=False, sheet_name="KSKTDulieucan-NM")
                ws1 = writer.sheets["KSKTDulieucan-NM"]
                ws1.freeze_panes = "A2"
                ws1.auto_filter.ref = ws1.dimensions
                
                fill_red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                font_red = Font(color="9C0006")
                fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                font_green = Font(color="006100")
                
                for row_idx in range(2, len(df_out) + 2):
                    for col_idx in [25, 27]:
                        cell = ws1.cell(row=row_idx, column=col_idx)
                        val_str = str(cell.value or '')
                        if '🔴' in val_str or val_str == 'Vượt dung sai':
                            cell.fill = fill_red
                            cell.font = font_red
                        elif '🟢' in val_str or val_str == 'OK':
                            cell.fill = fill_green
                            cell.font = font_green

                if not df_021_export.empty:
                    df_021_export.to_excel(writer, index=False, sheet_name="021")
                    ws2 = writer.sheets["021"]
                    ws2.freeze_panes = "A2"
                    ws2.auto_filter.ref = ws2.dimensions

            st.session_state['excel_bytes'] = output.getvalue()
            progress_text.markdown('<div class="progress-text" style="color: #28a745;">✅ Hoàn tất đối soát! Dữ liệu đã sẵn sàng hiển thị bên dưới.</div>', unsafe_allow_html=True)
            progress_bar.progress(100)

        except Exception as e:
            progress_text.markdown(f'<div class="progress-text" style="color: #dc3545;">❌ Lỗi trong quá trình xử lý: {e}</div>', unsafe_allow_html=True)

# HIỂN THỊ KẾT QUẢ TRỰC QUAN TRÊN WEB (DASHBOARD)
if 'df_out' in st.session_state:
    df_res = st.session_state['df_out']
    total_cnt = len(df_res)
    safe_cnt = int((df_res['Đánh giá KSNB'] == '🟢 An toàn').sum())
    vuot_cnt = int((df_res['Dung sai'] == 'Vượt dung sai').sum())
    missing_cnt = int((df_res['Đánh giá KSNB'] == '🔴 Rủi ro (Thiếu dữ liệu ERP)').sum())
    ops_err_cnt = int((df_res['Ghi chú'] != '').sum())

    safe_pct = round(safe_cnt / total_cnt * 100, 1) if total_cnt else 0
    vuot_pct = round(vuot_cnt / total_cnt * 100, 1) if total_cnt else 0
    missing_pct = round(missing_cnt / total_cnt * 100, 1) if total_cnt else 0

    st.markdown("### 📊 Bảng Điều Khiển Kết Quả Đối Soát (KSNB Dashboard)")
    
    st.markdown(f"""
    <div class="kpi-grid">
        <div class="kpi-card pastel-blue">
            <div class="kpi-title">Tổng Phiếu Cân</div>
            <div class="kpi-value" style="color: #0071e3;">{total_cnt:,}</div>
            <div class="kpi-sub">Đã lọc giữ phiếu lớn nhất</div>
        </div>
        <div class="kpi-card pastel-green">
            <div class="kpi-title">🟢 An Toàn (Chuẩn)</div>
            <div class="kpi-value" style="color: #1e8e3e;">{safe_cnt:,}</div>
            <div class="kpi-sub">Chiếm {safe_pct}% tổng phiếu</div>
        </div>
        <div class="kpi-card pastel-red">
            <div class="kpi-title">🔴 Vượt Dung Sai (>0.3%)</div>
            <div class="kpi-value" style="color: #d93025;">{vuot_cnt:,}</div>
            <div class="kpi-sub">Chiếm {vuot_pct}% tổng phiếu</div>
        </div>
        <div class="kpi-card pastel-amber">
            <div class="kpi-title">🟡 Thiếu Dữ Liệu ERP</div>
            <div class="kpi-value" style="color: #e37400;">{missing_cnt:,}</div>
            <div class="kpi-sub">Chiếm {missing_pct}% (Net & Gross = 0)</div>
        </div>
        <div class="kpi-card pastel-purple">
            <div class="kpi-title">🟣 Cảnh Báo Vận Hành</div>
            <div class="kpi-value" style="color: #8e24aa;">{ops_err_cnt:,}</div>
            <div class="kpi-sub">Lỗi bảo vệ / Lẫn xe / Kho HAB</div>
        </div>
    </div>
    
    <div class="ratio-bar-wrap">
        <div style="display:flex; justify-content:space-between; font-size:13px; font-weight:600; color:#424245;">
            <span>Tỷ trọng phân bổ trạng thái phiếu cân</span>
            <span>🟢 An toàn: {safe_pct}%  |  🔴 Vượt dung sai: {vuot_pct}%  |  🟡 Thiếu ERP: {missing_pct}%</span>
        </div>
        <div class="ratio-bar">
            <div style="width:{safe_pct}%; background:#34c759;"></div>
            <div style="width:{vuot_pct}%; background:#ff3b30;"></div>
            <div style="width:{missing_pct}%; background:#ff9500;"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if 'excel_bytes' in st.session_state:
        st.download_button(
            label="📥 Tải Xuống File Báo Cáo Chuẩn KSNB (.xlsx - Đã Tô Màu & Auto Filter)",
            data=st.session_state['excel_bytes'],
            file_name="Bao_Cao_KSNB_Ra_Cong.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.markdown("#### 🔍 Tra cứu & Soi lỗi trực tiếp trên Web")
    search_kw = st.text_input("🔎 Tìm nhanh theo Số phiếu cân, Biển số xe hoặc Số chứng từ:", placeholder="Nhập biển số xe (VD: 50E89471) hoặc mã chứng từ...")
    
    df_view = df_res.copy()
    if search_kw.strip():
        kw = search_kw.strip().upper()
        mask_kw = (
            df_view['Số phiếu cân'].astype(str).str.upper().str.contains(kw, na=False) |
            df_view['Số xe'].astype(str).str.upper().str.contains(kw, na=False) |
            df_view['Số chứng từ'].astype(str).str.upper().str.contains(kw, na=False)
        )
        df_view = df_view[mask_kw]

    tab1, tab2, tab3, tab4 = st.tabs([
        f"🔴 Vượt Dung Sai ({len(df_view[df_view['Dung sai'] == 'Vượt dung sai'])})",
        f"🟣 Cảnh Báo Bảo Vệ / Kho ({len(df_view[df_view['Ghi chú'] != ''])})",
        f"🟡 Thiếu Dữ Liệu ERP ({len(df_view[df_view['Đánh giá KSNB'].str.contains('Thiếu dữ liệu ERP', na=False)])})",
        f"📋 Tất Cả Phiếu Cân ({len(df_view)})"
    ])

    cols_show = ['STT', 'Số phiếu cân', 'Bàn cân', 'Số xe', 'Số chứng từ', 'Khối lượng Net', 'Khối lượng Gross', 'KL bàn cân lớn nhất', 'Chênh lệch', 'Dung sai', 'Đánh giá KSNB', 'Ghi chú']

    with tab1:
        st.dataframe(df_view[df_view['Dung sai'] == 'Vượt dung sai'][cols_show], use_container_width=True, height=420)
    with tab2:
        st.dataframe(df_view[df_view['Ghi chú'] != ''][cols_show], use_container_width=True, height=420)
    with tab3:
        st.dataframe(df_view[df_view['Đánh giá KSNB'].str.contains('Thiếu dữ liệu ERP', na=False)][cols_show], use_container_width=True, height=420)
    with tab4:
        st.dataframe(df_view[cols_show], use_container_width=True, height=420)
