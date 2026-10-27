import io
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

# HÀM LỌC VÀ CHUẨN HÓA DỮ LIỆU
def xuly_105(df, is_cungung=True):
    try:
        header_idx = df.astype(str).apply(lambda row: row.str.contains('Loại Giao Dịch', case=False, na=False).any(), axis=1).idxmax()
        if pd.isna(header_idx): return pd.DataFrame()
        df.columns = df.iloc[header_idx]
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
        df = df.loc[:, df.columns.notna()]
        df = df.loc[:, ~df.columns.duplicated()]
        
        if df.empty or 'Loại Giao Dịch' not in df.columns: return df
        if is_cungung:
            remove_gd = ['COGS Recognition', 'WIP Return', 'WIP Issue', 'WIP Completion Return', 'WIP Completion', 'WIP Byproduct Return', 'WIP Byproduct Completion', 'Residual Qty Receipt', 'Residual Qty Issue']
            df = df[~df['Loại Giao Dịch'].isin(remove_gd)]
            if 'Mã Mặt Hàng' in df.columns: df = df[df['Mã Mặt Hàng'].astype(str).str.startswith(('3', '4'), na=False)]
        else:
            df = df[~df['Loại Giao Dịch'].isin(['COGS Recognition'])]
            
        df = df[df['Loại Giao Dịch'].notna()]
        if all(col in df.columns for col in ['Mă Kho', 'Khu Vực Kho', 'Locator', 'LPN', 'Loại Giao Dịch']):
            df['Chuỗi'] = df['Mă Kho'].astype(str) + df['Khu Vực Kho'].astype(str) + df['Locator'].astype(str) + df['LPN'].astype(str) + df['Loại Giao Dịch'].astype(str)
        else: df['Chuỗi'] = ""
        return df
    except: return pd.DataFrame()

def xuly_021(df, is_cungung=True):
    try:
        header_idx = df.astype(str).apply(lambda row: row.str.contains('Số chứng từ', case=False, na=False).any(), axis=1).idxmax()
        if pd.isna(header_idx): return pd.DataFrame()
        df.columns = df.iloc[header_idx]
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
        df = df.loc[:, df.columns.notna()]
        df = df.loc[:, ~df.columns.duplicated()]
        
        if df.empty or 'Loại giao dịch' not in df.columns: return df
        if is_cungung: remove_gd = ['WIP Issue', 'WIP Completion Return', 'Return to vender', 'Return to vendor']
        else: remove_gd = ['Return to Vendor', 'Return to vender']
        df = df[~df['Loại giao dịch'].isin(remove_gd)]
        return df
    except: return pd.DataFrame()

# HÀM TÔ MÀU EXCEL
def style_excel_cells(val):
    val_str = str(val).strip()
    if val_str in ['Vượt dung sai', 'Lỗi Gross/Net', '0'] or '🔴' in val_str:
        return 'background-color: #FFC7CE; color: #9C0006' 
    elif '🟡' in val_str:
        return 'background-color: #FFEB9C; color: #9C6500' 
    elif val_str in ['OK', 'Hợp lệ', '1'] or '🟢' in val_str:
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
            # --- BƯỚC 1: ĐỌC DỮ LIỆU TRẠM CÂN ---
            progress_text.markdown('<div class="progress-text">⏳ Đang đọc và làm sạch dữ liệu Trạm Cân... (10%)</div>', unsafe_allow_html=True)
            progress_bar.progress(10)
            xls_tc = pd.ExcelFile(f_tramcan_nm)
            sheet_tc = 'tramcan' if 'tramcan' in xls_tc.sheet_names else 0
            df_tc = pd.read_excel(f_tramcan_nm, sheet_name=sheet_tc)
            
            tc_chungtu_list = []
            if 'Lấy phiếu xuất.1' in df_tc.columns:
                tc_chungtu_list = [str(x).strip().upper() for x in df_tc['Lấy phiếu xuất.1'].tolist() if pd.notna(x)]
            elif 'SỐ CHỨNG TỪ' in df_tc.columns:
                tc_chungtu_list = [str(x).strip().upper() for x in df_tc['SỐ CHỨNG TỪ'].tolist() if pd.notna(x)]

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

            # --- TÍNH TOÁN SHEET 021 ---
            progress_text.markdown('<div class="progress-text">⏳ Đang tự động hóa VLOOKUP đa tầng cho bảng 021... (55%)</div>', unsafe_allow_html=True)
            progress_bar.progress(55)
            
            df_021_export = pd.DataFrame()
            if not df_021_all.empty:
                df_021_export = df_021_all.copy()
                def safe_str(col): return df_021_export[col].astype(str).replace('nan', '') if col in df_021_export.columns else ""
                
                df_021_export['Chuỗi 2 '] = safe_str('Mă kho') + safe_str('Subinventory') + safe_str('Locator') + safe_str('LPN') + safe_str('Loại giao dịch')
                df_021_export['Chuỗi Pick'] = safe_str('Mă kho') + safe_str('Subinventory') + safe_str('Locator') + safe_str('LPN') + "Sales Order Pick"

                if not df_105_all.empty and 'Chuỗi' in df_105_all.columns:
                    map_105 = df_105_all[df_105_all['Chuỗi'] != ""].drop_duplicates(subset=['Chuỗi'], keep='first').set_index('Chuỗi')
                    
                    df_021_export['Ngày xuất'] = df_021_export['Chuỗi 2 '].map(map_105['Ngày Giao Dịch'] if 'Ngày Giao Dịch' in map_105.columns else {})
                    df_021_export['Thời gian thao tác'] = df_021_export['Chuỗi 2 '].map(map_105['Creation Date'] if 'Creation Date' in map_105.columns else {})
                    df_021_export['Người thực hiện'] = df_021_export['Chuỗi 2 '].map(map_105['Created by User'] if 'Created by User' in map_105.columns else {})
                    df_021_export['Số Gross'] = df_021_export['Chuỗi 2 '].map(map_105['Khối lượng Gross'] if 'Khối lượng Gross' in map_105.columns else {})
                    
                    df_021_export['Thời gian pick SO'] = df_021_export['Chuỗi Pick'].map(map_105['Creation Date'] if 'Creation Date' in map_105.columns else {})
                    df_021_export['Người Pick SO'] = df_021_export['Chuỗi Pick'].map(map_105['Created by User'] if 'Created by User' in map_105.columns else {})
                else:
                    df_021_export['Ngày xuất'] = df_021_export['Thời gian thao tác'] = df_021_export['Người thực hiện'] = df_021_export['Số Gross'] = ""
                    df_021_export['Thời gian pick SO'] = df_021_export['Người Pick SO'] = ""

                df_021_export['Chuyển text'] = safe_str('Số chứng từ') + "A" 
                df_021_export['Chuỗi 1'] = df_021_export['Chuyển text'] + safe_str('Loại giao dịch')
                
                df_021_export['Check ra cổng'] = df_021_export['Số chứng từ'].apply(
                    lambda x: 1 if any(str(x).strip().upper() in str(tc_val) for tc_val in tc_chungtu_list) else 0
                )
                
                cols_10_first = ['Check ra cổng', 'Thời gian pick SO', 'Người Pick SO', 'Ngày xuất', 'Thời gian thao tác', 'Người thực hiện', 'Số Gross', 'Chuỗi 2 ', 'Chuỗi 1', 'Chuyển text']
                final_021_cols = cols_10_first + [c for c in df_021_all.columns if c not in cols_10_first and c != 'Chuỗi 2 ' and c != 'Chuỗi Pick']
                df_021_export = df_021_export[final_021_cols]

            # --- TÍNH TOÁN SHEET KSKT (FIX LỖI NET/GROSS) ---
            progress_text.markdown('<div class="progress-text">⏳ Đang lấy dữ liệu Net/Gross chuẩn từ 021 và tính dung sai... (75%)</div>', unsafe_allow_html=True)
            progress_bar.progress(75)
            
            dict_net, dict_gross = {}, {}
            
            # CHỈ DÙNG DF_021_EXPORT ĐỂ LẤY CẢ NET LẪN GROSS (Giống 100% file Excel gốc của bạn)
            if not df_021_export.empty:
                df_calc = df_021_export.copy()
                if 'Số chứng từ' in df_calc.columns:
                    phieu_series = df_calc['Số chứng từ'].astype(str).str.replace('.0', '', regex=False).str.strip().str.upper()
                    
                    if 'Số lượng' in df_calc.columns:
                        dict_net = pd.to_numeric(df_calc['Số lượng'], errors='coerce').groupby(phieu_series).sum().to_dict()
                    if 'Số Gross' in df_calc.columns:
                        dict_gross = pd.to_numeric(df_calc['Số Gross'], errors='coerce').groupby(phieu_series).sum().to_dict()

            final_cols = ['STT', 'Số phiếu cân', 'Bàn cân', 'Số xe', 'Giờ vào', 'Giờ ra', 'Tên hàng hóa', 'Khối lượng vào', 'Khối lượng ra', 'Khối lượng hàng hóa', 'Số chứng từ', '1', '2', '3', '4', '5', 'Khối lượng Net', 'Khối lượng Gross', '1-CỔNG TRƯỚC', '2-CỔNG TRƯỚC', '1-CỔNG SAU', '2-CỔNG SAU', 'KL bàn cân lớn nhất', 'Chênh lệch', 'Dung sai', 'Chênh lệch Max - Min', 'Check net -gross', 'Thời gian lưu bãi (Phút)', 'Đánh giá KSNB', 'Ghi chú']
            df_out = pd.DataFrame(columns=final_cols)
            
            def get_tc_col(col_names):
                for c in col_names:
                    if c in df_tc.columns: return df_tc[c]
                return ""

            df_out['STT'] = range(1, len(df_tc) + 1)
            df_out['Số phiếu cân'] = get_tc_col(['PHIẾU CÂN', 'Số phiếu cân'])
            df_out['Bàn cân'] = get_tc_col(['BÀN CÂN RA', 'CỔNG'])
            df_out['Số xe'] = get_tc_col(['BS RA 1', 'BS VÀO 1', 'Số xe'])
            df_out['Giờ vào'] = get_tc_col(['NGÀY GIỜ VÀO', 'Giờ vào'])
            df_out['Giờ ra'] = get_tc_col(['NGÀY GIỜ RA', 'Giờ ra'])
            df_out['Tên hàng hóa'] = get_tc_col(['TÊN HÀNG HÓA'])
            df_out['Khối lượng vào'] = pd.to_numeric(get_tc_col(['KL VÀO']), errors='coerce').fillna(0)
            df_out['Khối lượng ra'] = pd.to_numeric(get_tc_col(['KL RA']), errors='coerce').fillna(0)
            df_out['Khối lượng hàng hóa'] = pd.to_numeric(get_tc_col(['KHỐI LƯỢNG PHIẾU CÂN']), errors='coerce').fillna(0)
            
            if 'Lấy phiếu xuất.1' in df_tc.columns:
                df_out['Số chứng từ'] = df_tc['Lấy phiếu xuất.1']
            else:
                df_out['Số chứng từ'] = get_tc_col(['SỐ CHỨNG TỪ'])
            
            for i in range(1, 6): df_out[str(i)] = ""
                
            split_vouchers = df_out['Số chứng từ'].astype(str).replace('nan', '').str.replace('+', ',', regex=False).str.split(',', expand=True)
            for i in range(min(5, len(split_vouchers.columns))):
                df_out[str(i+1)] = split_vouchers[i].str.strip()

            df_out['Khối lượng Net'] = 0.0
            df_out['Khối lượng Gross'] = 0.0

            for i in range(1, 6):
                c = str(i)
                if c in df_out.columns:
                    vouchers = df_out[c].astype(str).str.replace('.0', '', regex=False).str.strip().str.upper()
                    df_out['Khối lượng Net'] += vouchers.map(dict_net).fillna(0)
                    df_out['Khối lượng Gross'] += vouchers.map(dict_gross).fillna(0)

            cong_cols = ['1-CỔNG TRƯỚC', '2-CỔNG TRƯỚC', '1-CỔNG SAU', '2-CỔNG SAU']
            for c in cong_cols:
                if c in df_tc.columns: df_out[c] = pd.to_numeric(df_tc[c], errors='coerce')

            existing_cong = [c for c in cong_cols if c in df_out.columns and df_out[c].notna().any()]
            if existing_cong:
                df_out['KL bàn cân lớn nhất'] = df_out[existing_cong].max(axis=1).fillna(df_out['Khối lượng hàng hóa'])
                temp_cong = df_out[existing_cong].replace(0, np.nan)
                df_out['Chênh lệch Max - Min'] = temp_cong.max(axis=1) - temp_cong.min(axis=1)
            else:
                df_out['KL bàn cân lớn nhất'] = df_out['Khối lượng hàng hóa']
                df_out['Chênh lệch Max - Min'] = np.nan

            df_out['Chênh lệch'] = df_out['Khối lượng Gross'] - df_out['KL bàn cân lớn nhất']
            df_out['Dung sai'] = np.where(df_out['Khối lượng Gross'] != 0, (df_out['Chênh lệch'] / df_out['Khối lượng Gross']).round(4), 0)
            df_out['Dung sai'] = df_out['Dung sai'].apply(lambda x: 'Vượt dung sai' if abs(x) > 0.003 else 'OK')
            df_out['Check net -gross'] = np.where(df_out['Khối lượng Gross'] - df_out['Khối lượng Net'] > 0, "Hợp lệ", "Lỗi Gross/Net")

            df_out['Giờ vào_dt'] = pd.to_datetime(df_out['Giờ vào'], errors='coerce')
            df_out['Giờ ra_dt'] = pd.to_datetime(df_out['Giờ ra'], errors='coerce')
            df_out['Thời gian lưu bãi (Phút)'] = (df_out['Giờ ra_dt'] - df_out['Giờ vào_dt']).dt.total_seconds() / 60
            df_out['Thời gian lưu bãi (Phút)'] = df_out['Thời gian lưu bãi (Phút)'].round(0)
            
            def danh_gia_rui_ro(row):
                if row['Check net -gross'] == 'Lỗi Gross/Net' or row['Dung sai'] == 'Vượt dung sai': return '🔴 Rủi ro Cao'
                elif pd.notna(row['Thời gian lưu bãi (Phút)']) and row['Thời gian lưu bãi (Phút)'] > 180: return '🟡 Cảnh báo (Lưu bãi lâu)'
                else: return '🟢 An toàn'
                    
            df_out['Đánh giá KSNB'] = df_out.apply(danh_gia_rui_ro, axis=1)
            df_out = df_out.drop(columns=['Giờ vào_dt', 'Giờ ra_dt'])

            # --- BƯỚC 6: ÁP DỤNG MÀU SẮC VÀ XUẤT FILE ---
            progress_text.markdown('<div class="progress-text">⏳ Đang tô màu và đóng gói File Báo Cáo... (95%)</div>', unsafe_allow_html=True)
            progress_bar.progress(95)
            
            try:
                style_out = df_out.style.map(style_excel_cells, subset=['Dung sai', 'Check net -gross', 'Đánh giá KSNB'])
                style_021 = df_021_export.style.map(style_excel_cells, subset=['Check ra cổng']) if not df_021_export.empty else df_021_export
            except AttributeError:
                style_out = df_out.style.applymap(style_excel_cells, subset=['Dung sai', 'Check net -gross', 'Đánh giá KSNB'])
                style_021 = df_021_export.style.applymap(style_excel_cells, subset=['Check ra cổng']) if not df_021_export.empty else df_021_export
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                style_out.to_excel(writer, index=False, sheet_name="KSKTDulieucan-NM")
                if not df_021_export.empty:
                    style_021.to_excel(writer, index=False, sheet_name="021")

            progress_text.markdown('<div class="progress-text" style="color: #28a745;">✅ Đã hoàn tất! Lỗi Gross/Net đã bị triệt tiêu hoàn toàn.</div>', unsafe_allow_html=True)
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
