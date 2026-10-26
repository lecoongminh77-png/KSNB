import io
import pandas as pd
import numpy as np
import streamlit as st
import time

st.set_page_config(
    page_title="Kiểm Tra Hàng Ra Cổng - KSNB",
    page_icon="🚛",
    layout="wide"
)

# CSS Giao diện chuẩn Apple
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

# Thanh điều hướng
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Trang chủ", use_container_width=True):
        st.switch_page("app_ksnb.py")

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)

# Tiêu đề
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

# CÁC HÀM LỌC DỮ LIỆU ERP CHUẨN
def xuly_105_cungung(df):
    if df.empty or len(df.columns) < 4: return df
    remove_gd = ['COGS Recognition', 'WIP Return', 'WIP Issue', 'WIP Completion Return', 
                 'WIP Completion', 'WIP Byproduct Return', 'WIP Byproduct Completion', 
                 'Residual Qty Receipt', 'Residual Qty Issue']
    df = df[~df.iloc[:, 0].isin(remove_gd)]
    df = df[df.iloc[:, 0].notna()] 
    df = df[df.iloc[:, 3].astype(str).str.startswith(('3', '4'))]
    return df

def xuly_105_admin(df):
    if df.empty: return df
    df = df[~df.iloc[:, 0].isin(['COGS Recognition'])]
    return df

def xuly_021_cungung(df):
    if df.empty or len(df.columns) < 18: return df
    remove_gd = ['WIP Issue', 'WIP Completion Return', 'Return to vender', 'Return to vendor']
    df = df[~df.iloc[:, 17].isin(remove_gd)]
    return df

def xuly_021_admin(df):
    if df.empty or len(df.columns) < 18: return df
    df = df[~df.iloc[:, 17].isin(['Return to Vendor', 'Return to vender'])]
    return df

# XỬ LÝ CHÍNH
if st.button("🚀 Bắt đầu đối soát & Xuất Báo Cáo", use_container_width=True):
    if not f_tramcan_nm:
        st.warning("⚠️ Vui lòng nạp File Trạm Cân Nhà Máy!")
    else:
        # KHỞI TẠO THANH TIẾN TRÌNH
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # --- BƯỚC 1: ĐỌC DỮ LIỆU TRẠM CÂN ---
            progress_text.markdown('<div class="progress-text">⏳ Đang đọc dữ liệu Trạm Cân... (10%)</div>', unsafe_allow_html=True)
            progress_bar.progress(10)
            
            xls_tc = pd.ExcelFile(f_tramcan_nm)
            sheet_tc = 'tramcan' if 'tramcan' in xls_tc.sheet_names else 0
            df_tc = pd.read_excel(f_tramcan_nm, sheet_name=sheet_tc)
            
            # --- BƯỚC 2: ĐỌC VÀ LỌC ERP ---
            progress_text.markdown('<div class="progress-text">⏳ Đang xử lý bộ lọc cho các file 021 và 105... (35%)</div>', unsafe_allow_html=True)
            progress_bar.progress(35)
            
            # 105 (Chỉ dùng để tính toán, không xuất)
            frames_105 = []
            if f_105_cu: frames_105.append(xuly_105_cungung(pd.read_excel(f_105_cu, skiprows=7)))
            if f_105_admin: frames_105.append(xuly_105_admin(pd.read_excel(f_105_admin, skiprows=7)))
            df_105_all = pd.concat(frames_105, ignore_index=True) if frames_105 else pd.DataFrame()
            
            # 021 (Dùng tính toán + Xuất ra form)
            frames_021 = []
            if f_021_cu: frames_021.append(xuly_021_cungung(pd.read_excel(f_021_cu, skiprows=7)))
            if f_021_admin: frames_021.append(xuly_021_admin(pd.read_excel(f_021_admin, skiprows=7)))
            df_021_all = pd.concat(frames_021, ignore_index=True) if frames_021 else pd.DataFrame()

            # --- BƯỚC 3: TRÍCH XUẤT NET/GROSS ---
            progress_text.markdown('<div class="progress-text">⏳ Đang thu thập dữ liệu Net/Gross từ hệ thống ERP... (60%)</div>', unsafe_allow_html=True)
            progress_bar.progress(60)
            
            dict_net = {}
            dict_gross = {}
            
            # Gộp tạm 105 và 021 để tạo bộ từ điển dò tìm VLOOKUP
            df_erp = pd.concat([df_105_all, df_021_all], ignore_index=True) if not df_105_all.empty or not df_021_all.empty else pd.DataFrame()
            
            if not df_erp.empty:
                df_erp_search = df_erp.copy()
                df_erp_search.columns = df_erp_search.columns.astype(str).str.lower().str.strip()
                
                # Tìm động các cột quan trọng
                erp_col_phieu = next((c for c in df_erp_search.columns if "chứng từ" in c or "phiếu" in c or "tham chiếu" in c), None)
                erp_col_net = next((c for c in df_erp_search.columns if "net" in c or "khối lượng" in c), None)
                erp_col_gross = next((c for c in df_erp_search.columns if "gross" in c), None)
                
                if erp_col_phieu:
                    df_erp_search['phieu_clean'] = df_erp_search[erp_col_phieu].astype(str).str.replace('.0', '', regex=False).str.strip().str.upper()
                    if erp_col_net: dict_net = df_erp_search.groupby('phieu_clean')[erp_col_net].sum().to_dict()
                    if erp_col_gross: dict_gross = df_erp_search.groupby('phieu_clean')[erp_col_gross].sum().to_dict()

            # --- BƯỚC 4: ÉP FORM KSKTDulieucan-NM ---
            progress_text.markdown('<div class="progress-text">⏳ Đang ép khung dữ liệu và tính toán Chênh lệch, Dung sai... (80%)</div>', unsafe_allow_html=True)
            progress_bar.progress(80)
            
            # Cấu trúc 33 cột chuẩn y hệt file gốc
            final_cols = [
                'STT', 'Số phiếu cân', 'Bàn cân', 'Số xe', 'Giờ vào', 'Giờ ra', 'Tên hàng hóa', 
                'Khối lượng vào', 'Khối lượng ra', 'Khối lượng hàng hóa', 'Số chứng từ', 
                '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', 
                'Khối lượng Net', 'Khối lượng Gross', 
                '1-CỔNG TRƯỚC', '2-CỔNG TRƯỚC', '1-CỔNG SAU', '2-CỔNG SAU', 
                'KL bàn cân lớn nhất', 'Chênh lệch', 'Dung sai', 'Chênh lệch Max - Min', 
                'Check net -gross', 'Ghi chú'
            ]
            df_out = pd.DataFrame(columns=final_cols)
            
            # Hàm lấy an toàn từ file trạm cân thô
            def get_tc_col(col_names):
                for c in col_names:
                    if c in df_tc.columns: return df_tc[c]
                return ""

            df_out['STT'] = range(1, len(df_tc) + 1)
            df_out['Số phiếu cân'] = get_tc_col(['PHIẾU CÂN', 'Số phiếu cân'])
            df_out['Bàn cân'] = get_tc_col(['BÀN CÂN RA', 'CỔNG', 'Bàn cân'])
            df_out['Số xe'] = get_tc_col(['BS RA 1', 'BS VÀO 1', 'Số xe'])
            df_out['Giờ vào'] = get_tc_col(['NGÀY GIỜ VÀO', 'Giờ vào'])
            df_out['Giờ ra'] = get_tc_col(['NGÀY GIỜ RA', 'Giờ ra'])
            df_out['Tên hàng hóa'] = get_tc_col(['TÊN HÀNG HÓA', 'Tên hàng hóa'])
            df_out['Khối lượng vào'] = pd.to_numeric(get_tc_col(['KL VÀO']), errors='coerce').fillna(0)
            df_out['Khối lượng ra'] = pd.to_numeric(get_tc_col(['KL RA']), errors='coerce').fillna(0)
            df_out['Khối lượng hàng hóa'] = pd.to_numeric(get_tc_col(['KHỐI LƯỢNG PHIẾU CÂN', 'KL Hàng']), errors='coerce').fillna(0)
            df_out['Số chứng từ'] = get_tc_col(['Đưa số chứng từ về 1 cột', 'SỐ CHỨNG TỪ'])
            
            # Gán dữ liệu 10 cột tách (Trường hợp 1 xe nhiều chứng từ)
            for i in range(1, 11):
                col_str, col_int = str(i), i
                if col_int in df_tc.columns: df_out[col_str] = df_tc[col_int]
                elif col_str in df_tc.columns: df_out[col_str] = df_tc[col_str]

            df_out['Khối lượng Net'] = 0.0
            df_out['Khối lượng Gross'] = 0.0

            # Quét 10 cột, lấy Net/Gross từ Dict và CỘNG DỒN
            for i in range(1, 11):
                col_str = str(i)
                if col_str in df_out.columns:
                    vouchers = df_out[col_str].astype(str).str.replace('.0', '', regex=False).str.strip().str.upper()
                    df_out['Khối lượng Net'] += vouchers.map(dict_net).fillna(0)
                    df_out['Khối lượng Gross'] += vouchers.map(dict_gross).fillna(0)

            # Cổng cân
            cong_cols = ['1-CỔNG TRƯỚC', '2-CỔNG TRƯỚC', '1-CỔNG SAU', '2-CỔNG SAU']
            for c in cong_cols:
                if c in df_tc.columns: df_out[c] = pd.to_numeric(df_tc[c], errors='coerce')

            # Lấy Max Cổng cân và chênh lệch Max-Min
            existing_cong = [c for c in cong_cols if c in df_out.columns and df_out[c].notna().any()]
            if existing_cong:
                df_out['KL bàn cân lớn nhất'] = df_out[existing_cong].max(axis=1).fillna(df_out['Khối lượng hàng hóa'])
                temp_cong = df_out[existing_cong].replace(0, np.nan)
                df_out['Chênh lệch Max - Min'] = temp_cong.max(axis=1) - temp_cong.min(axis=1)
            else:
                df_out['KL bàn cân lớn nhất'] = df_out['Khối lượng hàng hóa']
                df_out['Chênh lệch Max - Min'] = np.nan

            # Tính Chênh lệch, Dung sai, Lỗi Net/Gross
            df_out['Chênh lệch'] = df_out['Khối lượng Gross'] - df_out['KL bàn cân lớn nhất']
            df_out['Dung sai'] = np.where(df_out['Khối lượng Gross'] != 0, (df_out['Chênh lệch'] / df_out['Khối lượng Gross']).round(4), 0)
            df_out['Dung sai'] = df_out['Dung sai'].apply(lambda x: 'Vượt dung sai' if abs(x) > 0.003 else 'OK')
            df_out['Check net -gross'] = np.where(df_out['Khối lượng Gross'] - df_out['Khối lượng Net'] > 0, "Hợp lệ", "Lỗi Gross/Net")

            # --- BƯỚC 5: TẠO FORM 021 & ĐÓNG GÓI EXCEL ---
            progress_text.markdown('<div class="progress-text">⏳ Đang tạo form 021 và đóng gói file báo cáo... (95%)</div>', unsafe_allow_html=True)
            progress_bar.progress(95)
            
            # Form 021
            if not df_021_all.empty:
                cols_021 = ['Check ra cổng', 'Thời gian pick SO', 'Người Pick SO', 'Ngày xuất', 'Thời gian thao tác', 'Người thực hiện', 'Số Gross', 'Chuỗi 2 ', 'Chuỗi 1', 'Chuyển text']
                df_021_export = pd.DataFrame(columns=cols_021 + df_021_all.columns.tolist())
                # Copy toàn bộ data đã lọc sang
                for c in df_021_all.columns: 
                    df_021_export[c] = df_021_all[c]
            else: 
                df_021_export = pd.DataFrame()

            # Ghi ra File
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_out.to_excel(writer, index=False, sheet_name="KSKTDulieucan-NM")
                if not df_021_export.empty:
                    df_021_export.to_excel(writer, index=False, sheet_name="021")

            progress_text.markdown('<div class="progress-text" style="color: #28a745;">✅ Đã hoàn tất ép form và đối soát dữ liệu!</div>', unsafe_allow_html=True)
            progress_bar.progress(100)
            
            # Nút tải
            st.download_button(
                label="📥 Tải File Báo Cáo Chuẩn KSNB",
                data=output.getvalue(),
                file_name="Bao_Cao_KSNB_Ra_Cong.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            progress_text.markdown(f'<div class="progress-text" style="color: #dc3545;">❌ Lỗi trong quá trình xử lý: {e}</div>', unsafe_allow_html=True)
