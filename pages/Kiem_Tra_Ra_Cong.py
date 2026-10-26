import io
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(
    page_title="Kiểm Tra Hàng Ra Cổng - KSNB",
    page_icon="🚛",
    layout="wide"
)

# CSS giao diện chuẩn Apple
st.markdown("""
<style>
    [data-testid="stSidebarNav"] { display: none !important; }
    .stApp {
        background-color: #f5f5f7;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif;
        color: #1d1d1f;
    }
    .apple-header-box {
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
        padding: 20px 0 10px 0;
        margin-bottom: 5px;
    }
    .apple-main-title {
        font-size: 38px;
        font-weight: 700;
        letter-spacing: -0.025em;
        margin: 0;
    }
    .apple-sub-text {
        font-size: 16px;
        font-weight: 600;
        text-align: right;
        margin: 0;
        line-height: 1.35;
    }
    .apple-sub-text span {
        color: #0071e3;
        font-size: 14px;
        display: block;
        margin-top: 2px;
        font-weight: 500;
    }
    div.stButton > button {
        background: #0071e3 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding: 10px 24px !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2) !important;
    }
    .upload-box {
        background-color: white;
        padding: 15px;
        border-radius: 12px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.03);
        margin-bottom: 15px;
        border: 1px solid #e2e8f0;
    }
    .section-title { font-size: 18px; font-weight: 600; margin-bottom: 10px; color: #1d1d1f; }
</style>
""", unsafe_allow_html=True)

# Thanh điều hướng
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button("🏠 Trang chủ", use_container_width=True):
        st.switch_page("app_ksnb.py")

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)

# Tiêu đề trang
st.markdown("""
<div class="apple-header-box">
    <div>
        <h1 class="apple-main-title">🚛 Kiểm Tra Hàng Ra Cổng (Trạm Cân)</h1>
    </div>
    <div>
        <div class="apple-sub-text">
            Đối soát dung sai & Chênh lệch Net/Gross.
            <span>Lê Công Minh Workspace ›</span>
        </div>
    </div>
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

# HÀM LỌC DỮ LIỆU CHUẨN
def xuly_105_cungung(df):
    if df.empty: return df
    # Cột A (index 0): Loại bỏ các GD không cần thiết + bỏ trống
    remove_giao_dich = ['COGS Recognition', 'WIP Return', 'WIP Issue', 'WIP Completion Return', 
                        'WIP Completion', 'WIP Byproduct Return', 'WIP Byproduct Completion', 
                        'Residual Qty Receipt', 'Residual Qty Issue']
    df = df[~df.iloc[:, 0].isin(remove_giao_dich)]
    df = df[df.iloc[:, 0].notna()] # Bỏ Blanks
    
    # Cột D (index 3): Chỉ giữ sản phẩm đầu 3 và 4
    df = df[df.iloc[:, 3].astype(str).str.startswith(('3', '4'))]
    return df

def xuly_105_admin(df):
    if df.empty: return df
    # Cột A (index 0): Bỏ COGS Recognition
    df = df[~df.iloc[:, 0].isin(['COGS Recognition'])]
    return df

def xuly_021_cungung(df):
    if df.empty: return df
    # Cột R (index 17): Bỏ WIP Issue, WIP Completion Return, Return to vender/vendor
    remove_gd = ['WIP Issue', 'WIP Completion Return', 'Return to vender', 'Return to vendor']
    df = df[~df.iloc[:, 17].isin(remove_gd)]
    return df

def xuly_021_admin(df):
    if df.empty: return df
    # Cột R (index 17): Bỏ Return to Vendor
    df = df[~df.iloc[:, 17].isin(['Return to Vendor', 'Return to vender'])]
    return df

# NÚT XỬ LÝ
if st.button("🚀 Xử lý Lọc dữ liệu & Bắt đầu đối soát", use_container_width=True):
    if not f_tramcan_nm:
        st.warning("⚠️ Vui lòng nạp File Trạm Cân Nhà Máy!")
    else:
        with st.spinner("Đang áp dụng bộ lọc và tính toán chênh lệch..."):
            try:
                # 1. Đọc Trạm Cân
                df_tc = pd.read_excel(f_tramcan_nm)
                
                # 2. Đọc và lọc ERP (Giả sử ERP xuất ra bị dư 7 dòng đầu tiên - skiprows=7)
                erp_frames = []
                
                if f_105_cu:
                    df_temp = pd.read_excel(f_105_cu, skiprows=7)
                    df_temp = xuly_105_cungung(df_temp)
                    erp_frames.append(df_temp)
                    
                if f_105_admin:
                    df_temp = pd.read_excel(f_105_admin, skiprows=7)
                    df_temp = xuly_105_admin(df_temp)
                    erp_frames.append(df_temp)
                    
                if f_021_cu:
                    df_temp = pd.read_excel(f_021_cu, skiprows=7)
                    df_temp = xuly_021_cungung(df_temp)
                    erp_frames.append(df_temp)
                    
                if f_021_admin:
                    df_temp = pd.read_excel(f_021_admin, skiprows=7)
                    df_temp = xuly_021_admin(df_temp)
                    erp_frames.append(df_temp)
                
                # Chuẩn hóa tên cột ERP sau khi gộp
                if erp_frames:
                    df_erp = pd.concat(erp_frames, ignore_index=True)
                    df_erp.columns = df_erp.columns.astype(str).str.lower().str.strip()
                else:
                    df_erp = pd.DataFrame()

                # 3. HÀM ĐỐI SOÁT
                def clean_phieu(val):
                    return str(val).replace('.0', '').strip().upper()

                df_tc.columns = df_tc.columns.astype(str).str.lower().str.strip()
                col_phieu = next((c for c in df_tc.columns if "phiếu" in c or "chứng từ" in c), None)
                
                if col_phieu:
                    df_tc['số phiếu chuẩn'] = df_tc[col_phieu].apply(clean_phieu)
                    
                    col_net, col_gross = "khối lượng net (erp)", "khối lượng gross (erp)"
                    
                    if not df_erp.empty:
                        erp_col_phieu = next((c for c in df_erp.columns if "chứng từ" in c or "phiếu" in c), None)
                        erp_col_net = next((c for c in df_erp.columns if "net" in c), None)
                        erp_col_gross = next((c for c in df_erp.columns if "gross" in c), None)
                        
                        if erp_col_phieu:
                            df_erp['số phiếu chuẩn'] = df_erp[erp_col_phieu].apply(clean_phieu)
                            dict_net = dict(zip(df_erp['số phiếu chuẩn'], pd.to_numeric(df_erp[erp_col_net], errors='coerce'))) if erp_col_net else {}
                            dict_gross = dict(zip(df_erp['số phiếu chuẩn'], pd.to_numeric(df_erp[erp_col_gross], errors='coerce'))) if erp_col_gross else {}
                            
                            df_tc[col_net] = df_tc['số phiếu chuẩn'].map(dict_net).fillna(0)
                            df_tc[col_gross] = df_tc['số phiếu chuẩn'].map(dict_gross).fillna(0)
                        else:
                            df_tc[col_net], df_tc[col_gross] = 0, 0
                    else:
                        df_tc[col_net], df_tc[col_gross] = 0, 0

                    # Tính KL bàn cân
                    cong_cols = [c for c in df_tc.columns if "cổng" in c]
                    if cong_cols:
                        df_tc['kl bàn cân (max)'] = df_tc[cong_cols].apply(pd.to_numeric, errors='coerce').max(axis=1).fillna(0)
                    else:
                        col_kl = next((c for c in df_tc.columns if "khối lượng hàng hóa" in c or "khối lượng ra" in c), None)
                        df_tc['kl bàn cân (max)'] = pd.to_numeric(df_tc[col_kl], errors='coerce').fillna(0) if col_kl else 0

                    # Tính Chênh Lệch và Dung sai
                    df_tc['chênh lệch (kg)'] = df_tc[col_gross] - df_tc['kl bàn cân (max)']
                    df_tc['dung sai (%)'] = np.where(df_tc[col_gross] != 0, (df_tc['chênh lệch (kg)'] / df_tc[col_gross]).round(4), 0)
                    df_tc['cảnh báo dung sai'] = df_tc['dung sai (%)'].apply(lambda x: 'VUOT DUNG SAI' if abs(x) > 0.003 else 'OK')
                    df_tc['check net-gross erp'] = np.where(df_tc[col_gross] - df_tc[col_net] > 0, "Hợp lệ", "Lỗi Gross/Net")

                # Xuất file
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_tc.to_excel(writer, index=False, sheet_name="DoiSoat_TramCan_NM")
                    # Có thể xuất thêm 1 sheet chứa data ERP đã lọc để bạn dễ kiểm tra chéo
                    if not df_erp.empty:
                        df_erp.to_excel(writer, index=False, sheet_name="ERP_Da_Loc_Gop")

                st.success("✅ Đã lọc dữ liệu thành công và hoàn tất đối soát!")
                
                # Nút tải báo cáo
                st.download_button(
                    label="📥 Tải File Báo Cáo Đối Soạt (Đã áp dụng bộ lọc ERP)",
                    data=output.getvalue(),
                    file_name="Bao_Cao_Doi_Soat_Sau_Loc.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ Lỗi trong quá trình xử lý: {e}")
