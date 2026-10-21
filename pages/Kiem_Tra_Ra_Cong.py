import sys
import os
import requests

# 1. Ép Python nhận diện thư mục gốc (nơi chứa ai_chat.py)
root_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root_path not in sys.path:
    sys.path.insert(0, root_path)

# 2. Khai báo các thư viện khác
import io
import pandas as pd
import numpy as np
import streamlit as st
from streamlit_lottie import st_lottie

# 3. Import module nội bộ
from ai_chat import hien_thi_tro_ly_ai

# Hàm tải hình động
def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()


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
</style>
""", unsafe_allow_html=True)

# Thanh điều hướng (Chỉ giữ duy nhất nút Trang chủ)
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


# 4 Ô NẠP FILE (2 Cột x 2 Hàng)
st.markdown("### 📥 Tải Dữ Liệu Đầu Vào (4 File Nguồn)")
col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    f_021 = st.file_uploader("1. File ERP (Mã 021)", type=["xlsx", "xls", "xlsb"])
    f_tramcan_nm = st.file_uploader("3. File Trạm Cân (Nhà Máy)", type=["xlsx", "xls", "xlsb"])
    st.markdown('</div>', unsafe_allow_html=True)
with col2:
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    f_105 = st.file_uploader("2. File ERP (Mã 105)", type=["xlsx", "xls", "xlsb"])
    f_tramcan_smc = st.file_uploader("4. File Trạm Cân (SMC)", type=["xlsx", "xls", "xlsb"])
    st.markdown('</div>', unsafe_allow_html=True)

# Xử lý tự động
if st.button("🚀 Bắt đầu đối soát & Xuất Báo Cáo", use_container_width=True):
    if not (f_021 or f_105 or f_tramcan_nm or f_tramcan_smc):
        st.warning("⚠️ Vui lòng nạp ít nhất 1 file Trạm Cân và 1 file ERP để đối soát!")
    else:
        with st.spinner("Đang xử lý kết hợp dữ liệu và tính toán Chênh lệch, Dung sai..."):
            try:
                # Đọc File Trạm Cân NM
                df_nm = pd.read_excel(f_tramcan_nm) if f_tramcan_nm else pd.DataFrame()
                
                # Đọc File Trạm Cân SMC
                df_smc = pd.read_excel(f_tramcan_smc) if f_tramcan_smc else pd.DataFrame()
                
                # Đọc File ERP (Gộp 021 và 105)
                df_erp_021 = pd.read_excel(f_021, skiprows=7) if f_021 else pd.DataFrame()
                df_erp_105 = pd.read_excel(f_105, skiprows=7) if f_105 else pd.DataFrame()
                
                # Chuyển đổi tên cột chuẩn cho ERP
                erp_frames = []
                for df_e in [df_erp_021, df_erp_105]:
                    if not df_e.empty:
                        df_e.columns = df_e.columns.astype(str).str.lower().str.strip()
                        erp_frames.append(df_e)
                
                df_erp = pd.concat(erp_frames, ignore_index=True) if erp_frames else pd.DataFrame()

                # CÁC HÀM XỬ LÝ & ĐỐI SOÁT DỮ LIỆU
                def clean_phieu(val):
                    return str(val).replace('.0', '').strip().upper()

                def tinh_toan_tram_can(df_tramcan, is_smc=False):
                    if df_tramcan.empty:
                        return pd.DataFrame()
                    
                    df_tc = df_tramcan.copy()
                    df_tc.columns = df_tc.columns.astype(str).str.lower().str.strip()
                    
                    # Tìm cột Số Phiếu / Chứng Từ
                    col_phieu = next((c for c in df_tc.columns if "phiếu" in c or "chứng từ" in c), None)
                    if not col_phieu:
                        return df_tc
                        
                    df_tc['số phiếu chuẩn'] = df_tc[col_phieu].apply(clean_phieu)
                    
                    # Cột ERP Gross/Net
                    col_net, col_gross = "khối lượng net", "khối lượng gross"
                    if not df_erp.empty:
                        # VLOOKUP lấy Net, Gross từ ERP dựa theo phiếu
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
                            df_tc[col_net] = 0
                            df_tc[col_gross] = 0
                    else:
                        df_tc[col_net] = 0
                        df_tc[col_gross] = 0

                    # Tính KL bàn cân (ví dụ lấy Max của các cột có chữ "cổng")
                    cong_cols = [c for c in df_tc.columns if "cổng" in c]
                    if cong_cols:
                        df_tc['kl bàn cân lớn nhất'] = df_tc[cong_cols].apply(pd.to_numeric, errors='coerce').max(axis=1).fillna(0)
                    else:
                        # Fallback về khối lượng hàng hóa trạm cân
                        col_kl = next((c for c in df_tc.columns if "khối lượng hàng hóa" in c or "khối lượng ra" in c), None)
                        df_tc['kl bàn cân lớn nhất'] = pd.to_numeric(df_tc[col_kl], errors='coerce').fillna(0) if col_kl else 0

                    # Tính Chênh Lệch và Dung sai
                    df_tc['chênh lệch'] = df_tc[col_gross] - df_tc['kl bàn cân lớn nhất']
                    df_tc['dung sai'] = np.where(df_tc[col_gross] != 0, (df_tc['chênh lệch'] / df_tc[col_gross]).round(4), 0)
                    df_tc['cảnh báo dung sai'] = df_tc['dung sai'].apply(lambda x: 'VUOT DUNG SAI' if abs(x) > 0.003 else 'OK')

                    # Check Net-Gross
                    df_tc['check net - gross'] = np.where(df_tc[col_gross] - df_tc[col_net] > 0, "Hợp lệ", "Lỗi Gross/Net")

                    return df_tc

                # Chạy đối soát cho 2 sheet
                df_ketqua_nm = tinh_toan_tram_can(df_nm, is_smc=False)
                df_ketqua_smc = tinh_toan_tram_can(df_smc, is_smc=True)

                # Xuất ra file Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    if not df_ketqua_nm.empty:
                        df_ketqua_nm.to_excel(writer, index=False, sheet_name="KSKTDulieucanNM")
                    if not df_ketqua_smc.empty:
                        df_ketqua_smc.to_excel(writer, index=False, sheet_name="KSKTDulieucanSMC")

                st.success("✅ Đã xử lý đối soát xong Dữ liệu Trạm Cân và ERP!")
                
                # Layout tải file
                col_dl1, col_dl2 = st.columns([1, 1])
                with col_dl1:
                    st.download_button(
                        label="📥 Tải File Báo Cáo KSKT (Trạm cân NM & SMC)",
                        data=output.getvalue(),
                        file_name="Bao_Cao_Doi_Soat_Ra_Cong.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Lỗi khi xử lý dữ liệu: {e}")
hien_thi_tro_ly_ai()
