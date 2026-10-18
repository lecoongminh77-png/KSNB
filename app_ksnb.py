import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="KSNB - Kiểm Soát Tồn Kho", layout="wide", page_icon="🛡️")

st.title("🛡️ HỆ THỐNG KIỂM SOÁT TỒN KHO HÀNG NGÀY")
st.caption("Kiểm soát nội bộ (KSNB) • Tự động quét lỗi tỷ trọng & kiểm tra phẩm cấp cuộn")

# 1. KHU VỰC NẠP DỮ LIỆU
with st.sidebar:
    st.header("📂 Nạp Dữ Liệu")
    uploaded_file = st.file_uploader("Kéo thả file Excel tồn kho (.xlsx):", type=["xlsx"])
    st.markdown("---")

def load_data(file_source):
    headers = [
        'Kho', 'Ma_hang', 'Ten_hang', 'Ma_loai', 'Phan_loai', 'Phan_loai_SX',
        'Tieu_chuan', 'Grade', 'BMT', 'TCT', 'APT', 'Kho_mm', 'Do_ma',
        'Phu_son', 'Ma_mau', 'Ma_cuon', 'DVT', 'Gross', 'Net', 'Loi', 'So_met',
        'A', 'B', 'C', 'Ty_trong', 'Don_gia', 'Thanh_tien', 'Xuat_xu', 'Don_hang'
    ]
    df = pd.read_excel(file_source, sheet_name=0, skiprows=13, header=None)
    df = df.iloc[:, :len(headers)]
    df.columns = headers
    
    # Loại bỏ các dòng trống / dòng tổng kết
    df = df.dropna(subset=['Kho', 'Net']).copy()
    
    # Ép kiểu dữ liệu và chuẩn hóa
    df['Kho'] = df['Kho'].astype(str).str.strip()
    df['Phan_loai'] = df['Phan_loai'].fillna('Khác').astype(str).str.strip()
    df['Net_KG'] = pd.to_numeric(df['Net'], errors='coerce').fillna(0)
    df['Net_Tan'] = df['Net_KG'] / 1000
    df['So_met'] = pd.to_numeric(df['So_met'], errors='coerce').fillna(0)
    df['BMT'] = pd.to_numeric(df['BMT'], errors='coerce')
    df['Ty_trong'] = pd.to_numeric(df['Ty_trong'], errors='coerce').fillna(0)
    return df

# Nếu chưa nạp file thì nhắc nhở người dùng kéo file vào
if uploaded_file is None:
    st.info("👈 **Chào bạn! Hãy kéo thả file báo cáo tồn kho Excel (`.xlsx`) vào ô bên trái để bắt đầu phân tích.**")
    st.stop()

try:
    df_raw = load_data(uploaded_file)
    st.sidebar.success(f"Đã nạp file: {uploaded_file.name}")
except Exception as e:
    st.error(f"Lỗi khi đọc file: {e}")
    st.stop()

# 2. BỘ LỌC
with st.sidebar:
    st.header("🔍 Bộ Lọc Nhanh")
    kho_list = sorted([str(x) for x in df_raw['Kho'].unique()])
    selected_kho = st.multiselect("Lọc Vị trí kho:", options=kho_list, default=kho_list)
    
    pl_list = sorted([str(x) for x in df_raw['Phan_loai'].unique()])
    selected_pl = st.multiselect("Lọc Phân loại:", options=pl_list, default=pl_list)
    
    search_coil = st.text_input("Tìm kiếm Mã cuộn / Mã hàng:")

df = df_raw[(df_raw['Kho'].isin(selected_kho)) & (df_raw['Phan_loai'].isin(selected_pl))]
if search_coil:
    df = df[df['Ma_cuon'].astype(str).str.contains(search_coil, case=False, na=False) | 
            df['Ma_hang'].astype(str).str.contains(search_coil, case=False, na=False)]

# 3. THẺ CHỈ SỐ TỔNG QUAN
c1, c2, c3, c4 = st.columns(4)
c1.metric("Tổng Tồn Kho", f"{df['Net_Tan'].sum():,.2f} Tấn", f"{len(df):,} Cuộn/Lô")
c2.metric("Hàng Loại 1", f"{df[df['Phan_loai'] == '1']['Net_Tan'].sum():,.2f} Tấn", f"{(df['Phan_loai'] == '1').sum():,} Cuộn")
c3.metric("Chờ XL & Phế phẩm", f"{df[df['Phan_loai'].isin(['CXL', 'PP'])]['Net_Tan'].sum():,.2f} Tấn", f"{df['Phan_loai'].isin(['CXL', 'PP']).sum():,} Cuộn")

rot_xk = df[df['Phan_loai'].isin(['A1', 'A2'])]
c4.metric("Hàng Rớt XK (A1, A2)", f"{rot_xk['Net_Tan'].sum():,.2f} Tấn", f"{len(rot_xk):,} Cuộn")

st.markdown("---")

# 4. BIỂU ĐỒ BÁO CÁO
col_a, col_b = st.columns(2)
with col_a:
    kho_sum = df.groupby('Kho')['Net_Tan'].sum().reset_index()
    fig1 = px.bar(kho_sum, x='Kho', y='Net_Tan', text_auto='.2f', color='Kho', title="Khối lượng theo Kho (Tấn)")
    st.plotly_chart(fig1, use_container_width=True)

with col_b:
    pl_sum = df.groupby('Phan_loai')['Net_Tan'].sum().reset_index()
    fig2 = px.pie(pl_sum, names='Phan_loai', values='Net_Tan', hole=0.35, title="Cơ cấu theo Phân loại chất lượng")
    st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")

# 5. CẢNH BÁO KSNB
st.subheader("🚨 DANH SÁCH BẤT THƯỜNG CẦN KIỂM SOÁT")
t1, t2, t3 = st.tabs(["⚠️ Lệch Tỷ Trọng / Thiếu Mét", "⚠️ Kiểm Tra Quy Chuẩn A2", "📋 Bảng Dữ Liệu Chi Tiết"])

with t1:
    err_tytrong = df[(df['Net_KG'] > 0) & (df['So_met'] == 0)][['Kho', 'Ma_cuon', 'Ma_hang', 'Ten_hang', 'Net_KG', 'So_met', 'Ty_trong']]
    st.write(f"Tìm thấy **{len(err_tytrong)}** dòng có khối lượng Net nhưng thiếu số mét (tỷ trọng = 0):")
    st.dataframe(err_tytrong, use_container_width=True)

with t2:
    a2_all = df[df['Phan_loai'] == 'A2']
    st.write(f"Tổng số cuộn A2: **{len(a2_all)}** cuộn (Quy định KSNB: BMT phải ≥ 0.57mm).")
    a2_err = a2_all[a2_all['BMT'] < 0.57]
    if len(a2_err) == 0:
        st.success("✅ Toàn bộ cuộn hàng A2 đều đúng quy định độ dày!")
    else:
        st.error(f"⚠️ Có {len(a2_err)} cuộn A2 vi phạm tiêu chuẩn độ dày:")
        st.dataframe(a2_err[['Kho', 'Ma_cuon', 'Ten_hang', 'BMT', 'Net_KG']], use_container_width=True)

with t3:
    st.dataframe(df[['Kho', 'Ma_cuon', 'Ma_hang', 'Ten_hang', 'Phan_loai', 'BMT', 'Net_KG', 'So_met', 'Ty_trong']], use_container_width=True)
