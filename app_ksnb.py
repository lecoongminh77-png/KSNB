import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Dashboard KSNB - Tồn Kho Thép", layout="wide")

@st.cache_data
def load_data():
    headers = [
        'Kho', 'Ma_hang', 'Ten_hang', 'Ma_loai', 'Phan_loai', 'Phan_loai_SX',
        'Tieu_chuan', 'Grade', 'BMT', 'TCT', 'APT', 'Kho_mm', 'Do_ma',
        'Phu_son', 'Ma_mau', 'Ma_cuon', 'DVT', 'Gross', 'Net', 'Loi', 'So_met',
        'A', 'B', 'C', 'Ty_trong', 'Don_gia', 'Thanh_tien', 'Xuat_xu', 'Don_hang'
    ]
    df = pd.read_excel('15-10-2026.xlsx', sheet_name='Sheet1', skiprows=13, header=None)
    df = df.iloc[:, :len(headers)]
    df.columns = headers
    
    df['Net_KG'] = pd.to_numeric(df['Net'], errors='coerce').fillna(0)
    df['Net_Tan'] = df['Net_KG'] / 1000
    df['So_met'] = pd.to_numeric(df['So_met'], errors='coerce').fillna(0)
    df['BMT'] = pd.to_numeric(df['BMT'], errors='coerce')
    df['Ty_trong'] = pd.to_numeric(df['Ty_trong'], errors='coerce').fillna(0)
    return df

df = load_data()

st.title("🛡️ BẢNG ĐIỀU HÀNH KIỂM SOÁT TỒN KHO")
st.caption("Dữ liệu báo cáo ngày: 15/10/2026 • Tự động quét từ file hệ thống")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Tổng tồn kho", f"{df['Net_Tan'].sum():,.2f} Tấn", f"{len(df):,} Cuộn/Lô")
c2.metric("Hàng Loại 1 (Chuẩn)", f"{df[df['Phan_loai'] == '1']['Net_Tan'].sum():,.2f} Tấn", f"{(df['Phan_loai'] == '1').sum():,} Cuộn")
c3.metric("Chờ XL & Phế phẩm", f"{df[df['Phan_loai'].isin(['CXL', 'PP'])]['Net_Tan'].sum():,.2f} Tấn", f"{df['Phan_loai'].isin(['CXL', 'PP']).sum():,} Cuộn")
c4.metric("Hàng rớt XK (A1, A2)", f"{df[df['Phan_loai'].isin(['A1', 'A2'])]['Net_Tan'].sum():,.2f} Tấn", f"{df['Phan_loai'].isin(['A1', 'A2']).sum():,} Cuộn")

st.divider()

col_left, col_right = st.columns(2)
with col_left:
    st.subheader("📊 Tồn kho theo Vị trí kho")
    kho_summary = df.groupby('Kho')['Net_Tan'].sum().reset_index()
    fig_kho = px.bar(kho_summary, x='Kho', y='Net_Tan', text_auto='.2f', color='Kho', title="Khối lượng theo kho (Tấn)")
    st.plotly_chart(fig_kho, use_container_width=True)

with col_right:
    st.subheader("📑 Cơ cấu theo Phân loại chất lượng")
    pl_summary = df.groupby('Phan_loai')['Net_Tan'].sum().reset_index()
    fig_pl = px.pie(pl_summary, names='Phan_loai', values='Net_Tan', hole=0.4, title="Tỷ lệ phân loại phẩm cấp")
    st.plotly_chart(fig_pl, use_container_width=True)

st.divider()

st.subheader("🚨 CẢNH BÁO BẤT THƯỜNG CẦN KIỂM TRA GẤP")
tab1, tab2 = st.tabs(["Lệch Tỷ Trọng / Thiếu Dữ Liệu", "Kiểm Tra Hàng A2"])

with tab1:
    anomalies = df[(df['Net_KG'] > 0) & (df['So_met'] == 0)][['Kho', 'Ma_hang', 'Ten_hang', 'Ma_cuon', 'Net_KG', 'So_met', 'Ty_trong']]
    st.write(f"Phát hiện **{len(anomalies)}** dòng có khối lượng Net nhưng số mét = 0:")
    st.dataframe(anomalies, use_container_width=True)

with tab2:
    a2_check = df[df['Phan_loai'] == 'A2'][['Kho', 'Ma_cuon', 'Ten_hang', 'BMT', 'Net_KG']]
    st.write(f"Tổng số cuộn gắn nhãn A2: **{len(a2_check)}** cuộn. Quy định: Chỉ gắn cho thép dày (BMT ≥ 0.57mm).")
    invalid_a2 = a2_check[a2_check['BMT'] < 0.57]
    if len(invalid_a2) == 0:
        st.success("✅ Toàn bộ cuộn hàng A2 đều thỏa mãn tiêu chuẩn độ dày (BMT ≥ 0.57mm)!")
    else:
        st.error(f"⚠️ Phát hiện {len(invalid_a2)} cuộn vi phạm tiêu chuẩn độ dày:")
        st.dataframe(invalid_a2, use_container_width=True)
