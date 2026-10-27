import streamlit as st

st.set_page_config(layout="wide")
st.title("Hệ Thống Quản Lý Nội Bộ")

# Tạo 3 thẻ điều hướng ngang
tab1, tab2, tab3 = st.tabs(["🚛 Trạm Cân & ERP", "🔍 Kiểm Soát Nội Bộ", "📦 Quản Lý Tồn Kho"])

with tab1:
    st.header("Đối soát dữ liệu 4 bên")
    # Bỏ code phần đối soát cổng ở đây

with tab2:
    st.header("Kiểm Tra Mẫu & Đối Chiếu HRC")
    # Bỏ code kiểm soát HRC ở đây

with tab3:
    st.header("Tổng Hợp Nhập Xuất Tồn")
    # Bỏ code dashboard NXT ở đây
