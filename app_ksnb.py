import streamlit as st

st.set_page_config(layout="wide")

# Tạo menu bên trái bằng Radio button
with st.sidebar:
    st.markdown("### LÊ CÔNG MINH")
    chuc_nang = st.radio(
        "📍 DANH MỤC CHỨC NĂNG",
        ["🏠 Trang Chủ", "🚛 Trạm Cân & ERP", "🔍 Kiểm Soát Nội Bộ", "📦 Quản Lý Tồn Kho"]
    )

# Routing nội dung theo nút được chọn
if chuc_nang == "🏠 Trang Chủ":
    st.title("Tổng Quan Hệ Thống")
    st.write("Chào mừng bạn đến với công cụ kiểm soát tự động.")

elif chuc_nang == "🚛 Trạm Cân & ERP":
    st.title("Quản lý Trạm Cân")
    
elif chuc_nang == "🔍 Kiểm Soát Nội Bộ":
    st.title("Đối chiếu HRC")
    
elif chuc_nang == "📦 Quản Lý Tồn Kho":
    st.title("Báo cáo NXT")
