import streamlit as st
from streamlit_option_menu import option_menu # Cần chạy: pip install streamlit-option-menu

# Mở rộng toàn màn hình
st.set_page_config(layout="wide", page_title="Hệ Thống Quản Lý")

# --- MENU ĐIỀU HƯỚNG BÊN TRÁI (SIDEBAR) ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2942/2942821.png", width=100) # Logo của bạn
    st.markdown("### LÊ CÔNG MINH")
    
    # Tạo menu đẹp mắt
    chuc_nang = option_menu(
        menu_title="DANH MỤC CHÍNH",
        options=["Trang Chủ", "Trạm Cân & ERP", "Kiểm Soát Nội Bộ", "Quản Lý Tồn Kho"],
        icons=["house", "truck", "shield-check", "box-seam"], # Lấy icon từ bootstrap icons
        menu_icon="cast",
        default_index=0,
        styles={
            "container": {"padding": "0!important", "background-color": "#fafafa"},
            "icon": {"color": "#0066cc", "font-size": "18px"}, 
            "nav-link": {"font-size": "16px", "text-align": "left", "margin":"0px"},
            "nav-link-selected": {"background-color": "#0066cc"},
        }
    )

# --- NỘI DUNG TƯƠNG ỨNG VỚI MENU ---
if chuc_nang == "Trang Chủ":
    st.title("📊 Tổng Quan Hệ Thống")
    st.write("Chào mừng bạn đến với hệ thống quản lý nội bộ.")
    
    # Chia 3 cột để hiển thị các chỉ số tóm tắt (KPIs) thay vì để các nút to
    col1, col2, col3 = st.columns(3)
    col1.metric("Phiếu Cân Hôm Nay", "124 Phiếu", "12 phiều chờ")
    col2.metric("Lỗi Cảnh Báo", "3 Lỗi", "-2 so với hôm qua", delta_color="inverse")
    col3.metric("Tồn Kho Thực Tế", "1,450 Tấn", "+50 Tấn")

elif chuc_nang == "Trạm Cân & ERP":
    st.header("🚛 Trạm Cân & Hàng Ra Cổng")
    st.info("Đối soát 4 bên (ERP 021, 105 và Trạm Cân NM, SMC), tính dung sai tự động.")
    # Chỗ này bạn bỏ code hiển thị bảng hoặc nút upload file của bạn vào

elif chuc_nang == "Kiểm Soát Nội Bộ":
    st.header("🔍 Kiểm Soát & Đối Chiếu HRC")
    st.warning("Hệ thống tự động phát hiện sai lệch và trùng mã.")
    # Code xử lý kiểm soát...

elif chuc_nang == "Quản Lý Tồn Kho":
    st.header("📦 Dashboard Tổng Hợp NXT")
    # Code biểu đồ, bảng tồn kho...
