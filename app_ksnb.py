import streamlit as st

# 1. CẤU HÌNH TRANG (Phải để trên cùng)
st.set_page_config(page_title="Hệ Thống Kiểm Soát", page_icon="🧊", layout="wide")

# 2. CSS HACK GIAO DIỆN (Làm nổi thẻ, bo góc, bóng mờ)
st.markdown("""
<style>
    /* Đổi nền trang thành màu xám nhạt sang trọng */
    .stApp { background-color: #F4F7FC; }
    
    /* Style cho các thẻ chỉ số (Metric Cards) */
    div[data-testid="metric-container"] {
        background-color: white;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        transition: transform 0.2s ease-in-out;
    }
    div[data-testid="metric-container"]:hover {
        transform: translateY(-5px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    
    /* Style cho các khu vực chứa nội dung (Container) */
    div[data-testid="stVerticalBlock"] > div {
        background-color: transparent;
    }
    
    /* Ẩn bớt các menu mặc định của Streamlit cho chuyên nghiệp */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# 3. SIDEBAR (Menu bên trái)
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2942/2942821.png", width=80)
    st.markdown("## 👨‍💻 LÊ CÔNG MINH")
    st.markdown("---")
    # Menu dùng Radio button mặc định nhưng hiển thị gọn gàng
    menu = st.radio(
        "ĐIỀU HƯỚNG",
        ["📊 Tổng Quan (Dashboard)", "🚛 Hàng Ra Cổng (Trạm Cân)", "🔍 Đối Chiếu HRC", "📦 Tổng Hợp Tồn Kho"]
    )
    st.markdown("---")
    st.caption("Cập nhật lần cuối: Hôm nay")

# 4. NỘI DUNG CHÍNH (Thay đổi theo Menu)
if menu == "📊 Tổng Quan (Dashboard)":
    st.title("Hiệu Suất Hoạt Động Hôm Nay")
    st.markdown("Theo dõi tự động dữ liệu Trạm Cân và Tồn Kho.")
    
    st.write("") # Tạo khoảng trắng
    
    # Khu vực KPI (4 thẻ)
    col1, col2, col3, col4 = st.columns(4)
    col1.metric(label="🚚 Số Chuyến Xe", value="124", delta="+12 hôm nay")
    col2.metric(label="⚠️ Sai Lệch Trạm Cân", value="3", delta="-2 chuyến", delta_color="inverse")
    col3.metric(label="📦 Tồn Kho HRC", value="4,250 T", delta="-150 T")
    col4.metric(label="✔️ Mẫu Đạt Chuẩn", value="98.5%", delta="0.5%")
    
    st.markdown("---")
    
    # Khu vực Bảng và Biểu đồ
    c1, c2 = st.columns((6, 4))
    with c1:
        st.subheader("Bảng Cảnh Báo Gần Nhất")
        st.info("Chỗ này bạn dùng st.dataframe() để hiển thị bảng dữ liệu lỗi từ file Excel")
    with c2:
        st.subheader("Tiến độ Nhập/Xuất")
        st.success("Chỗ này bạn dùng st.bar_chart() để vẽ biểu đồ")

elif menu == "🚛 Hàng Ra Cổng (Trạm Cân)":
    st.title("Phân Hệ Trạm Cân & ERP")
    st.write("Tải file dữ liệu để đối soát 4 bên (ERP 021, 105 và Trạm Cân NM, SMC)")
    # Thêm code uplaod file của bạn vào đây...
    
else:
    st.title(menu)
    st.write("Đang phát triển tính năng này...")
